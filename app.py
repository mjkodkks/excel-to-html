import base64
from datetime import datetime
import glob
import os
import re
import subprocess
import minify_html
import pandas as pd
import shutil
from bs4 import BeautifulSoup
import csv
from pathlib import Path
from urllib.parse import unquote
from simple_salesforce import Salesforce
from dotenv import load_dotenv

load_dotenv()

sf = Salesforce(
    username=os.getenv('SALESFORCE_USERNAME'),
    password=os.getenv('SALESFORCE_PASSWORD'),
    security_token=os.getenv('SALESFORCE_SECURITY_TOKEN'), 
    domain=os.getenv('SALESFORCE_DOMAIN')
)

base_sfc_url = os.getenv('SALESFORCE_FILE_URL', 'https://your_instance.salesforce.com/sfc/servlet.shepherd/document/download/')

try:
    sf_identity = sf.User.describe()
    print("✅ Connected to Salesforce successfully.")
except Exception as e:
    print("❌ Failed to connect to Salesforce:", e)

input_folder = Path("input_files")
output_folder = Path("output_html")
output_result_folder = Path("output_result")
field_size_limit = 400000

def input_to_html():
    print("## Start INPUT TO HTML ##")
    if output_folder.exists():
        shutil.rmtree(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)

    input_files = (
    list(input_folder.glob("*.xlsx")) +
    list(input_folder.glob("*.xls")) +
    list(input_folder.glob("*.docx")) +
    list(input_folder.glob("*.doc"))
)

    if not input_files:
        print("No compatible files found.")
        return

    for index, file in enumerate(input_files):
        print(f"Processing file: {file.name}")
        base_name = file.stem
        file_output_folder = output_folder / f"{index}_{base_name}"
        file_output_folder.mkdir(parents=True, exist_ok=True)

        try:
            subprocess.run([
                "soffice", "--headless", "--convert-to", "html",
                "--outdir", str(file_output_folder.resolve()), str(file.resolve())
            ], check=True)

        except subprocess.CalledProcessError as e:
            print(f"Error converting {file.name}: {e}")

    print("Conversion completed!")

def fetch_existing_file_titles():
    query = """
        SELECT ContentDocument.Title, ContentDocumentId 
        FROM ContentVersion 
        WHERE FileExtension IN ('png', 'jpg', 'jpeg', 'gif', 'bmp')
    """
    results = sf.query_all(query)
    title_to_id = {}
    for record in results['records']:
        title = record['ContentDocument']['Title'].lower().strip()
        if title not in title_to_id:
            title_to_id[title] = record['ContentDocumentId']
    return title_to_id

def rename_images_and_refs():
    print("## Renaming images sequentially and updating HTML src ##")
    rename_map = {}
    for folder in output_folder.iterdir():
        if not folder.is_dir():
            continue

        html_files = list(folder.glob("*.html"))
        if not html_files:
            continue

        html_file = html_files[0]
        with open(html_file, "r", encoding="utf-8") as f:
            soup = BeautifulSoup(f.read(), "html.parser")

        images = soup.find_all("img")

        for idx, img in enumerate(images, start=1):
            src = img.get("src", "")
            src = unquote(src)
            if src and Path(src).suffix.lower() in [".png", ".jpg", ".jpeg", ".gif"]:
                original_path = folder / Path(src).name
                if original_path.exists():
                    ext = Path(src).suffix.lower()
                    new_name = f"{folder.name.split('_', 1)[1]}_html_{idx}{ext}"
                    new_path = folder / new_name
                    print(f"🔄 Renaming {original_path.name} → {new_name}")
                    original_path.rename(new_path)
                    rename_map[new_name.lower().strip()] = str(new_path)
                    img["src"] = new_name

        with open(html_file, "w", encoding="utf-8") as f:
            f.write(str(soup))

    return rename_map

def upload_missing_images(image_map, existing_titles):
    print("## Uploading only unmatched images ##")
    image_lookup = {}
    for title, local_path in image_map.items():
        if title in existing_titles:
            print(f"⚠️ Skipping upload, already exists: {title}")
            image_lookup[title] = existing_titles[title]
        else:
            with open(local_path, "rb") as f:
                data = f.read()
            encoded_file = base64.b64encode(data).decode("utf-8")

            try:
                response = sf.ContentVersion.create({
                    "Title": Path(local_path).name,
                    "PathOnClient": Path(local_path).name,
                    "VersionData": encoded_file
                })
                version_id = response.get("id")
                query = f"SELECT ContentDocumentId FROM ContentVersion WHERE Id = '{version_id}'"
                result = sf.query(query)
                doc_id = result['records'][0]['ContentDocumentId']
                image_lookup[title] = doc_id
                print(f"🖼️ Uploaded image: {local_path}")
            except Exception as e:
                print(f"❌ Failed to upload image: {local_path} — {e}")
    return image_lookup

def update_html_images(image_lookup):
    for folder in output_folder.iterdir():
        if not folder.is_dir():
            continue
        html_files = list(folder.glob("*.html"))
        if not html_files:
            continue

        html_file = html_files[0]
        with open(html_file, "r", encoding="utf-8") as f:
            soup = BeautifulSoup(f.read(), "html.parser")

        for img in soup.find_all("img"):
            src = img.get("src", "")
            src = unquote(src)
            filename = Path(src).name.lower().strip()
            matched_doc_id = image_lookup.get(filename)
            if matched_doc_id:
                new_src = f"{base_sfc_url}{matched_doc_id}"
                print(f"🔄 Replaced src '{src}' with '{new_src}'")

                # Create a wrapper <span> around the <img>
                wrapper = soup.new_tag("span", style="display: inline-flex; align-items: center; gap: 0.5em;")
                img["src"] = new_src
                img["style"] = "vertical-align: middle; max-height: 2000px;"
                img.replace_with(wrapper)
                wrapper.append(img)

                # Check if there's a NavigableString/text right after the image
                next_sibling = wrapper.next_sibling
                if next_sibling and isinstance(next_sibling, str):
                    text_span = soup.new_tag("span")
                    text_span.string = next_sibling.strip()
                    wrapper.append(text_span)
                    next_sibling.extract()
            else:
                print(f"⚠️ No match found for image: {filename}")

        with open(html_file, "w", encoding="utf-8") as f:
            f.write(str(soup))

def read_all_html_files():
    print("## Read All HTML Files ##")
    if not output_folder.is_dir():
        print("output_html folder not found.")
        return
    html_folders = [folder for folder in output_folder.iterdir() if folder.is_dir()]
    all_contents = {}

    for folder in html_folders:
        folder_index = os.path.basename(folder).split("_")[0]
        print(f"Processing folder: {folder}")

        files = list(folder.glob("*.html"))
        contents_list = []
        for file in files:
            print(f"Reading file: {file}")
            try:
                filename = Path(file).stem
                with open(file, 'r', encoding='utf-8') as f:
                    contentTxt = f.read()
                    soup = BeautifulSoup(contentTxt, "html.parser")

                    tdBgcolor = soup.body.select("table td[bgcolor]")
                    tdAlign = soup.body.select("table td[align]")
                    dataSheetsValue = soup.find_all(attrs={"data-sheets-value": True})
                    fontAllTag = soup.body.find_all("font")

                    font_sizes = {
                        "1": "x-small",
                        "2": "small",
                        "3": "medium",
                        "4": "large",
                        "5": "x-large",
                        "6": "xx-large",
                        "7": "-webkit-xxx-large"
                    }

                    for font in fontAllTag:
                        style_parts = []
                        if font.has_attr("size"):
                            size = font_sizes.get(font["size"], "medium")
                            style_parts.append(f"font-size: {size};")
                        if font.has_attr("color"):
                            style_parts.append(f"color: {font['color']};")
                        if font.has_attr("face"):
                            style_parts.append(f"font-family: '{font['face']}', Arial, sans-serif;")
                        span_tag = soup.new_tag("span", style=" ".join(style_parts))
                        span_tag.string = font.getText()
                        font.replace_with(span_tag)

                    for td in tdBgcolor:
                        bgcolor = td["bgcolor"]
                        existing_style = td.get("style", "")
                        td["style"] = f"background-color: {bgcolor}; {existing_style}".strip()

                    for td in tdAlign:
                        align = td["align"]
                        if align == 'middle':
                            align = 'center'
                        existing_style = td.get("style", "")
                        td["style"] = f"text-align: {align}; {existing_style}".strip()

                    for tag in dataSheetsValue:
                        del tag["data-sheets-value"]

                body_content = soup.body.encode_contents().decode("utf-8")
                body_content = minify_html.minify(body_content, keep_closing_tags=True, minify_js=False, minify_css=False, do_not_minify_doctype=True)

                contents_list.append({
                    "title": filename,
                    "content": body_content,
                    "parent_title": folder,
                    "is_field_exceed": len(body_content) >= field_size_limit
                })

            except Exception as e:
                print(f"Error reading {file}: {e}")

        all_contents[folder_index] = contents_list

    return all_contents

def bulk_import_html_to_salesforce(all_contents):
    print("## Bulk Uploading to Salesforce Knowledge ##")

    for folder_index, contents in all_contents.items():
        for index, content in enumerate(contents):
            if content["is_field_exceed"]:
                print(f"⚠️ Skipping '{content['title']}' — field size too large.")
                continue

            try:
                response = sf.Knowledge__kav.create({
                    "Title": content["title"],
                    "UrlName": f"url-{datetime.now().strftime('%Y%m%d%H%M%S%f')}",
                    "Answer__c": content["content"],
                    "RecordTypeId": "012N00000036GnwIAE",
                    "Language": "en_US"
                })
                print(f"✅ Created article: {content['title']} (Id: {response['id']})")
            except Exception as e:
                print(f"❌ Failed to create article: {content['title']}", e)

def remove_all_black_color_tags():
    print("## Removing all #000000 styles and attributes from HTML ##")

    if not output_folder.exists():
        print("No output_html folder found.")
        return

    html_folders = [folder for folder in output_folder.iterdir() if folder.is_dir()]
    
    for folder in html_folders:
        html_files = list(folder.glob("*.html"))
        for html_file in html_files:
            try:
                with open(html_file, "r", encoding="utf-8") as f:
                    content = f.read()
                content = re.sub(r'(style\s*=\s*["\'][^"\']*)#000000\s*;?', r'\1', content, flags=re.IGNORECASE)
                content = re.sub(r'\s*color\s*=\s*["\']#000000["\']', '', content, flags=re.IGNORECASE)
                content = content.replace("#000000", "")

                with open(html_file, "w", encoding="utf-8") as f:
                    f.write(content)

                print(f"🧹 Cleaned all #000000 from: {html_file}")
            except Exception as e:
                print(f"Error cleaning {html_file}: {e}")

# Pipeline
input_to_html()
image_map = rename_images_and_refs()
existing_titles = fetch_existing_file_titles()
image_lookup = upload_missing_images(image_map, existing_titles)
update_html_images(image_lookup)
remove_all_black_color_tags()
html_data = read_all_html_files()
bulk_import_html_to_salesforce(html_data)

def is_excel(filename: str) -> bool:
    pattern = r'.*\.(xls|xlsx)$'
    return bool(re.match(pattern, filename, re.IGNORECASE))

def is_csv(filename: str) -> bool:
    pattern = r'.*\.(csv)$'
    return bool(re.match(pattern, filename, re.IGNORECASE)) 

# %%
