# %%
from datetime import datetime
import glob
import os
import re
import subprocess
import time
import minify_html
import pandas as pd
import shutil
from bs4 import BeautifulSoup
import csv

input_folder = "excel_files"  # Folder containing Excel files
output_folder = "output_html"  # Folder to save HTML outputs

def excel_to_HTML():
    print("## Start EXCEL TO HTML ##")
    # Remove all items in output_html
    shutil.rmtree(output_folder, ignore_errors=True)

    # Ensure output directory exists
    os.makedirs(output_folder, exist_ok=True)

    # Get all Excel files in the folder
    excel_files = [f for f in os.listdir(input_folder) if f.endswith((".xlsx", ".xls"))]

    # Convert each Excel file separately
    for index, file in enumerate(excel_files):
        print(f"Processing file: {file}")
        base_name, _ = os.path.splitext(file)  # Get file name without extension
        file_output_folder = os.path.join(output_folder, f"{index}_{base_name}")

        # Create a separate folder for each file
        os.makedirs(file_output_folder, exist_ok=True)

        # Full input file path
        input_path = os.path.join(input_folder, file)

        # Convert Excel to HTML inside the specific folder
        subprocess.run([
            "soffice",
            "--headless",
            "--convert-to", "html",
            "--outdir", file_output_folder,
            input_path
        ], check=True)

    print("Conversion completed!")

def read_all_html_files():
    print("## Read All HTML Files ##")
    os.makedirs(output_folder, exist_ok=True)
    html_folders = glob.glob(output_folder + '/*')  # Get all subfolders
    all_contents = {}

    for folder in html_folders:
        folder_index = os.path.basename(folder).split("_")[0]  # Extract folder index
        print(f"Processing folder: {folder}")

        files = glob.glob(folder + '/*.html')  # Get all HTML files in the folder
        contents_list = []
        for file in files:
            print(f"Reading file: {file}")
            try:
                filename = os.path.basename(file)
                with open(file, 'r', encoding='utf-8') as f:
                    contentTxt = f.read()
                    soup = BeautifulSoup(contentTxt, "html.parser")
                    
                    # Modify styles
                    tdBgcolor = soup.body.select("table td[bgcolor]")
                    tdAlign = soup.body.select("table td[align]")

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

                    body_content = soup.body.decode_contents()

                    # Remove unnecessary attributes
                    patternRemoveUnusedAttr = r'\s*data-sheets-value=\'\{.*?\}\''
                    patternRemoveTag = r'<br.*?\/>|<img.*?>'
                    body_content = re.sub(patternRemoveUnusedAttr, '', body_content)
                    body_content = re.sub(patternRemoveTag, '', body_content)
                    body_content = minify_html.minify(body_content, minify_js=False, minify_css=False, remove_processing_instructions=True, keep_spaces_between_attributes=True)

                    contents_list.append({
                        "title": filename,
                        "content": body_content
                    })
            except Exception as e:
                print(f"Error reading {file}: {e}")

        # Store content list by folder index
        all_contents[folder_index] = contents_list

    return all_contents

def create_output_files(all_contents):
    print("## Creating Output CSV Files ##")
    
    for folder_index, contents in all_contents.items():
        output_filename = f"output-{folder_index}.csv"
        print(f"Creating: {output_filename}")

        if not contents:
            print(f"No content found for folder {folder_index}, skipping.")
            continue

        dataCreateCSV = {
            "Knowledge__kav": [],
            "Id": [],
            "RecordTypeId": [],
            "Title": [],
            "UrlName": [],
            "Summary": [],
            "Answer": [],
            "Categorie__c": [],
            "Category__c": []
        }

        for index, content in enumerate(contents):
            ts = datetime.now().strftime("%Y%m%d%H%M%S%f")[:17]
            urlMock = f"URL-{ts}{index}"
            dataCreateCSV["Knowledge__kav"].append(index)
            dataCreateCSV["Id"].append("test")
            dataCreateCSV["RecordTypeId"].append("012N00000036GnwIAE")
            dataCreateCSV["Title"].append(content["title"] + "_(test-html-import)")
            dataCreateCSV["UrlName"].append(urlMock)
            dataCreateCSV["Summary"].append(content["title"])
            dataCreateCSV["Answer"].append(content["content"])
            dataCreateCSV["Categorie__c"].append("")
            dataCreateCSV["Category__c"].append("Knowledge Material")

        df = pd.DataFrame(dataCreateCSV)
        df.to_csv(output_filename, index=False, sep=",", quoting=csv.QUOTE_NONNUMERIC, quotechar='"', escapechar="\\")

    print("All output CSV files created successfully!")


def is_excel_or_csv(filename: str) -> bool:
    pattern = r'.*\.(csv|xls|xlsx)$'
    return bool(re.match(pattern, filename, re.IGNORECASE))

def is_excel(filename: str) -> bool:
    pattern = r'.*\.(xls|xlsx)$'
    return bool(re.match(pattern, filename, re.IGNORECASE))

def is_csv(filename: str) -> bool:
    pattern = r'.*\.(csv)$'
    return bool(re.match(pattern, filename, re.IGNORECASE))

excel_to_HTML()
html_data = read_all_html_files()
create_output_files(html_data)