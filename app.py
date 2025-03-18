# %%
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

input_folder = "excel_files"  # Folder containing Excel files
output_folder = "output_html"  # Folder to save HTML outputs
output_result_folder = "output_result"  # Folder to output result outputs
field_size_limit = 400000

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
                filename = Path(file).stem
                with open(file, 'r', encoding='utf-8') as f:
                    contentTxt = f.read()
                    soup = BeautifulSoup(contentTxt, "html.parser")
                    tables = soup.body.select("table")
                    name_of_sheet = soup.body.find_all("a", attrs={"name": True})
                    count_table = len(tables)
                    is_table_more_than_one = count_table > 1

                    if is_table_more_than_one:
                        print(count_table)
                    
                    # Modify styles
                    tdBgcolor = soup.body.select("table td[bgcolor]")
                    tdAlign = soup.body.select("table td[align]")
                    fontTag = soup.body.select("td font")
                    dataSheetsValue = soup.find_all(attrs={"data-sheets-value": True})
                    brAndImage = soup.find_all(["br", "img"])

                    for font in fontTag:
                        color = font.get("color", "")
                        if color != "#000000":
                            continue
                        text = font.get_text()
                        font.replace_with(text)

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
                    
                    # Remove the data-sheets-value attribute from all elements
                    for tag in dataSheetsValue:
                        del tag["data-sheets-value"]

                    # Remove <br> and <img> tags
                    for tag in brAndImage:
                        tag.decompose()

                    for index, table in enumerate(tables):
                        body_content = table.prettify()
                        if is_table_more_than_one:
                            title = f'{filename}_{name_of_sheet[index].getText()}'
                        else:
                            title = filename

                        content_length = len(body_content)
                        print(f"## content length : {content_length}")
                        is_field_exceed = content_length >= field_size_limit
                        if is_field_exceed:
                            print('| field size limit exceed!', end=" ")

                        # Remove unnecessary attributes
                        # patternRemoveUnusedAttr = r'\s*data-sheets-value=\'\{.*?\}\''
                        # patternRemoveTag = r'<br.*?\/>|<img.*?>'
                        # body_content = re.sub(patternRemoveUnusedAttr, '', body_content)
                        # body_content = re.sub(patternRemoveTag, '', body_content)
                        body_content = minify_html.minify(body_content, keep_closing_tags=True, minify_js=False, minify_css=False, remove_processing_instructions=True, keep_spaces_between_attributes=True)

                        contents_list.append({
                            "title": title,
                            "content": body_content,
                            "parent_title": folder,
                            "is_field_exceed": is_field_exceed
                        })
            except Exception as e:
                print(f"Error reading {file}: {e}")

        # Store content list by folder index
        all_contents[folder_index] = contents_list

    return all_contents

def create_output_files(all_contents):
    print("## Creating Output CSV Files ##")

    # Remove all items in output_html
    shutil.rmtree(output_result_folder, ignore_errors=True)

    # Ensure output directory exists
    os.makedirs(output_result_folder, exist_ok=True)
    
    for folder_index, contents in all_contents.items():
        output_filename = f"output-{folder_index}.csv"
        output_filename_html = f"output-{folder_index}.html"
        print(f"Creating: {output_filename}")

        output_result_path = os.path.join(output_result_folder, output_filename)
        

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
            dataCreateCSV["Categorie__c"].append("Auto Import")
            dataCreateCSV["Category__c"].append("Knowledge Material")

        df = pd.DataFrame(dataCreateCSV)
        df.to_csv(output_result_path, index=False, sep=",", quoting=csv.QUOTE_NONNUMERIC, quotechar='"', escapechar="\\")

        for index, content in enumerate(contents):
            with open(os.path.join(output_result_folder, output_filename_html), "w", encoding="utf-8") as f:
                f.write(content["content"])

    ## create result sum all file to one csv file.
    sum_all_content = []
    for item in all_contents.items():
        key, value = item 
        if len(item) == 0:
            return
        for sheet in value:
            sum_all_content.append((sheet))
    
    dataCreateCSVOne = {
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
    
    for index, obj in enumerate(sum_all_content): 
        if obj["is_field_exceed"]:
            print("found obj that has field exceed")
            print(obj["title"], end=" | ")
            continue
        ts = datetime.now().strftime("%Y%m%d%H%M%S%f")[:17]
        urlMock = f"URL-{ts}{index}"
        
        dataCreateCSVOne["Knowledge__kav"].append(index)
        dataCreateCSVOne["Id"].append("test")
        dataCreateCSVOne["RecordTypeId"].append("012N00000036GnwIAE")

        dataCreateCSVOne["Title"].append(obj["title"])
        dataCreateCSVOne["UrlName"].append(urlMock)
        dataCreateCSVOne["Summary"].append(obj["title"])
        dataCreateCSVOne["Answer"].append(obj["content"])
        dataCreateCSVOne["Categorie__c"].append("Auto Import")
        dataCreateCSVOne["Category__c"].append("Knowledge Material")
    
    df = pd.DataFrame(dataCreateCSVOne)
    name_result_one_csv = "output.csv"
    path_of_one_file = os.path.join(output_result_folder, name_result_one_csv)
    df.to_csv(path_of_one_file, index=False, sep=",", quoting=csv.QUOTE_NONNUMERIC, quotechar='"', escapechar="\\")
    print(f"Creating: {name_result_one_csv}")
    print("One for all CSV file created successfully!")


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
# %%
