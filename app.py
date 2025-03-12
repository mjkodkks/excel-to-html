# %%
import glob
import os
import re
import subprocess
import time
import minify_html
import pandas as pd
import shutil
from bs4 import BeautifulSoup

input_folder = "excel_files"  # Folder containing Excel files
output_folder = "output_html"  # Folder to save HTML outputs

def excel_to_HTML():
    print("## Start EXCEL TO HTML ##")
    # Remove all item in output_html
    shutil.rmtree(output_folder, ignore_errors=True)

     # Ensure output directory exists
    os.makedirs(output_folder, exist_ok=True)

    # Get all Excel files in the folder
    excel_files = [f for f in os.listdir(input_folder) if f.endswith((".xlsx", ".xls"))]

    # Convert each Excel file separately
    for index, file in enumerate(excel_files):
        print(file)
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

# Read all HTML files in the folder
def read_all_html_files(file_pattern=output_folder + '/**/*.html', recursive=True):
    print("read_all_html_files")
    os.makedirs(output_folder, exist_ok=True)
    files = glob.glob(file_pattern, recursive=recursive)
    print('allfile')
    print(files)
    contentsList = []
    for file in files:
        print(f"--- {file} ---")
        try:
            filename = os.path.basename(file)
            with open(file, 'r', encoding='utf-8') as f:
                contentTxt = f.read()
                print(len(contentTxt))
                soup = BeautifulSoup(contentTxt, "html.parser")
                body_content = soup.body.prettify().replace(",","")
                print(len(body_content))
                minified = minify_html.minify(body_content, minify_js=True, minify_css=True, remove_processing_instructions=True)
                print(len(minified))
                contentsList.append({
                    "title": filename,
                    "content": minified
                })
        except Exception as e:
            print(f"Error reading {file}: {e}")
        print("\n" + "="*50 + "\n")  # Separator for readability

    # print(contentsList)
    return contentsList

# Create output file from content
def create_output_file_from_content(contents: list[dict[str, str]], output_name= "output.csv", is_salesforce=False):
    print("create_csv_from_content")
    print(contents)
    if not contents:
        return 
    
    if is_excel_or_csv(output_name) is False:
        print("file type not support")
        return

    if os.path.exists(output_name):
        os.remove(output_name)
        print("Already delete previous output file")
    else:
        print("The file does not exist")

    if is_salesforce: 
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
        
        for content in contents:
            ts = round(time.time())
            urlMock = "URL-" + str(ts)
            dataCreateCSV["Knowledge__kav"].append("")
            dataCreateCSV["Id"].append("")
            dataCreateCSV["RecordTypeId"].append("012N00000036GnwIAE")
            dataCreateCSV["Title"].append(content["title"] + "(test-html-import)")
            dataCreateCSV["UrlName"].append(urlMock)
            dataCreateCSV["Summary"].append(content["title"])
            dataCreateCSV["Answer"].append(content["content"])
            dataCreateCSV["Categorie__c"].append("")
            dataCreateCSV["Category__c"].append("Knowledge Material")
    else:
        dataCreateCSV = {
            "Title": [],
            "Content": [],
        }   

        for content in contents:
            dataCreateCSV["Title"].append(content["title"])
            dataCreateCSV["Content"].append(content["content"])

    df = pd.DataFrame(dataCreateCSV)

    if (is_excel(output_name)):
        df.to_excel(output_name)
    
    elif (is_csv(output_name)):
        df.to_csv(output_name, index=True, quotechar="'")


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
readContent = read_all_html_files()
create_output_file_from_content(readContent,is_salesforce=True)


# %%


