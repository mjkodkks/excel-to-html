# %%
import glob
import os
import subprocess
import pandas as pd
import shutil

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
                contentsList.append({
                    "title": filename,  # Fix: use string keys
                    "content": contentTxt  # Fix: use string keys
                })
        except Exception as e:
            print(f"Error reading {file}: {e}")
        print("\n" + "="*50 + "\n")  # Separator for readability

    print(contentsList)
    return contentsList

# Create CSV from content
def create_csv_from_content(contents: list[dict[str, str]], output_name_csv= "output"):
    print("create_csv_from_content")
    print(contents)
    if not contents:
        return 

    output_name_csv = output_name_csv + ".csv"
    if os.path.exists(output_name_csv):
        os.remove(output_name_csv)
        print("Already delete csv")
    else:
        print("The file does not exist")

    dataCreateCSV = {
        "Title": [],
        "Content": [],
    }

    for content in contents:  # Fix: No need to use enumerate
        dataCreateCSV["Title"].append(content["title"])  # Fix: Use correct key
        dataCreateCSV["Content"].append(content["content"])  # Fix: Use correct key

    df = pd.DataFrame(dataCreateCSV)
    df.to_csv(output_name_csv, index=False)

excel_to_HTML()
readContent = read_all_html_files()
create_csv_from_content(readContent)


# %%



