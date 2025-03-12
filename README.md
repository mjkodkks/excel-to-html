# EXCEL TO HTML
## Requirements
- python 3.11 or above
- install LibreOffice https://www.libreoffice.org/download/download-libreoffice/
- map soffice to executable location.

### Map soffice MacOS
```
sudo ln -s /Applications/LibreOffice.app/Contents/MacOS/soffice /usr/local/bin/soffice

soffice --version
```

## How to run
- from your python and pip
```
pip install -r requirements.txt

```
- copy all excel file to folder ```excel_files```

```
python app.py

```

- check your output from ```output_html```