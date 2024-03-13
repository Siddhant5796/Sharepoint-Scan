# Sharepoint-Scan

import requests
from requests_ntlm import HttpNtlmAuth
from openpyxl import load_workbook
import re
from bs4 import BeautifulSoup
 
def fetch_excel_files(sharepoint_folder_link, username, password):
    session = requests.Session()
    session.auth = HttpNtlmAuth(username, password)
    response = session.get(sharepoint_folder_link)
    excel_files = []
    if response.status_code == 200:
        print(response.content)
        soup = BeautifulSoup(response.content, 'html.parser')
        #all_links= re.findall(link_pattern, str(soup))
        regex = "https?:\\/\\/(?:www\\.)?[-a-zA-Z0-9@:%._\\+~#=]{1,256}\\.[a-zA-Z0-9()]{1,6}\\b(?:[-a-zA-Z0-9()@:%_\\+.~#?&\\/=]*)"
        all_links = re.findall(regex, str(soup))
        #print([x[0] for x in all_links])
        print("all_links")
        print(all_links)
        for link in soup.find_all('a', href=True):
            href = link['href']
            if href.endswith('.xlsx'):
                excel_files.append(href)
    else:
        print(f"Failed to fetch data from {sharepoint_folder_link}. Status code: {response.status_code}")
    return excel_files
def main():
    sharepoint_folder_link = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
    username = "xxxxxxxxxxxxxx"
    password = "xxxxxxxxxxxx"
    excel_files = fetch_excel_files(sharepoint_folder_link, username, password)
    if excel_files:
        print("Excel files found:")
        for file_url in excel_files:
            session = requests.Session()
            session.auth = HttpNtlmAuth(username, password)
            response = session.get(file_url)
            if response.status_code == 200:
                wb = load_workbook(filename=BytesIO(response.content))
                for ws in wb.worksheets:
                    print("Worksheet:", ws.title)
                    for row in ws.iter_rows(values_only=True):
                        text = ' '.join(str(cell) for cell in row if cell)
                        pii_info = check_patterns(text)
                        if pii_info:
                            print("PII found in row:", pii_info)
            else:
                print(f"Failed to retrieve the file. Status code: {response.status_code}")
    else:
        print("No Excel files found.")
if __name__ == "__main__":
    main()
