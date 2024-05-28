#!/opt/homebrew/bin/python3.11
# Required pip modules:
# pip install requests
# pip install pypdf2
# pip install pdfplumber
# pip install pandas

import argparse
import os
import requests
from PyPDF2 import PdfReader, PdfWriter
import io
import pdfplumber
import re
import pandas as pd
import csv

def parse_args():
    parser = argparse.ArgumentParser(description="Hamvoip Directory Tool")
    parser.add_argument("-a", "--all", action="store_true", help="Generate all CSVs and XML")
    parser.add_argument("-c", "--cisco", action="store_true", help="Generate hamvoip_cisco.xml")
    parser.add_argument("-d", "--dapnet", action="store_true", help="Generate hamvoip_dapnet.csv")
    parser.add_argument("-f", "--fanvil", action="store_true", help="Generate hamvoip_fanvil.csv")
    parser.add_argument("-o", "--other", action="store_true", help="Generate hamvoip_other.csv")
    parser.add_argument("-r", "--remove", action="store_true", help="Remove CSV and XML files")
    parser.add_argument("-u", "--users", action="store_true", help="Generate hamvoip_users.csv")
    
    args = parser.parse_args()

    # Show help if no argument is given
    if not any(vars(args).values()):
        parser.print_help()
        parser.exit()

    return args

def fetch_extensions_pdf_url():
    response = requests.get('https://hamvoip.nl/download.php')
    if response.status_code == 200:
        page_content = response.text
        extensions_link = re.search(r'<a.*?href="downloads/extentions_(\d+\.\d+).pdf".*?>Download</a>', page_content, re.IGNORECASE)
        if extensions_link:
            version_number = extensions_link.group(1)
            return f"https://hamvoip.nl/downloads/extentions_{version_number}.pdf", version_number
        else:
            print("No extensions download URL found.")
            exit(1)
    else:
        print(f"Failed to fetch the download page. Status code: {response.status_code}")
        exit(1)

def download_decrypt_pdf(url, password):
    response_pdf = requests.get(url)
    if response_pdf.status_code == 200:
        pdf_data = io.BytesIO(response_pdf.content)
    else:
        raise Exception("Failed to download the PDF file.")

    pdf_reader = PdfReader(pdf_data)
    if pdf_reader.is_encrypted:
        pdf_reader.decrypt(password)
        pdf_writer = PdfWriter()
        for page_num in range(len(pdf_reader.pages)):
            pdf_writer.add_page(pdf_reader.pages[page_num])
        decrypted_pdf_data = io.BytesIO()
        pdf_writer.write(decrypted_pdf_data)
        decrypted_pdf_data.seek(0)
    else:
        decrypted_pdf_data = pdf_data

    return decrypted_pdf_data

def extract_text_from_pdf(pdf_data):
    extracted_text = []
    with pdfplumber.open(pdf_data) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            extracted_text.append(text)
    return "\n".join(extracted_text)

def extract_extensions(text):
    pattern_3digits = re.compile(r'(3\d{2}) - ([A-Z0-9]+) - ([^0-9\n]+)')
    pattern_4digits = re.compile(r'(\d{4,}) - (.*?)(?=\d{4,} -|\n|$)')

    data_3digits = []
    data_4digits = []
    data_longer_than_4digits = []

    matches_3digits = pattern_3digits.findall(text)
    for match in matches_3digits:
        number, callsign, name = match
        data_3digits.append([int(number), callsign.strip().upper(), name.strip()])

    matches_4digits = pattern_4digits.findall(text)
    for match in matches_4digits:
        number, name = match
        number = int(number)
        if len(str(number)) == 4:
            data_4digits.append([number, name.strip()])
        else:
            data_longer_than_4digits.append([number, name.strip()])

    return data_3digits, data_4digits, data_longer_than_4digits

def generate_users_csv(data_3digits):
    if data_3digits:
        df_users = pd.DataFrame(data_3digits, columns=['extension', 'callsign', 'name'])
        df_users = df_users[['extension', 'callsign', 'name']]
        df_users.sort_values(by='extension', inplace=True)
        df_users.to_csv('hamvoip_users.csv', index=False, header=['Extension', 'Callsign', 'Name'])
        print_summary('hamvoip_users.csv', len(data_3digits))

def generate_other_csv(data_4digits, data_longer_than_4digits):
    if data_4digits or data_longer_than_4digits:
        df_other = pd.DataFrame(data_4digits + data_longer_than_4digits, columns=['Extension', 'Name'])
        df_other.sort_values(by='Extension', inplace=True)
        df_other.to_csv('hamvoip_other.csv', index=False)
        print_summary('hamvoip_other.csv', len(data_4digits) + len(data_longer_than_4digits))

def generate_fanvil_csv(data_3digits, data_4digits, data_longer_than_4digits):
    if data_3digits or data_4digits or data_longer_than_4digits:
        combined_data = [[ext, f"{callsign} - {name}"] for ext, callsign, name in data_3digits] + data_4digits + data_longer_than_4digits
        df_fanvil = pd.DataFrame(combined_data, columns=['work', 'name'])
        df_fanvil['mobile'] = ''
        df_fanvil['other'] = ''
        df_fanvil['ring'] = 'Default'
        df_fanvil['groups'] = ''
        df_fanvil.sort_values(by='work', inplace=True)
        df_fanvil.to_csv('hamvoip_fanvil.csv', index=False, header=["name", "work", "mobile", "other", "ring", "groups"], quoting=csv.QUOTE_NONNUMERIC)
        print_summary('hamvoip_fanvil.csv', len(combined_data))

def generate_cisco_xml(data_3digits, data_4digits, data_longer_than_4digits):
    combined_data = [[ext, f"{callsign} - {name}"] for ext, callsign, name in data_3digits] + data_4digits + data_longer_than_4digits
    df_cisco = pd.DataFrame(combined_data, columns=['Extension', 'Name'])
    with open('hamvoip_cisco.xml', 'w') as f:
        f.write('<CiscoIPPhoneDirectory>\n')
        f.write('<Title>Hamvoip Directory</Title>\n')
        f.write('<Prompt>Please select number to dial…</Prompt>\n')
        for index, row in df_cisco.iterrows():
            f.write('<DirectoryEntry>\n')
            f.write(f'<Name>{row["Name"]}</Name>\n')
            f.write(f'<Telephone>{row["Extension"]}</Telephone>\n')
            f.write('</DirectoryEntry>\n')
        f.write('</CiscoIPPhoneDirectory>\n')
        print_summary('hamvoip_cisco.xml', len(combined_data))

def generate_dapnet_csv(data_3digits):
    if data_3digits:
        df_users = pd.DataFrame(data_3digits, columns=['Extension', 'Callsign', 'Name'])
        df_users = df_users[['Extension', 'Callsign']]
        df_users.sort_values(by='Extension', inplace=True)
        df_users.to_csv('hamvoip_dapnet.csv', index=False, header=['extension', 'callsign'])
        print_summary('hamvoip_dapnet.csv', len(data_3digits))

def remove_files():
    files_to_remove = ['hamvoip_users.csv', 'hamvoip_cisco.xml', 'hamvoip_fanvil.csv', 'hamvoip_other.csv', 'hamvoip_dapnet.csv']
    for file_path in files_to_remove:
        if os.path.exists(file_path):
            os.remove(file_path)
            print(f"Removed: {file_path}")
        else:
            print(f"Not found: {file_path}")

def print_summary(file_name, entries):
    print(f"File '{file_name}' created with {entries} entries.")

# Main script execution
args = parse_args()

if args.remove:
    remove_files()
    exit()

if not any(vars(args).values()):
    print("No options provided. Use --help for usage information.")
    exit()

pdf_url, version_number = fetch_extensions_pdf_url()
password = 'passw0rd'
decrypted_pdf_data = download_decrypt_pdf(pdf_url, password)
pdf_text = extract_text_from_pdf(decrypted_pdf_data)
data_3digits, data_4digits, data_longer_than_4digits = extract_extensions(pdf_text)

if args.users:
    generate_users_csv(data_3digits)
elif args.fanvil:
    generate_fanvil_csv(data_3digits, data_4digits, data_longer_than_4digits)
elif args.cisco:
    generate_cisco_xml(data_3digits, data_4digits, data_longer_than_4digits)
elif args.other:
    generate_other_csv(data_4digits, data_longer_than_4digits)
elif args.dapnet:
    generate_dapnet_csv(data_3digits)
elif args.all:
    generate_users_csv(data_3digits)
    generate_fanvil_csv(data_3digits, data_4digits, data_longer_than_4digits)
    generate_cisco_xml(data_3digits, data_4digits, data_longer_than_4digits)
    generate_other_csv(data_4digits, data_longer_than_4digits)
    generate_dapnet_csv(data_3digits)

print("Process completed successfully.")
