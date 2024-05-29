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
    """
    Parse and return command line arguments.
    """
    parser = argparse.ArgumentParser(description="Hamvoip Directory Tool")
    parser.add_argument("-a", "--all", action="store_true", help="Generate all CSVs and XML")
    parser.add_argument("-c", "--cisco", action="store_true", help="Generate hamvoip_cisco.xml")
    parser.add_argument("-d", "--dapnet", action="store_true", help="Generate hamvoip_dapnet.csv")
    parser.add_argument("-f", "--fanvil", action="store_true", help="Generate hamvoip_fanvil.csv")
    parser.add_argument("-o", "--other", action="store_true", help="Generate hamvoip_other.csv")
    parser.add_argument("-u", "--users", action="store_true", help="Generate hamvoip_users.csv")
    parser.add_argument("-y", "--yealink", action="store_true", help="Generate hamvoip_yealink.csv")
    parser.add_argument("-r", "--remove", action="store_true", help="Remove CSV and XML files")
    
    args = parser.parse_args()

    # Show help if no argument is given
    if not any(vars(args).values()):
        parser.print_help()
        parser.exit()

    return args

def fetch_extensions_pdf_url():
    """
    Fetch and return the URL of the latest extensions PDF from the Hamvoip website.
    """
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
    """
    Download and decrypt the PDF from the given URL using the provided password.
    """
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
    """
    Extract and return text from the decrypted PDF data.
    """
    extracted_text = []
    with pdfplumber.open(pdf_data) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            extracted_text.append(text)
    return "\n".join(extracted_text)

def extract_extensions(text):
    """
    Extract and return 3-digit, 4-digit, and longer-than-4-digit extensions from the extracted text.
    """
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
    """
    Generate hamvoip_users.csv file from 3-digit extensions data.
    """
    if data_3digits:
        df_users = pd.DataFrame(data_3digits, columns=['extension', 'callsign', 'name'])
        df_users = df_users[['extension', 'callsign', 'name']]
        df_users.sort_values(by='extension', inplace=True)
        df_users.to_csv('hamvoip_users.csv', index=False, header=['Extension', 'Callsign', 'Name'])
        print_summary('hamvoip_users.csv', len(data_3digits))

def generate_other_csv(data_4digits, data_longer_than_4digits):
    """
    Generate hamvoip_other.csv file from 4-digit and longer-than-4-digit extensions data.
    """
    if data_4digits or data_longer_than_4digits:
        df_other = pd.DataFrame(data_4digits + data_longer_than_4digits, columns=['Extension', 'Name'])
        df_other.sort_values(by='Extension', inplace=True)
        df_other.to_csv('hamvoip_other.csv', index=False)
        print_summary('hamvoip_other.csv', len(data_4digits) + len(data_longer_than_4digits))

def generate_fanvil_csv(data_3digits, data_4digits, data_longer_than_4digits):
    """
    Generate hamvoip_fanvil.csv file from 3-digit, 4-digit, and longer-than-4-digit extensions data.
    """
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
    """
    Generate hamvoip_cisco.xml file from 3-digit, 4-digit, and longer-than-4-digit extensions data.
    """
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
    """
    Generate hamvoip_dapnet.csv file from 3-digit extensions data.
    """
    if data_3digits:
        df_dapnet = pd.DataFrame(data_3digits, columns=['Extension', 'Callsign', 'Name'])
        df_dapnet = df_dapnet[['Extension', 'Callsign']]
        df_dapnet.sort_values(by='Extension', inplace=True)
        df_dapnet.to_csv('hamvoip_dapnet.csv', index=False, header=['extension', 'callsign'])
        print_summary('hamvoip_dapnet.csv', len(data_3digits))

def generate_yealink_csv(data_3digits, data_4digits, data_longer_than_4digits):
    """
    Generate hamvoip_yealink.csv file from 3-digit, 4-digit, and longer-than-4-digit extensions data.
    """
    combined_data = [[ext, f"{callsign} - {name}"] for ext, callsign, name in data_3digits] + data_4digits + data_longer_than_4digits
    df_yealink = pd.DataFrame(combined_data, columns=['Extension', 'Name'])
    df_yealink['display_name'] = df_yealink['Name']
    df_yealink['office_number'] = df_yealink['Extension']
    df_yealink['mobile_number'] = ''
    df_yealink['other_number'] = ''
    df_yealink['line'] = '2'
    df_yealink['ring'] = 'Auto'
    df_yealink['auto_divert'] = ''
    df_yealink['priority'] = ''
    df_yealink['group_id_name'] = 'HamVoip'
    df_yealink['default_photo'] = 'Config:logo-hamvoip.jpg'
    df_yealink['photo_data'] = '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'

    df_yealink = df_yealink[['display_name', 'office_number', 'mobile_number', 'other_number', 'line', 'ring', 'auto_divert', 'priority', 'group_id_name', 'default_photo', 'photo_data']]
    df_yealink.sort_values(by='office_number', inplace=True)
    df_yealink.to_csv('hamvoip_yealink.csv', index=False)
    print_summary('hamvoip_yealink.csv', len(combined_data))

def remove_files():
    """
    Remove generated CSV and XML files.
    """
    files_to_remove = ['hamvoip_users.csv', 'hamvoip_cisco.xml', 'hamvoip_fanvil.csv', 'hamvoip_other.csv', 'hamvoip_dapnet.csv', 'hamvoip_yealink.csv']
    for file_path in files_to_remove:
        if os.path.exists(file_path):
            os.remove(file_path)
            print(f"Removed: {file_path}")
        else:
            print(f"Not found: {file_path}")

def print_summary(file_name, entries):
    """
    Print a summary of the generated file and the number of entries it contains.
    """
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

pdf_url, version_number = fetch_extensions_pdf_url()
print(f"Downloaded Hamvoip extensions PDF v{version_number}")

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
elif args.yealink:
    generate_yealink_csv(data_3digits, data_4digits, data_longer_than_4digits)
elif args.all:
    generate_users_csv(data_3digits)
    generate_other_csv(data_4digits, data_longer_than_4digits)
    generate_dapnet_csv(data_3digits)
    generate_fanvil_csv(data_3digits, data_4digits, data_longer_than_4digits)
    generate_yealink_csv(data_3digits, data_4digits, data_longer_than_4digits)
    generate_cisco_xml(data_3digits, data_4digits, data_longer_than_4digits)

print("Process completed successfully.")
