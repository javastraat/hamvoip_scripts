#!/usr/bin/python3
# Required pip modules:
# pip install requests
# pip install pypdf2
# pip install pdfplumber
# pip install pandas

import requests
from PyPDF2 import PdfReader, PdfWriter
import io
import pdfplumber
import re
import pandas as pd

# Fetch the content of the page
response = requests.get('https://hamvoip.nl/download.php')
if response.status_code == 200:
    page_content = response.text
    # Search for the anchor tag containing "extensions" in the href attribute
    extensions_link = re.search(r'<a.*?href="downloads/extentions_(\d+\.\d+).pdf".*?>Download</a>', page_content, re.IGNORECASE)
    if extensions_link:
        version_number = extensions_link.group(1)
        download_url = f"https://hamvoip.nl/downloads/extentions_{version_number}.pdf"
        print("Found extensions download URL:", download_url)
    else:
        print("No extensions download URL found.")
        # Stop execution if no download URL found
        exit(1)
else:
    print(f"Failed to fetch the download page. Status code: {response.status_code}")
    # Stop execution if unable to fetch the download page
    exit(1)

# Define the password
password = 'passw0rd'

# Download the PDF
response_pdf = requests.get(download_url)
if response_pdf.status_code == 200:
    pdf_data = io.BytesIO(response_pdf.content)
else:
    raise Exception("Failed to download the PDF file.")

# Decrypt the PDF
pdf_reader = PdfReader(pdf_data)
if pdf_reader.is_encrypted:
    pdf_reader.decrypt(password)
    # Create a PdfWriter object to write the decrypted content
    pdf_writer = PdfWriter()
    for page_num in range(len(pdf_reader.pages)):
        pdf_writer.add_page(pdf_reader.pages[page_num])
    # Write the decrypted PDF to a new BytesIO object
    decrypted_pdf_data = io.BytesIO()
    pdf_writer.write(decrypted_pdf_data)
    decrypted_pdf_data.seek(0)
else:
    decrypted_pdf_data = pdf_data

# Extract Text with Formatting using pdfplumber
extracted_text = []
with pdfplumber.open(decrypted_pdf_data) as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        extracted_text.append(text)

# Combine extracted text into a single string
combined_text = "\n".join(extracted_text)

# Display combined text for debugging
#print("Combined text:")
#print(combined_text)

# Extract all valid entries using regular expressions for 3-digit, 4-digit, and longer extensions
pattern_3digits = re.compile(r'(3\d{2}) - ([A-Z0-9]+) - ([^0-9\n]+)')
pattern_4digits = re.compile(r'(\d{4,}) - (.*?)(?=\d{4,} -|\n|$)')

# Create lists to store extracted data
data_3digits = []
data_4digits = []
data_longer_than_4digits = []

# Extract 3-digit extensions and names
matches_3digits = pattern_3digits.findall(combined_text)
for match in matches_3digits:
    number, callsign, name = match
    data_3digits.append([int(number), callsign.strip().upper()])

# Extract 4-digit and longer extensions and names
matches_4digits = pattern_4digits.findall(combined_text)
for match in matches_4digits:
    number, name = match
    number = int(number)
    if len(str(number)) == 4:
        data_4digits.append([number, name.strip()])
    else:
        data_longer_than_4digits.append([number, name.strip()])

# Create DataFrames for each category
df_3digits = pd.DataFrame(data_3digits, columns=['Extension', 'Callsign'])
df_4digits = pd.DataFrame(data_4digits, columns=['Extension', 'Name'])
df_longer_than_4digits = pd.DataFrame(data_longer_than_4digits, columns=['Extension', 'Name'])

# Sort the 3-digit DataFrame by extension
df_3digits_sorted = df_3digits.sort_values(by='Extension')

# Save the sorted 3-digit extensions to "hamvoip_users.csv" with only callsign and extension columns
df_3digits_sorted.rename(columns={'Extension':'extension', 'Callsign':'callsign'}).to_csv('hamvoip_users.csv', index=False)

# Sort the other CSV by extension
df_4digits_sorted = df_4digits.sort_values(by='Extension')
df_longer_than_4digits_sorted = df_longer_than_4digits.sort_values(by='Extension')

# Save the sorted 4-digit and longer extensions to "hamvoip_other.csv"
df_4digits_sorted.to_csv('hamvoip_other.csv', index=False)
df_longer_than_4digits_sorted.to_csv('hamvoip_other.csv', mode='a', header=False, index=False)

# Show the number of extensions
print(f"Number of 3-digit extensions found: {len(df_3digits_sorted)}")
print(f"Number of 4-digit extensions found: {len(df_4digits_sorted)}")
print(f"Number of extensions longer than 4 digits found: {len(df_longer_than_4digits_sorted)}")
