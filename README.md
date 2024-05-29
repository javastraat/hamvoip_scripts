# Hamvoip Directory Tool

Export Extension data from pdf to csv for hamvoip hamvoip.nl dapnet

change password before usage

This is a command-line tool for generating various CSV and XML files used with a Hamvoip Phone System. It fetches extension data from a PDF hosted on the Hamvoip website, decrypts it, and extracts extension information. Then, based on the provided command-line arguments, it generates different output files such as CSVs for different phone systems like Cisco, Yealink, and Fanvil, as well as an XML file for Cisco phones. Additionally, it provides an option to remove the generated files. The script is designed to be versatile, allowing the user to generate specific files or all files at once.

## Usage

```bash
usage: hamvoip_directory_tool.py [-h] [-a] [-c] [-d] [-f] [-o] [-u] [-y] [-r]

python hamvoip_directory_tool.py [options]
```
## Options
Available Options:
```
-a, --all: Generate all available CSVs and XML files.
-c, --cisco: Generate hamvoip_cisco.xml file for Cisco phones.
-d, --dapnet: Generate hamvoip_dapnet.csv file for the DAPNET system.
-f, --fanvil: Generate hamvoip_fanvil.csv file for Fanvil phones.
-o, --other: Generate hamvoip_other.csv file with other then user extensions.
-u, --users: Generate hamvoip_users.csv file with user extension and information.
-y, --yealink: Generate hamvoip_yealink.csv file for Yealink phones.
-r, --remove: Remove all generated CSV and XML files.
```
<text>
-a, --all: This argument generates all available CSVs and XML files. If this argument is provided, the script generates CSVs for Cisco, Yealink, Fanvil, and Other phone systems, as well as an XML file for Cisco phones.

-c, --cisco: This argument generates the hamvoip_cisco.xml file, which is specifically formatted for Cisco IP phones.

-d, --dapnet: This argument generates the hamvoip_dapnet.csv file, which is a CSV file used for the DAPNET (Decentralized Amateur Paging Network) system.

-f, --fanvil: This argument generates the hamvoip_fanvil.csv file, which is a CSV file used for Fanvil phones.

-o, --other: This argument generates the hamvoip_other.csv file, which contains other extensions (four or more digits).

-u, --users: This argument generates the hamvoip_users.csv file, which contains user extension with information such as extension, callsign, and name.

-y, --yealink: This argument generates the hamvoip_yealink.csv file, which is a CSV file used for Yealink phones.

-r, --remove: This argument removes all generated CSV and XML files. If this argument is provided, the script deletes the following files: hamvoip_users.csv, hamvoip_cisco.xml, hamvoip_fanvil.csv, hamvoip_other.csv, hamvoip_dapnet.csv, and hamvoip_yealink.csv.
</text>
## Requirements
Python 3.11</br>
Required pip modules: requests, pypdf2, pdfplumber, pandas
