# Load Folder to MySQL
Iterates through a folder of files and loads them to MySQL server, one-at-a-time.
Requires specification of desired table name, seperator (for seperated files), and Sheet Name (for excel files).

## Important Notes
You must create the target database before running the script.

## Supported Filetypes
Excel
- xlsx
- xlsb (binary excel files)
Seperated Text Files
- Comma-seperated (csv)
- Pipe Delimited
- Other normal-character delimiters

## Requirements
Also present in requirements.txt
- pandas
- sqlalchemy
- pymysql
- pyxlsb(if using binary excel files)

