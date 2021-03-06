# miRNA_DB_Parser
## 9.27.2021

This project is a set Python scripts that collect certain information from open miRNA databases (miRDB, TargetScan) via WebScrapping and write it to the (.xlsx) outfiles.

## miRDB_Parser
 miRDB_Parser uses POST method (with reasonable frequency) to collect data from miRDB web site and write it to outfile.
## TargetScan_DB_Parser
 TargetScan_DB_Parser.py downloads .xlsx file from TargetScan site for each miRNA in the input file to its directory, and after appending
 its content deletes each downloaded file. 
 
 This method has been chosen due to bad HTML markup of TargetScan web site and impossibility to use any SQL databases.

### Requested modules:

-[pandas](https://pandas.pydata.org/docs/getting_started/install.html)

-[bs4](https://pypi.org/project/beautifulsoup4/)

-[requests](https://pypi.org/project/requests/)

-[openpyxl](https://openpyxl.readthedocs.io/en/stable/#installation)

## How to install
To run this script you'll need an installed [Python 3](https://www.python.org/downloads/). Place the .py script and .xlsx file in the
same directory. 

Run from CLI (python miRNA_DB_Parser.py)
