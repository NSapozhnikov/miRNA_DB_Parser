# miRNA_DB_Parser
## 9.27.2021
## miRNA_DB_Parser
 miRNA_DB_Parser is a Python script that collects certain information from open miRNA databases (miRDB) and writes it to the outfile (.xlsx).
## TargetScan_DB_Parser
 TargetScan_DB_Parser.py downloads .xlsx file from TargetScan site for each miRNA in the input file to its directory, and after appending
 its content deletes each downloaded file.

### Requested modules:

-[pandas](https://pandas.pydata.org/docs/getting_started/install.html)

-[bs4](https://pypi.org/project/beautifulsoup4/)

-[requests](https://pypi.org/project/requests/)

-[openpyxl](https://openpyxl.readthedocs.io/en/stable/#installation)

## How to install
To run this script you'll need an installed [Python 3](https://www.python.org/downloads/). Place the .py script and .xlsx file in the
same directory. 

Run from CLI (python miRNA_DB_Parser.py)
