import requests
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import time
import re
import os
import sys
import datetime

time_start = datetime.datetime.now()
print (f"\nTime of start - {time_start.isoformat(sep=' ', timespec='seconds')}") 
with pd.ExcelWriter('TargetScan_output_table.xlsx') as writer:
    print('Writing to the outfile...')
    pd_df = pd.DataFrame({'miRNA': [],                                                       ### output file creation + header printing
                          'Gene_Symbol': [],
                          'Gene_Desciption': [],
                          'Target total context++ score': []})
    pd_df.to_excel(writer, sheet_name='TargetScan', header=False, index=False)
    
in_values = pd.read_excel('derg_de_mirna.xlsx', sheet_name='Ls - Lsk')['name']
num = 0
for miRNA in in_values:
    num +=1
    try:
        url_TargetScan = f"""http://www.targetscan.org/cgi-bin/targetscan/vert_72/targetscan.cgi?species=Rat&gid=&mir_sc=&mir_c=&mir_nc=&mir_vnc=&mirg={miRNA}"""
        resp = requests.get(url_TargetScan)                                                 ### GET the main URL page for this miRNA
        soup = BeautifulSoup(resp.text, 'html.parser')
        download_link_1 = soup.find_all(href=re.compile('species'), string='Download table')[0]
        url_download_link_1 = 'http://www.targetscan.org/cgi-bin/targetscan/vert_72/' + download_link_1.get('href')

        resp_download = requests.get(url_download_link_1)                                   ### GET page with downloadables

        soup = BeautifulSoup(resp_download.text, 'html.parser')
        download_link_2 = soup.find_all(href=re.compile('.xlsx'))[0]
        url_download_link_2 = 'http://www.targetscan.org/' + download_link_2.get('href')
        download_table = requests.get(url_download_link_2)                                  ### downloading .xlsx table for the miRNA
        open(f"{miRNA}_output.xlsx", 'wb').write(download_table.content)
        
        pd_df = pd.read_excel(f"{miRNA}_output.xlsx")
        pd_df = pd_df.loc[pd_df['Total context++ score']<= -0.5]                            ### filter for score values
        
        with pd.ExcelWriter('TargetScan_output_table.xlsx', engine='openpyxl', mode='a') as writer:
            book = load_workbook('TargetScan_output_table.xlsx')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            sheet = book.active
            pd_df.to_excel(writer, sheet_name='TargetScan', header=True, index=False, startrow=sheet.max_row+1)      ### writing to the outfile
            book.save('TargetScan_output_table.xlsx')
        os.remove(f"{miRNA}_output.xlsx")                                                   ### deleting downloaded file
        
    except AttributeError:                                                                  ### if None info was found for this miRNA
        continue
    finally:
        time.sleep(0.5)
    print(f"{num}/{len(in_values)}", end='\r')
time_end = datetime.datetime.now()
print(f"\nTime of finish - {time_end.isoformat(sep=' ', timespec='seconds')}. Time of executing - {time_end - time_start}.")
print('\nDone!')