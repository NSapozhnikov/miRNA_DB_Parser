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
                          'Ortholog of target gene': [],
                          'Gene name': [],
                          '3P-seq tags + 5': [],
                          'Conservative total sites': [],
                          'Conserved 8mer sites': [],
                          'Conserved 7mer-m8 sites': [],
                          'Conserved 7mer-A1 sites': [],
                          'Poorly conserved sites total': [],
                          'Poorly conserved 8mer sites': [],
                          'Poorly conserved 7mer-m8 sites': [],
                          'Poorly conserved 7mer-A1 sites': [],
                          '6mer sites': [],
                          'Representative miRNA': [],
                          'Cumulative weighted context++ score': [],
                          'Target total context++ score': []})
    pd_df.to_excel(writer, sheet_name='TargetScan', header=True, index=False)
    
    in_values = pd.read_excel('total_m_mi.xlsx', sheet_name='total miRNA')['miRNA']
    num = 0
    max_row = 1
    empty_table = []
    no_page = []
    no_table = []
    for miRNA in in_values:
        num +=1
        pd_df = pd.DataFrame({'miRNA': [],                                                       
                              'Ortholog of target gene': [],
                              'Gene name': [],
                              '3P-seq tags + 5': [],
                              'Conservative total sites': [],
                              'Conserved 8mer sites': [],
                              'Conserved 7mer-m8 sites': [],
                              'Conserved 7mer-A1 sites': [],
                              'Poorly conserved sites total': [],
                              'Poorly conserved 8mer sites': [],
                              'Poorly conserved 7mer-m8 sites': [],
                              'Poorly conserved 7mer-A1 sites': [],
                              '6mer sites': [],
                              'Representative miRNA': [],
                              'Cumulative weighted context++ score': [],
                              'Target total context++ score': []})
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
            
            pd_df1 = pd.read_excel(f"{miRNA}_output.xlsx")
            pd_df1 = pd_df1.drop(columns=['Representative transcript'])                            ### drop useless columns
            pd_df1 = pd_df1.loc[pd_df1['Total context++ score']<= -0.5]                            ### filter for score values
            if len(pd_df1.columns) == 17:
                pd_df['Ortholog of target gene'] = pd_df1['Ortholog of target gene']
                pd_df['Gene name'] = pd_df1['Gene name']
                pd_df['3P-seq tags + 5'] = pd_df1['3P-seq tags + 5']
                pd_df['Conservative total sites'] = pd_df1['Conserved sites total']
                pd_df['Conserved 8mer sites'] = pd_df1['Conserved 8mer sites']
                pd_df['Conserved 7mer-m8 sites'] = pd_df1['Conserved 7mer-m8 sites']
                pd_df['Conserved 7mer-A1 sites'] = pd_df1['Conserved 7mer-A1 sites']
                pd_df['Poorly conserved sites total'] = pd_df1['Poorly conserved sites total']
                pd_df['Poorly conserved 8mer sites'] = pd_df1['Poorly conserved 8mer sites']
                pd_df['Poorly conserved 7mer-m8 sites'] = pd_df1['Poorly conserved 7mer-m8 sites']
                pd_df['Poorly conserved 7mer-A1 sites'] = pd_df1['Poorly conserved 7mer-A1 sites']
                pd_df['6mer sites'] = pd_df1['6mer sites']
                pd_df['Representative miRNA'] = pd_df1['Representative miRNA']
                pd_df['Cumulative weighted context++ score'] = pd_df1['Cumulative weighted context++ score']
                pd_df['Target total context++ score'] = pd_df1['Total context++ score']                         ### adding content for large tables
            elif len(pd_df1.columns) == 12:
                pd_df['Ortholog of target gene'] = pd_df1['Ortholog of target gene']
                pd_df['Gene name'] = pd_df1['Gene name']
                pd_df['3P-seq tags + 5'] = pd_df1['3P-seq tags + 5']
                pd_df['Conservative total sites'] = pd_df1['Total sites']
                pd_df['Conserved 8mer sites'] = pd_df1['8mer sites']
                pd_df['Conserved 7mer-m8 sites'] = pd_df1['7mer-m8 sites']
                pd_df['Conserved 7mer-A1 sites'] = pd_df1['7mer-A1 sites']     
                pd_df['6mer sites'] = pd_df1['6mer sites']
                pd_df['Representative miRNA'] = pd_df1['Representative miRNA']
                pd_df['Cumulative weighted context++ score'] = pd_df1['Cumulative weighted context++ score']
                pd_df['Target total context++ score'] = pd_df1['Total context++ score']                         ### adding content for small tables            
            pd_df.loc[:, 'miRNA'] = miRNA                                                                       ### adding the requested miRNA
            print(f"Adding {len(pd_df.index)} rows for {miRNA}...")                                                                   
            if len(pd_df.index) == 0:
                empty_table.append(miRNA)                                                        

            pd_df.to_excel(writer, sheet_name='TargetScan', header=False, index=False, startrow=max_row)      ### write to the outfile
            os.remove(f"{miRNA}_output.xlsx")                                                   ### delete downloaded file
            max_row += len(pd_df.index)
            
        except AttributeError:
            no_page.append(miRNA)
            print(f"No {miRNA} was found in TargetScan DB... Skipping...")                       ### if None info was found for this miRNA
        except IndexError:
            no_table.append(miRNA)
            print(f"No downloadable table was found for {miRNA}... Skipping...")                ### no download link is found
        except ConnectionError:
            print(f"Connection error for {miRNA}. Skipping...")
        finally:
            time.sleep(0.5)
        print(f"{num}/{len(in_values)}", end='\n')
    time_end = datetime.datetime.now()
    print('No rows were added for this list of miRNA\'s: ')
    for k in empty_table:
        print(k, end=' ')
    print('No page or downloadable link has been found for: ')
    for i in no_page:
        print(i, end=' ')
    for j in no_table:
        print(j, end=' ')
print(f"\nTime of finish - {time_end.isoformat(sep=' ', timespec='seconds')}. Time of executing - {time_end - time_start}.")
print('\nDone!')
