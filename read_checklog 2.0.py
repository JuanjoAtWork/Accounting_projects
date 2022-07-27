# Read CHECKs Log from a share file to pivot in consolidation trust file _________________________________
import pandas as pd
import os
from datetime import datetime
from add_commchecks import *


#select_month = input("Enter month [##] :")
file_path = r"V:\Acct2022\Trust\Files\CheckLogs"
file_name = r"Confie - Corporate Deposit Log.xlsx"
#dest_name = r"check_log.xlsx"

comm_path = f"{file_path}\\{file_name}"
#dest_path = f"{file_path}\\{dest_name}"
#booleans = []
db2 = pd.DataFrame()
db1 = pd.read_excel(comm_path, sheet_name="1. Confie - Corporate Deposit L", usecols="A:N")

# FORMAT order columns in new dataFrame _________________________________________________
cols_pos=[1,2,0,4,5,6,7,3]
for i in cols_pos:
    #print(f"db2[{db1.columns[i]}] = db1[{db1.columns[i]}]")
    db2[db1.columns[i]] = db1[db1.columns[i]]


# FILTER selecting rows based on condition __________________________________________________    
# filter only description with 'COMM' or 'INCENTIE BONUES'
export_table = db2[db2['Description'].str.contains("COMM|INCENTIVE BONUS|INCENTIVE/BONUS|HEALTHCARE|MONTHLY PAYMENT|CONTINGENT",  na=False)]

# filter only date with selected month 
export_table = export_table[export_table['Date Deposited'].dt.month == int(select_month)]

# assign column with source
export_table['Source'] = "smartsheets"


# merge addtional checks and smartsheets
final_book = pd.concat([export_table, additional_book] , ignore_index=True )

# export to file__________________________________________________________
timemark = datetime.today().strftime('%Y-%m-%d %H.%M.%S')
dest_name = f"check_log {timemark}.xlsx"
dest_path = f"{file_path}\\{dest_name}"

with pd.ExcelWriter(dest_path) as writer:
    final_book.to_excel(writer, sheet_name='CheckLog')
    #table3.to_excel(writer, sheet_name='Sheet2')


#print('start excel.exe "' + dest_path + '"')
os.system('start excel.exe "' + dest_path + '"')

