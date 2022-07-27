import os
import pandas as pd
import glob
import numpy as np
from datetime import datetime
 
my_list = []
my_db = pd.DataFrame()

df = pd.DataFrame()
temp_folder = r'\\frwy\main\MXTJ\ACCT\Acct\Acct2022\Freeway\Temp'
f_list = r'\\frwy\main\MXTJ\ACCT\Acct\Acct2022\Trust\Trust check list.xlsx'
#temp_folder = r'\\frwy\main\MXTJ\ACCT\Acct\Acct2020\Freeway\CC Batches'
destination_Folder = r'\\frwy\main\MXTJ\ACCT\Acct\Acct2022\Projects\Trust Categories\Dashboard'
# ____________________________________________________________________________________
# select month and start date ************************
selected_month = '2022-07'                          #*
#*****************************************************

txt_m = int(selected_month[-2:])
txt_y = int(selected_month[2:4])
start_date = datetime(txt_y, txt_m, 1).strftime('%x') ### CONVERT String to date
#start_date = datetime.strptime(start_date, "%m/%d/%y")

# Create list of chosen month work book in TRUST
for i in glob.glob(temp_folder + r'\*\*' + selected_month + '*.xlsx'):
  if not "~$" in i and  not "Copy" in i:  
    # Adding file path to list 
    my_list.append(i)
    
    # read Sheet and find Trust Sheet
    my_book = pd.ExcelFile(i)
    found_sheet = False
    trust_sheet_name = None
    for sht in my_book.sheet_names:
        if 'Trust' in sht or 'Operating' in sht or 'OP' in sht:
            trust_sheet_name = sht
            found_sheet = True
            print('file %s has Trust sheet' % i)
            break
        elif found_sheet == True:
            exit
            break
    
    # Add found sheet to dataFrame
    if not trust_sheet_name is None:

        df = pd.read_excel(i, sheet_name=trust_sheet_name, header=None,
                               names=['Account','Name','Date','vc','Amount','Category','BAI Code','Type','BAI Description','Note'],
                               usecols="A:J")
                              #usecols=[0, 1, 2, 3, 4, 5, 6, 7, 8])

        
        # Filter onyl dates starting from the 1srt of selected month
        #df = df[df['Date'] >= start_date]
        

        # Add file name column 
        file_name = i.replace(temp_folder + '\\' , "")   
        df['file_name'] = file_name
        #my_db = my_db.append(df, ignore_index=True)
        my_db = pd.concat([my_db, df], ignore_index=True)
        df = None
        #except Exception:
        #    print('%s was empty. ' %i)
    else:
        print('file %s has no Trust sheet' % i)


 # vlookup for only valid accounts
df2 = pd.read_excel(f_list , sheet_name="Acc")
my_db = pd.merge(my_db, df2[['Name','Status']], on='Name', how='left')

my_db = my_db[my_db['Status'].str.contains('DO NOT USE')==False]  # mod 03.03.22



# Display data >>>>>>>>>>>>
#table = pd.pivot_table(my_db, values = 'Amount', index = ['Account'], columns = ['Date'], aggfunc = 'count')

# Export Data >>>>>>>>>>>>>>>>

# Saved Merge Files in one
timemark = datetime.today().strftime('%Y-%m-%d %H.%M.%S')
xl_file_name = r'\TrustDash' + timemark + '.xlsx'
#table.to_excel(destination_Folder + r'\TrustDash' + timemark + '.xlsx') 
#pylint: disable=abstract-class-instantiated
with pd.ExcelWriter(destination_Folder + r'\TrustDash' + timemark + '.xlsx') as writer:
    #table.to_excel(writer, sheet_name='Sheet1')
    my_db.to_excel(writer, sheet_name='Sheet2')



#__________________________________________________________________________________________________________________________
# END: 
# open distination folder________________________________________________________________________________________________ 
d_path = os.path.realpath(destination_Folder)
os.startfile(d_path)
os.system('start excel.exe "' + destination_Folder + xl_file_name + '"')

