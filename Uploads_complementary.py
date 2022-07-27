import os
import re
import glob
import pandas as pd
import numpy as np
from datetime import date, datetime

#____________________________________________________________________________________________________________
# INPUT:
#**********************************************
m = '07' # << type MONTH                      #*
d = '25' # << type DAY   {2digits}            #*
y = '22' # << type YEAR                       #*
current_date = datetime(int(y),int(m),int(d)) #*
#*********************************************

# Directory # 
destination_Folder = r'\\frwy\main\MXTJ\ACCT\Acct\Acct2022\Projects\Trust Categories\Dashboard'
db1 = pd.DataFrame()
db2 = pd.DataFrame()
db3 = pd.DataFrame()                
# Directory # 
#--------------------------------------------------------------------------#
# new _______________________________________________________________________________
my_db = pd.DataFrame()
log_f = []
dls_older = r'\\frwy\main\MXTJ\ACCT\Acct\Acct2021\Freeway\Banks Download'
f_names =['North West','Suntrust 1211','Planters','Georgia Banking 5703','Farmers 1139',
'Colony Bank','PNC Bank','Bank of Terell 3301', 'Synovus Bank', 'Regions 7611',
'Community 3053','Cadence Bank','Regions','PNC 6041','Community 2651','Prosperity']

trust_check_list = r'\\frwy\main\MXTJ\ACCT\Acct\Acct2022\Trust\Trust check list.xlsx'
#destination_Folder = r'V:\Acct2021\Projects\Deposits\bankDLs'

#____________________________________________________________________________________________________________
#_________________________________________________________________________________________________________
# FUNCTIONS:
#____________________________________________________________________________________________________________
# Set date from string
def valid_date(datestring):
    new_date = None
    #print(datestring + ' Type: ' + str(type(datestring)))
    try:
        mat=re.match('(\d+[-/]\d+[-/]\d+)', datestring)
        if mat is not None:
            this_date = mat.group(0)
            if '/' in this_date:
                new_date = datetime.strptime(this_date, '%m/%d/%Y')
                
            elif '-' in this_date:
                new_date = datetime.strptime(this_date, '%m-%d-%Y')
                
            return new_date
                    
    except ValueError:
        #new_date = datestring
        pass
        
    return new_date

# set arrange of columns
def find_col_date_in_Headers(s_file_name, f, df):
    sample_list = df.columns
    col_date_position = [sample_list.get_loc(i) for i in sample_list if 'Date' in str(i) or 'date' in str(i)]
    d1 = pd.DataFrame()
    if not col_date_position == []:
        #-new line-----------------------------------------------------#
        if not 'csv' in f:
            d1 = pd.read_excel(f, index_col=None, na_values=['NA'], usecols = col_date_position)
        else:
            d1 = pd.read_csv(f, index_col=None, na_values=['NA'], usecols = col_date_position)
        #-end new line-------------------------------------------------#
        d1.columns = ['bank_date']
        try:
            d1['bank_date'] = pd.to_datetime(d1['bank_date'], format='%m%d%Y')
        except:
            d1.fillna(0)
            d1['bank_date']=d1['bank_date'].astype(str)
            d1['bank_date'] = d1['bank_date'].apply(valid_date)

        d1['folder'] = s_file_name[0]
        d1['file'] = s_file_name[1]
        d1['last4'] = s_file_name[1].split(' ')[0]
    else:
        d1['bank_date'] = 'NA'
        d1['folder'] = s_file_name[0]
        d1['file'] = s_file_name[1]
        d1['last4'] = s_file_name[1].split(' ')[0]
    return d1

# Error in headers search for valid header names in first row 0 to 2 // Replace this with function 
def find_col_date_in_rows(s_file_name, f, df):
    for r in range(len(df)):
        df.columns = df.iloc[r]
        sample_list = df.columns
        col_date_position = [sample_list.get_loc(i) for i in sample_list if 'Date' in str(i) or 'date' in str(i)]
        d1 = pd.DataFrame()
        if not col_date_position == []:
            #-new line-----------------------------------------------------#
            if not 'csv' in f:
                d1 = pd.read_excel(f, index_col=None, na_values=['NA'], usecols = col_date_position)
            else:
                d1 = pd.read_csv(f, index_col=None, na_values=['NA'], usecols = col_date_position)
            #-end new line-------------------------------------------------#

            d1.columns = d1.iloc[r + 1]
            d1.columns = ['bank_date']
            
            try:
                d1['bank_date'] = pd.to_datetime(d1['bank_date'], format='%m%d%Y')
            except:
                d1.fillna(0)
                d1['bank_date']=d1['bank_date'].astype(str)
                d1['bank_date'] = d1['bank_date'].apply(valid_date)
            # #finally:
            # #   pass

            d1['folder'] = s_file_name[0]
            d1['file'] = s_file_name[1]
            d1['last4'] = s_file_name[1].split(' ')[0]
            break
        else:
           d1['bank_date'] = 'NA'
           d1['folder'] = s_file_name[0]
           d1['file'] = s_file_name[1]
           d1['last4'] = s_file_name[1].split(' ')[0]
    return d1


# # ______________________________________________________________________________________________________________________________#
# # BOA Bank _____________________________________________________________________________________________________________________#
# source_folder = r'\\frwy\main\MXTJ\ACCT\Acct\Acct2021\Freeway\Banks Download\BOA TN'
# # r'\*' + m + '.*' + d + '*.xls*'   was : r'\*07.13.*21.xls*'
# #files_list = glob.glob(source_folder + r'\*MAY MTD 2022.xls*') # monthlys
# files_list = glob.glob(source_folder + r'\*' + m + '.' + d + '.*' + y + '.xls*') # All
# for f in files_list:
#     if not "~$" in f:
#         print('adding: ' + f)
#         db = pd.read_excel(f, parse_dates=True)
#         db.columns = db.iloc[4]     # set row 4  as header 
#         db = db[(db['Row Type'] == 'Data') & (db['Data Type'] != 'Summary')]
        
#         # add column with file name 
#         file_name = f.replace(source_folder + '\\' , "")  
#         db['file_name'] = file_name
#         db['bank'] = "BOA"
#         db.rename(columns={"As of Date": "bank_date", "Account Number": "Account"}, inplace = True)  
#         db['bank_date'] = pd.to_datetime(db['bank_date'])  # convert string to dates 
#         #db['bank_date'] = pd.to_datetime(db['bank_date']).dt.strftime('%m/%d/%Y') # format date as mm/dd/yyyy
#         # compilation
#         #db1 = db1.append(db, ignore_index=True)
#         db1 = pd.concat([db1, db], ignore_index=True)
#         db = None

# # Display data >>>>>>>>>>>>  
# #db1['bank_date']=db1['bank_date'].astype(str)
# #table1 = pd.pivot_table(db1, values = 'Amount', index = ['bank','Account'], columns = ['bank_date'], aggfunc = 'count')

# # test 
# if not len(files_list) == 0:
#     table1_validation = True
#     table1 = pd.pivot_table(db1, values = 'Amount', index = ['bank','Account'], columns = ['bank_date'], aggfunc = 'count')
# else:
#     table1_validation = False
#     table1 = pd.DataFrame({'bank':['BOA'], 'Account':[488042589392], current_date:[0]}) # create empty table


# ________________________________________________________________________________________________________________________________________
# 53rd Bank ______________________________________________________________________________________________________________________________
source_folder = r'\\frwy\main\MXTJ\ACCT\Acct\Acct2021\Freeway\Banks Download\Fifth Third Bank'
files_list = glob.glob(source_folder + r'\*' + m + '.*' + d + '*.xls*')
#files_list = glob.glob(source_folder + r'\*MAY MTD 2022.xls*') # monthlys
for f in files_list:
    if not "~$" in f:
        print('adding: ' + f)
        db = pd.read_excel(f, header=None, names=['Bank ref','Account','bank_date','BAI Code','Description','Blank1','Amount','Blank2','ref1','ref2'],
         usecols=[0, 1, 2, 3, 4, 5, 6, 7, 8, 9])

        #db['bank_date'] = pd.to_datetime(db['bank_date']).dt.strftime('%m/%d/%Y') # format date as mm/dd/yyyy
        file_name = f.replace(source_folder + '\\' , "")  
        db['file_name'] = file_name
        db['bank'] = "53rd"
        #db2 = db2.append(db, ignore_index=True)
        db2 = pd.concat([db2, db], ignore_index=True)
        db = None

# format column date
#db2['bank_date']=db2['bank_date'].astype(str)

# Display data in pivot >>>>>>>>>>>> 
if not len(files_list) == 0:
    table2_validation = True
    table2 = pd.pivot_table(db2, values = 'Amount', index = ['bank','Account'], columns = ['bank_date'], aggfunc = 'count')
else:
    table2_validation = False
    table2 = pd.DataFrame({'bank':['53rd'], 'Account':[7934337234], current_date:[0]}) # create empty table
# ________________________________________________________________________________________________________________________________________
# the rest of Bank accounts___________________________________________________________________________________________________________________
# PROCESS:  Search for month input number in each file from  folder list(f_names). 
# f_names: list of folder to search files 
for f in f_names:
    c_path = os.path.join(dls_older, f)
    this_files = glob.glob(c_path + r'\*' + m + '.' + d + '*xls*')
    #this_files = glob.glob(c_path + r'\*MAY MTD 2022.xls*')
        #-new line-----------------------------------------------------#
    this_files_csv = glob.glob(c_path + r'\*' + m + '.' + d + '*csv*')
    #this_files_csv = glob.glob(c_path + r'\*MAY MTD 2022.csv*')
    this_files = this_files + this_files_csv
    #-end new line-------------------------------------------------#
    
    # for each folder. get files forselected month
    for f in this_files:
        if not "~$" in f:
            log_f.append(f)
            print('reading file: %s' % f)
            if not 'csv' in f:
                df = pd.read_excel(f)
            else:
                df = pd.read_csv(f)
            
            # save file name 
            file_name = f.replace(dls_older + '\\' , "")   
            s_file_name = file_name.split('\\') 

            # find date column in DataFrame 
            d1 = find_col_date_in_Headers(s_file_name, f, df)
            if len(d1) == 0:
                d1 = find_col_date_in_rows(s_file_name, f, df)
            
            # consolidate data in one table 
            #my_db = my_db.append(d1, ignore_index=True)
            my_db = pd.concat([my_db,d1], ignore_index=True)

            # dump info in placeholders 
            df = None
            d1 = None
            col_date_position = None

# Additional columns _______________

balance_table = pd.read_excel(open(trust_check_list, 'rb'), sheet_name='Acc', usecols='A:H') 
# set same type 
balance_table['last4'] = balance_table['last4'].astype(int)
my_db['last4']= my_db['last4'].astype(int)
my_db = pd.merge(my_db, balance_table[['last4','Accounts2']], on='last4', how='left') # last4 
my_db.rename(columns= {"Accounts2": "Account", "folder": "bank", "last4": "Amount"}, inplace = True) 

#my_db['bank_date']= my_db['bank_date'].astype(str)
table3 = pd.pivot_table(my_db, values = 'Amount', index = ['bank','Account'], columns = ['bank_date'], aggfunc = 'count')

# ______________________________________________________________________________________________________________________________#
# US Bank  _____________________________________________________________________________________________________________________#
source_folder = r'\\frwy\main\MXTJ\ACCT\Acct\Acct2021\Freeway\Banks Download\US Bank'
files_list = glob.glob(source_folder + r'\*' + m + '.' + d + '.*' + y + '.xls*')
#files_list2 = glob.glob(source_folder + r'\*' + m + '.*' + d + '*.csv')
#files_list = files_list1 + files_list2
   
# view list # 
for f in files_list:
#    print(f)
# Read data >>>>>>>>>>>>
    if not "~$" in f:
        print('adding: ' + f)
        # if '.csv' in f:
        #     db1 = pd.read_csv(f,header=None, 
        #         names= ['Type', 'Date', 'routing', 'Account', 'Name', 'Currency','BIA Code', 'Description', 'Deposit_Type', 'Amount', 'text','ref','text2'],
        #         usecols=[0,1,2,3,4,5,6,7,8,9,10,11,12])
        # else:
        db = pd.read_excel(f,header=None, 
        names= ['Type', 'Date', 'routing', 'Account', 'Name', 'Currency','BIA Code', 'Description', 'Deposit_Type', 'Amount', 'text','ref','text2'],
        usecols=[0,1,2,3,4,5,6,7,8,9,10,11,12])
            
        # filter non-blanks in column 'text'
        d1 = db[(db['Deposit_Type'] == 'Credit') | (db['Deposit_Type'] == 'Debit')]
        
        #d1 = pd.DataFrame()
        # Format Date
        date_list = []
        for d in d1['Date']:
                date_list.append(datetime.strptime(str(d), '%m%d%Y'))
            
        
        d1['bank_date'] = date_list
        #d1.loc[:,('bank_date')] = date_list
        #d1['bank_date'] = pd.to_datetime(d1['bank_date'], format='%m%d%Y')
        #d1['Amount'] = db['Amount']
        #d1['Account'] = db['Account']
        # add column with file name 
        file_name = f.replace(source_folder + '\\' , "")   
        d1['file_name'] = file_name
        #d1.loc[:,('file_name')] = file_name
        #d1['bank'] = "US Bank"    
        d1.loc[:,('bank')] = "US Bank"    
        # compilation
        #db3 = db3.append(d1, ignore_index=True)
        db3 = pd.concat([db3, d1], ignore_index=True)
        db = None
        d1 = None
#----------------------------------

# Display data >>>>>>>>>>>>
#db3['bank_date']=db3['bank_date'].astype(str)
table4 = pd.pivot_table(db3, values = 'Amount', index = ['bank','Account'], columns = ['bank_date'], aggfunc = 'count')

#_________________________________________________________________________________________________________________________
# WRAP : 
# Merge tables 

# test
table = pd.concat([table2, table3, table4])

# if table1_validation and table2_validation:
#     table = pd.concat([table1 , table2, table3, table4])
#     table = table.reset_index()
# elif not table1_validation and table2_validation:
#     table = pd.concat([table2, table3, table4])
#     table = table.reset_index()
# elif table1_validation and not table2_validation:
#     table = pd.concat([table1, table3, table4])
#     table = table.reset_index()
# else:    
#     table = pd.concat([table3, table4])

# save
timemark = datetime.today().strftime('%Y-%m-%d %H.%M.%S')
xl_file_name = r'\Uploads_complementary_' + timemark + '.xlsx'

with pd.ExcelWriter(destination_Folder + xl_file_name) as writer:
    table.to_excel(writer, sheet_name='Sheet1')
    #table3.to_excel(writer, sheet_name='Sheet2')
    #db.to_excel(writer, sheet_name='Sheet2')

#__________________________________________________________________________________________________________________________
# END: 
# open distination folder________________________________________________________________________________________________ 
d_path = os.path.realpath(destination_Folder)
os.startfile(d_path)
os.system('start excel.exe "' + destination_Folder + xl_file_name + '"')
