# -*- coding: utf-8 -*-
"""
Important notes:

This code will generate a folder to store name-data excel sheet on Desktop, 
the folder is named as 'print_job', the excel sheet is named as 'name_data',

To use this code, need to install the following modules ---
pandas + langid + openpyxl + os, 
they are easy to install, installations can be found in Google.

The code takes the excel file path, which contains the raw name data, as input.

Path has to be in a certain form: (remember the 'r' in the front)
    file = r'excel path'

---ver. 1.0 Dsh 19/07/22
    The output excel contains three more columns, which are '中文', 'Engligh' and 'revisit'.
"""

#%%
import pandas as pd
import langid
from openpyxl import load_workbook
import os

########################### Useful functions##################################
def space_counter(name):
    '''
    Parameters
    ----------
    name : str
        Name in either chinese character or english

    Returns
    -------
    count : float
        Number of 'space' in the name

    '''
    count = 0
    for i in name:
        if (i.isspace()) == True:
            count += 1
    return count 


##############################################################################
###########################  Editions from here  #############################


# Indicate excel sheet 'path'!!!  Remeber the 'r'
file = r'C:\Users\work\Desktop\test.xlsx'

# Sheet format, 0 = headers on, None = No header 
header_ = 0   


#########################        End here        #############################
##############################################################################

# Creating a folder for printing on Desktop
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
newpath = desktop + '\print_job' 
if not os.path.exists(newpath):
    os.makedirs(newpath)
    
# Save a copy to a different file
new = r'C:\Users\work\Desktop\print_job\name_data.xlsx'
    
########################  Preparing the dataset  #############################

# Read the sheet and get the header of the name data
df0 = pd.read_excel(file, header=header_)
col_name = df0.columns[0]

# Preparing list for tesing and storing 
en=[]
zh=[]
re=[]

# Open workbook
workbook = load_workbook(filename=file)
sheet = workbook.active

# Modify the desired cell
sheet["C1"] = "中文"
sheet["D1"] = "English"
sheet["E1"] = "Revisit"

# Save the change in pandas
workbook.save(new)

# Read the new sheet
df = pd.read_excel(new, header=header_)

# Storing counters
num_zh = 0
num_en = 0
num_re = 0

# Allocation algorithm 
for i in range(len(df0[col_name])):
    a = df0[col_name][i]
    #print(langid.classify(i)[0])
    if langid.classify(a)[0] == 'zh':
        zh.append(a)
        df[df.columns[2]][num_zh] = a
        num_zh += 1
        #print(i,num)
    else :
        if space_counter(a) < 2 :
            fullname = a.split(' ')
            if len(fullname[0]) < 11 :
                en.append(fullname[0])
                df[df.columns[3]][num_en] = fullname[0]
                num_en += 1
                #print(num_en,fullname)
            else:
                re.append(a)
                df[df.columns[4]][num_re] = a
                num_re+=1
                #print('long name') 
        else:
           re.append(a)
           df[df.columns[4]][num_re] = a
           num_re+=1
           
# Test, Save and Warning           
if len(en) + len(zh) + len(re) == len(df0[col_name]):
    df.to_excel(new)
    print('\n\n\n New file is stored at \n',new,' \n Check the -Revisit- column \n\n\n\n\n')
else:
    raise Warning("Data length mismatch, check input data and restart the code")

