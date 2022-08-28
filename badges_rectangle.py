# -*- coding: utf-8 -*-
"""

Important notes:
    
This code will generate a html file for printing badges. 

Before using this code, 'name_data.xlsx' is needed to be revisit manualy, change the name format
in the revisit column!!!

The background images for badges and name are needed to be saved in 'print_job' folder!!!

Formats and parameters of the badges can be modified in parameters region, watch out the unit!!!
     

---ver. 1.0 Dsh 19/07/22
    The output html contains 10 badges on a single A4 sheet.

---ver. 2.0 Dsh 23/07/22
    'Revisit' name can now be added to the end of "English" column


"""

import pandas as pd
import os


##############################################################################
#####################  Parameters can be modified  ###########################
 
# Badge background and dimensions in cm !!!!
badge_bg = "'badge_bg.png'"
badge_width = 8.71
badge_height = 5.51 

# Name box background and dimensions in cm !!!!!!
name_bg = "'name_bg.png'"
name_width = 7.8
name_height = 3

# Name box relative postion to the badge, in cm !!!!!
name_margin_top = 2

# Font of the name
font = 'Heiti'

##############################################################################
##############################################################################

# Extract desktop path for storing 
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
newpath = desktop + '\print_job' 

# Read the updated data
file = r'%s\name_data.xlsx' % (newpath) 

# Sheet format, 0 = headers on, None = No header 
header_ = 0 

# Read the sheet and get the header of the name data
df = pd.read_excel(file, header=header_)

#Preparing name data set
zh =[]
en =[]

for i in range(len(df['中文'])) :
    if type(df['中文'][i]) is str:
        zh.append(df['中文'][i]) 

for i in range(len(df['English'])) :
    if type(df['English'][i]) is str:
        en.append(df['English'][i]) 

number = len(en)

for i in range(len(df['Revisit'])) :
    if type(df['Revisit'][i]) is str:
        en.append(df['Revisit'][i]) 
        df['English'][number] = df['Revisit'][i]
        number += 1

name_data = zh + en
df.to_excel(file)

# Html algorithm and design
head = '''
<html lang="en">
<head>
    <title>Print</title>
    <style>
       * {padding:0; margin: 0;}
       @page { margin: 2cm }
       @media print  
       {
       div {
           page-break-inside: avoid;
           }
       
       }

       .page {
           background-color: transparent;
           margin-top: 0.5cm;
           margin-left: 0.5cm;
 /* A4 paper dimension */
           width: 21cm;
           height: 29.5cm;
           border-style: solid;
           border-width: 1px;
           border-color: transparent;
           padding-left: 50px;
       }
       .badge{
           background-image: url(%s); 
           background-repeat: no-repeat;
           background-size: %fcm %fcm;
           /*background-color: yellow;*/
           width: %fcm;
           height: %fcm;
           border-style: solid;
           border-color: transparent;
           border-width: 0.5px;
           float: left;
       }
       .name{
           background-image: url(%s);
           background-repeat: no-repeat;
           background-size: %fcm %fcm;
           width:%fcm;
           height: %fcm;
           margin-top: %fcm;
           margin-left: auto;
           margin-right: auto;
       }
       .centre{
           padding-top: 0.3cm;
       }
       p {
           font-family: %s;
           font-size: 45pt;
           text-align: center;;
       }
   </style>
</head>
'''% (badge_bg, badge_width,badge_height,badge_width,badge_height,\
       name_bg,  name_width,name_height,name_width,name_height,\
    name_margin_top,font)
    
c = '''
<div class = badge>
    <div class = name>
        <div class = centre>
            <p>%s</p>
        </div>
    </div>
</div>
''' 

# Generate badges with different names
Cc = str(c)
Names = c%(name_data[0])
for i in range(1,len(name_data)):
    #print(zh[i])
    d = str(Cc)
    D = d%(name_data[i])
    #print(D)
    Names += D

body = '''
<body>
    <div class = page>
        %s
    </div>
</body>
</html>
''' % (Names)


# Create html file
F = open(r'%s\badge.html' %(newpath),"w",encoding="utf-8")
F.write(head + body)
              
# Saving the data into the HTML file
F.close()

    
    
