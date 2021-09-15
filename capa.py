# -*- coding: utf-8 -*-
"""

@author: limyuguang
"""

import numpy as np
import pandas as pd

#data - extract the CAPA value (int)
data = np.genfromtxt(r'',delimiter=',')

max_rows = len(data) #46
max_cols = len(data[0]) #11


new_rows = (max_rows - 1) * 10 #exclude the header and mar-dec (10mths)
new_cols = 3 #FB, CAPA and MONTH
new_data = np.zeros([new_rows, new_cols])


k = 0

for i in range(1, max_rows):
    for j in range(1, max_cols):
        new_data[k][2] = data[i][j] #inserting the CAPA value into the new_data from the original data
        k = k + 1 #k is a counter for the rows in new_data
        j = j + 1 #j is the counter for the columns in data
    i = i + 1



#data2 - extract the FB (string)
data2 = np.genfromtxt(r'', delimiter=',', dtype=str)
fb_list = [] #list to store all the FB names

for i in range(1, max_rows):
        fb_list.append(data2[i][0])
        i = i + 1

month_list = ['March','April','May','June','July','August',
              'September','October','November', 'December'] #all the months in the file

df = pd.DataFrame(new_data)

count = 0
for index, row in df.iterrows():
    if(index % 10 == 0): #10 because there is only 10 months
        count = count + 1
    df.iloc[index, 0] = fb_list[count-1]

count = 0
for index, row in df.iterrows():
    if(count == 10):
        count = 0
    df.iloc[index, 1] = month_list[count]
    count = count + 1

df["Stage"] = 'Default'

df1 = pd.read_excel(r"")

milestone_list = []
milestone_list.append(df1.loc[0, 'Date'].month_name()) #bronze
milestone_list.append(df1.loc[1, 'Date'].month_name()) #silver
milestone_list.append(df1.loc[2, 'Date'].month_name()) #gold
milestone_list.append(df1.loc[3, 'Date'].month_name()) #titanium

#dictionary to change the month into integer
month_dict = {'March' : 3,
              'April' : 4,
              'May' : 5,
              'June' : 6,
              'July' : 7,
              'August' : 8,
              'September' : 9,
              'October' : 10,
              'November' : 11,
              'December' : 12
              }

for index, row in df.iterrows():
    if (month_dict[df.iloc[index, 1]] >= month_dict[milestone_list[0]]) and (month_dict[df.iloc[index, 1]] <= month_dict[milestone_list[1]]):
        df.loc[index, 'Stage'] = 'Bronze'
    elif (month_dict[df.iloc[index, 1]] > month_dict[milestone_list[1]]) and (month_dict[df.iloc[index, 1]] <= month_dict[milestone_list[2]]):
        df.loc[index, 'Stage'] = 'Silver'
    elif (month_dict[df.iloc[index, 1]] > month_dict[milestone_list[2]]) and (month_dict[df.iloc[index, 1]] <= month_dict[milestone_list[3]]):
        df.loc[index, 'Stage'] = 'Gold'
    elif (month_dict[df.iloc[index, 1]] > month_dict[milestone_list[3]]) and (month_dict[df.iloc[index, 1]] <= month_dict['December']):
        df.loc[index, 'Stage'] = 'Titanium'
    else:
        df.loc[index, 'Stage'] = 'Unknown'

df.rename(columns={0 : '',
                   1 : '',
                   2 : '',
                   }, inplace = True) #rename the columns

df.to_csv(r'', index = False)



