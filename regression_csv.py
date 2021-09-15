#!/usr/bin/env python
# coding: utf-8

# In[11]:


import pandas as pd
import os
import datetime


# In[12]:


dates_list = [] #this list stores all the relevant dates for regression
week_files = [] #this list stores all the folder files for regression

#this function is used to find all the relevant dates for a particular calender week
def find_date():
    for i in range(7):
        date = datetime.datetime.today() - datetime.timedelta(days=(6-i))
        date = date.strftime('%Y%m%d')
        dates_list.append(date)

find_date()

#this function is used to find the folder files that has the matching date as stated in the previous function
def find_file():
    path_3 = r""
    files = sorted(os.listdir(path_3))
    for i in range(len(dates_list)):
        for j in files:
            if(dates_list[i] in j):
                week_files.append(j)
find_file()

print(dates_list)
print(week_files)

#this function is used to find the latest YYYYMMDD-Merged.csv file
def find_latest_date():
    path_5 = r""
    files = os.listdir(path_5) 
    dates = []
    for j in files:
        if not("Latest" in j):
            j = int(j.split("-")[0])
            dates.append(j)
    
    latest_date = str(max(dates))
    
    return latest_date

latest = find_latest_date()

df2 = pd.read_csv(r"" + r"\\" + latest + "")

#this function is used to extract the actual date
def copy_actual(keyword, df1, actual_date):
    for index, row in df1.iterrows():
        if(df1.loc[index, 'INF'] == keyword):
            if(df1.loc[index, 'Result'] == 'FAILED'):
                return 1
            else:
                return actual_date
            
    return 2

#this function is used to extract the status        
def copy_status(keyword, df1, actual_date):
    for index, row in df1.iterrows():
        if(df1.loc[index, 'INF'] == keyword):
            if(df1.loc[index, 'Result'] == 'FAILED'):
                return df1.loc[index, 'Result']
            else:
                return df1.loc[index, 'Result']
    return 2

#this function is used to execute the copy_actual function and copy_status function defined earlier on  
def regression(df1, actual_date):
    for index, row in df2.iterrows():
        keyword = df2.loc[index, '']
        actual = copy_actual(keyword, df1, actual_date)
        if actual == 2:
            continue
        elif actual == 1:
            df2.loc[index, 'Actual'] = ""
        else:
            actual = str(actual)
            actual = datetime.datetime.strptime(actual, '%Y%m%d').strftime('%m/%d/%Y')
            df2.loc[index, 'Actual'] = actual



    for index, row in df2.iterrows():
        keyword = df2.loc[index, '']
        status = copy_status(keyword, df1, actual_date)
    
        if status == 2:
            continue
        else:
            df2.loc[index, ''] = status


# In[13]:

#this is the main function used to execute the previous functions
def open_file():
    path_4 = r""
    
    for k in week_files:
        folder = path_4 + r"\\" + k
        check_csv = os.listdir(folder)
        for w in check_csv:
            if("" in w):
                actual_date = int(k.split('_')[0])
                df1 = pd.read_csv(folder + r"\\" + r"")
                regression(df1, actual_date)
open_file()



timestamp = str(df2.loc[0, 'Timestamp'])

year = int(timestamp[0:4])
month = int(timestamp[4:6])
day = int(timestamp[6:8])


my_date = datetime.date(year, month, day)
year = str(my_date.isocalendar()[0])
cw = str(my_date.isocalendar()[1])
cw = year + cw

def generate_module_plan(x):
    if int(x) <= int(cw):
        return 1
    else:
        return 0
    
df2[''] = df2[''].apply(generate_module_plan)



df2.to_csv(r"" + r"\\" + latest + "", index=False)


# In[14]:

#Pattern Logic
selected_column = ['', '', '', '', '', '', '', '', '']
path_8 = r""
df3 = df2[selected_column]
df3.insert(6, "", "")

def calender_week(x):
    if(type(x) == str and x != ''):
        date = datetime.datetime.strptime(x, "%m/%d/%Y")
        year = str(date.isocalendar()[0])
        week = str(date.isocalendar()[1])
        return year + week

df3['Actual'] = df3['Actual'].apply(calender_week)

df3.rename(columns={'':'','': '','':'','':''},inplace=True)


path_8 = r""
df5 = pd.read_excel(path_8, sheet_name = '')
df6 = pd.read_excel(path_8, sheet_name = '')
df7 = pd.read_excel(path_8, sheet_name = '')

with pd.ExcelWriter(r"") as writer:
    df3.to_excel(writer, sheet_name = '',index=False)
    df5.to_excel(writer, sheet_name = '',index=False)
    df6.to_excel(writer, sheet_name = '',index=False)
    df7.to_excel(writer, sheet_name = '',index=False)

