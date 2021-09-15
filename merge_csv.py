#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import os   
import time


# In[2]:


today_date = time.strftime('%Y%m%d') #the date we generate a new timestamp
today_date = int(today_date) + 2
today_date = str(today_date)
path_1 = r""

#this function is to read the latest datestamp merged file as a dataframe
def clone():
    files = os.listdir(path_1) 
    dates = [] #a list of dates from the Merged_CSV folder that is being read
    
    for j in files:
        if not("Latest" in j or "Temp" in j):
            j = int(j.split("-")[0])
            dates.append(j)
    latest_date = str(max(dates)) #only take csv with the latest date and clone from it    
    
    for f in files:
        if(latest_date in f):
            df = pd.read_csv(path_1 + r"\\" + f)
    
    return df 

#this function merges all the module files to be a single Latest-Merged.csv
def merge():
    path_2 = r""
    files = os.listdir(path_2)
    df = pd.DataFrame()
    for f in files:
        if((".xlsx" in f) and (not "WeeklyDiff" in f) and (not "~$" in f)):
            data = pd.read_excel(path_2 + r"\\" + f, sheet_name = "")
            df = df.append(data, sort=False)
    
    df['Actual'] = ''
    df['Status'] = ''
    df.insert(1, "Timestamp", today_date) #insert the date for the new df at column B
    df = df.reset_index(drop=True) #this is to make sure the indexing is correct. we need to reuse this df without saving
    
    return df

old_df = clone() #df of previous week           
new_df = merge() #df of current week


# In[3]:

#this function copy and paste the "Actual" and "Status" information from the previous week to the Latest-Merged.csv
def copy_actual(keyword):
    for index, row in old_df.iterrows():
        if(old_df.loc[index, ''] == keyword):
            actual = old_df.loc[index, 'Actual']
            return actual
        
def copy_status(keyword):
    for index, row in old_df.iterrows():
        if(old_df.loc[index, ''] == keyword):
            status = old_df.loc[index, '']
            return status    

for index, row in new_df.iterrows():
    keyword = new_df.loc[index, '']
    actual = copy_actual(keyword)
    status = copy_status(keyword)
    
    new_df.loc[index, ''] = actual       
    new_df.loc[index, ''] = status
        

new_df.to_csv(path_1 + r"\\" + "", index=False)

