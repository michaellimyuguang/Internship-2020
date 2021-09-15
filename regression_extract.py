#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import os


# In[2]:


def merge():
    df = pd.DataFrame()
    path = r""
    files = os.listdir(path)

    for file in files:
        date = file
        date_parts = date.split('-')
        date = str(20) + date_parts[0] + date_parts[1] + date_parts[2]
        
        data = pd.read_csv(path + r"\\" + file + r"\\" + r"")
        data.insert(1, "", date)
        
        df = df.append(data, sort=False)
        
    df.to_csv(r"", index=False)
    
merge()

def merge_2():
    df = pd.read_csv(r"")
    path = r""
    
    files = os.listdir(path)
    
    for file in files:
        date = file
        date_parts = date.split('-')
        date = str(20) + date_parts[0] + date_parts[1] + date_parts[2]    
        data = pd.read_csv(path + r"\\" + file + r"\\" + r"")
        print(data)
        data.insert(1, "Timestamp", date)
        
    df = df.append(data, sort=False)
    
    df.to_csv(r"", index=False)

merge_2()
