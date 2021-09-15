#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import os


# In[2]:


def merge():
    path_7 = r""
    files = os.listdir(path_7)
    df4 = pd.DataFrame()
    for f in files:
        if not("Latest-" in f):
            data = pd.read_csv(path_7 + r"\\" + f)
            df4 = df4.append(data, sort=False)
    df4.to_csv(r"",index=False)
merge()

