# -*- coding: utf-8 -*-
"""

@author: limyuguang
"""

import os
import pandas as pd


#this is to combine all the FB together
path_1 = r""

#this will get the list of files in the folder
files = os.listdir(path_1)


df = pd.DataFrame() #empty dataframe

files_DA = [] #list that capture files with keyword in it

for f in files:
    if("" in f):
        files_DA.append(f)


for f in files_DA:
    data = pd.read_excel(path_1 + r"\\" + f, sheet_name = 0)
    df = df.append(data) #combine all the FB into 1 dataframe

df.to_excel(path_1 + r"\\" + "", index=False) #save the file in the same directory