# -*- coding: utf-8 -*-
"""
Created on Wed Apr 29 15:05:14 2020

@author: limyuguang
"""

import pandas as pd

path_1 = r""
path_2 = r""

df1 = pd.read_excel(path_1, sheet_name = 0) #lookup table
df2 = pd.read_excel(path_2, sheet_name = 0) #source

df2.columns = range(df2.shape[1])
df2 = df2[6:]
df2 = df2.reset_index(drop=True)

for index, row in df2.iterrows():
    if("" in df2.loc[index, 0] or "" in df2.loc[index, 0] or "" in df2.loc[index, 0]):
        df2 = df2.drop(index)

df2 = df2.reset_index(drop=True)

df2[""] = ""
df2[""] = "" #aka downstream consistency check

for index, row in df2.iterrows():
    string = str(df2.loc[index, 16]) #value in downstream column
    if string == "":
        df2.loc[index, ""] = ""
    else:
        VD = string.split(",")
        for i in range(len(VD)):
            if("" in VD[i]):
                VD_editted = VD[i].replace(" ", "") #make sure there is no weird spacing in the cell
                VD_editted = VD_editted.replace("v1", "")
                VD_editted = VD_editted.replace("v2", "")
                df2.loc[index, ""] = VD_editted #Get from downstream column
                break
            else:
                df2.loc[index, ""] = ""
                
for index, row in df1.iterrows():
    df1.loc[index, ""] = df1.loc[index, ""].replace(" ", "") #make sure there is no weird spacing in the cell
    df1.loc[index, ""] = df1.loc[index, ""].replace(" ", "") #make sure there is no weird spacing in the cell
    
def checker(keyword):
    for index, row in df1.iterrows():
        if (keyword == ""):
            return "blank"    
        elif(df1.loc[index, ""].upper() == keyword):
            name = df1.loc[index, ""]
            return name
    
for index, row in df2.iterrows(): 
    keyword = df2.loc[index, ""].upper()
    name = checker(keyword)
    if(name.upper() == "BLANK"):
        df2.loc[index, ""] = "empty"    
    elif(name.upper() in str(df2.loc[index, 30]).upper()): #find the name in the column
        df2.loc[index, ""] = ""
    else:
        df2.loc[index, ""] = ""

df2[""] = ""

for index, row in df2.iterrows():
    upstream = str(df2.loc[index, 15])
    if ('' in upstream.upper()):
        df2.loc[index, ""] = "consistent"
    else:
        df2.loc[index, ""] = "not consistent"

df2[""] = ""
for index, row in df2.iterrows():
    if(df2.loc[index, ""] == "empty"):
        df2.loc[index, ""] = ""
    elif(df2.loc[index, ""] == "" and df2.loc[index, ""] == ""):
        df2.loc[index, ""] = ""
    elif(df2.loc[index, ""] == "" and df2.loc[index, ""] == ""):
        df2.loc[index, ""] = ""
    elif(df2.loc[index, ""] == "" and df2.loc[index, ""] == ""):
        df2.loc[index, ""] = ""
    elif(df2.loc[index, ""] == "" and df2.loc[index, ""] == ""):
        df2.loc[index, ""] = ""
    
df2.rename(columns={0:'',1:'',2:'',3:'',4:'',5:'',6:'',7:'',8:'',9:'',10:'',11:'',12:'',13:'',14:'',15:'',16:'',17:'',18:'',19:'',20:'',21:'',22:'',23:'',24:'',25:'',26:'',27:'',28:'',29:'',30:'',31:'',32:'',33:'',34:'',35:''},inplace=True)

df2.to_excel(r"", index=False)
