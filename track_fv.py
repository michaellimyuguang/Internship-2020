# -*- coding: utf-8 -*-
"""
Created on Mon Mar 30 10:30:52 2020

@author: limyuguang
"""

import os
import pandas as pd
import datetime

def generate_input_folder():
    my_date = datetime.date.today()

    #print(my_date)
    #week = my_date.isocalendar()[1]
    
    #week = week - 1
    #pw = week - 18
    #print("pw:", pw)
    #print("week:", week)
    
    #friday
    if(my_date.isocalendar()[2] == 5):
        week = my_date.isocalendar()[1]
        pw = week - 18
        print("week:", week)
        print("pw:", pw)
    #monday, tuesday, wednesday
    else:
        week = my_date.isocalendar()[1] - 1
        pw = week - 18
        print("week:", week)
        print("pw:", pw)
        
    input_folder = "PW" + str(pw) + "_" + "CW" + str(week)

    return input_folder

input_folder = generate_input_folder()
path_1 = r""
path_1 = path_1 + r"\\" + input_folder
vt_names = []
df3 = pd.DataFrame()
matching_names = []

count = 0 #counter for the number of matching files

def file_names():
    files = os.listdir(path_1)
    for f in files:
        if not("~$" in f):
            vt_names.append(f)

file_names() #populate all the excel file names into vt_names

print("No. of files in " + input_folder + ": " + str(len(vt_names)))

def extract_pw(dp_ss, name, device_name, df1, keyword):
    
    list_type = []
    list_name = []
    list_device_name = []
    list_element = []
    
    temp_list_2 = []
    temp_list_3 = []
    
    for index, row in df1.iterrows():            
        if("" == str(df1.loc[index, df1.columns[0]])):
            temp_list_0 = row
            temp_list_0 = temp_list_0.values.tolist()
            
    for index, row in df1.iterrows():
        if("" == str(df1.loc[index, df1.columns[0]])):
            temp_list = row
            temp_list = temp_list.values.tolist()
            
        elif("" == str(df1.loc[index, df1.columns[0]])):
            temp_list = row
            temp_list = temp_list.values.tolist()
             
    for index, row in df1.iterrows():
        if(name != ""):
            if("" == str(df1.loc[index, df1.columns[0]])):            
                temp_list_2 = row
                temp_list_2 = temp_list_2.values.tolist()
                del temp_list_2[0]
        else:
            temp_list_2 = []
            
    if(name == ""):
        for index, row in df1.iterrows():
            if("" == str(df1.loc[index, df1.columns[0]])):            
                temp_list_3 = row
                temp_list_3 = temp_list_3.values.tolist()    
                del temp_list_3[0]
            
    del temp_list[0] #remove Average PW Target in the list
    del temp_list_0[0] #remove Project week (PW)
    
    loop = len(temp_list)
    
    if(len(temp_list_2) == 0):
        for k in range(loop):
            temp_list_2.append(" ")
    
    if(len(temp_list_3) == 0):
        for m in range(loop):
            temp_list_3.append(" ")        
                                        
        
    for i in range(loop):
        list_type.append(dp_ss)
        list_name.append(name)
        list_device_name.append(device_name)
        list_element.append(keyword)
        
    df2 = pd.DataFrame({"" : temp_list_0, "" : list_type, "": list_name, "" : list_device_name, "" : temp_list, "" : temp_list_2, "" : temp_list_3, "" : list_element})

    return df2

df4 = pd.read_excel(r"", sheet_name = 1)
df4[''] = ''

def keyword_checker(keyword):
    for index, row in df4.iterrows(): 
        if(keyword.upper() == str(df4.loc[index, df4.columns[5]]).upper()):             
            df4.loc[index, ''] = "Yes"
            return 1
          
for i in range(len(vt_names)):
    if(".xlsx" in vt_names[i]):    
        if(vt_names[i].count("_") == 4):
        
            df1 = pd.read_excel(path_1 + r"\\" + vt_names[i], sheet_name = 0)
            string = vt_names[i].split("_")
            dp_ss = string[2]
            name = string[3]
            keyword = name
        
            if(keyword_checker(keyword)):
                matching_names.append(vt_names[i])
                count = count + 1
                print(vt_names[i])
                device_name = string[4]
                device_name = device_name.split(".")[0]
                df2 = extract_pw(dp_ss, name, device_name, df1, keyword)
                df3 = df3.append(df2)
        
#        elif("" in vt_names[i]):
#            print(vt_names[i])
#            df1 = pd.read_excel(path_1 + r"\\" + vt_names[i], sheet_name = 0)
#            string = vt_names[i].split("_")
#            dp_ss = string[2]
#            name = dp_ss
#            keyword = dp_ss
#            device_names = string[2]
#            df2 = extract_pw(dp_ss, name, device_names, df1, keyword)
#            df3 = df3.append(df2)
            
        elif ("" in vt_names[i]):
            df1 = pd.read_excel(path_1 + r"\\" + vt_names[i], sheet_name = 0)
            string = vt_names[i].split("_")
            dp_ss = ""
            name = string[2].split(".")[0]
            device_name = ""
            keyword = name
            df2 = extract_pw(dp_ss, name, device_name, df1, keyword)
            df3 = df3.append(df2)
        
        else:

            df1 = pd.read_excel(path_1 + r"\\" + vt_names[i], sheet_name = 0)
            name = vt_names[i].split("_")
            string = vt_names[i].split("_")
        

        
            del name[0]
            del name[0]
            del name[0]
  
            keyword_list = name
            del keyword_list[len(keyword_list) - 1]
            keyword = "_"
            keyword = keyword.join(keyword_list)
        
            if(keyword_checker(keyword)):
                matching_names.append(vt_names[i])
                count = count + 1
                print(vt_names[i])                
                dp_ss = string[2]
                device_name = string[len(string) - 1]
                device_name = device_name.split(".")[0]

                for w in range(len(name)):
                    df2 = extract_pw(dp_ss, name[w], device_name, df1, keyword)
                    df3 = df3.append(df2)

df4[''] = ''

def cw_checker():
    for index, row in df1.iterrows():
        if(df1.loc[0, df1.columns[1]] == 19 or df1.loc[0, df1.columns[2]] == 19):
            return "Yes"
        else:
            return "No"
        
for index, row in df4.iterrows():
    if(df4.loc[index, ''] == "Yes"):
        df4.loc[index, ''] = cw_checker()

    
df3[""] = input_folder

print("" + input_folder + ": " + str(count))

  
df3.to_excel(r"" + r"\\" + input_folder + "", index=False)
df4.to_excel(r"" + r"\\" + "", index=False)

df6 = pd.DataFrame({"" : vt_names})
df6[""] = ""
df6[""] = input_folder

for index, row in df6.iterrows():
    for a in range(len(matching_names)):
        if(df6.loc[index, df6.columns[0]] == matching_names[a]):
            df6.loc[index, df6.columns[1]] = "Yes"
    
df6.to_excel(r"" + r"\\" + input_folder + "", index=False)

######################################################################################################################################

output_folder = r""
overview_names = []

def pw_files():
    files = os.listdir(output_folder)
    for f in files:
        if not( "~$" in f):
            if("_overview" in f):
                overview_names.append(f)

pw_files()                
df5 = pd.DataFrame()

print(overview_names)

for j in range(len(overview_names)):
    data = pd.read_excel(output_folder + r"\\" + overview_names[j], sheet_name = 0)
    df5 = df5.append(data, sort=False)

df5.to_excel(output_folder + r"\\" + "", index=False)

data_names = []

def data_files():
    files = os.listdir(output_folder)
    for f in files:
        if not( "~$" in f):
            if("_data_info" in f):
                data_names.append(f)

data_files()                
df7 = pd.DataFrame()

print(data_names)

for j in range(len(data_names)):
    data = pd.read_excel(output_folder + r"\\" + data_names[j], sheet_name = 0)
    df7 = df7.append(data, sort=False)

df7.to_excel(output_folder + r"\\" + "", index=False)


my_date = datetime.date.today()
if(my_date.isocalendar()[2] == 5):
    week = my_date.isocalendar()[1]
    pw = week - 18
    print("week:", week)
    print("pw:", pw)
else:
    week = my_date.isocalendar()[1] - 1
    pw = week - 18
    print("week:", week)
    print("pw:", pw)

df3 = pd.read_excel(r"" + r"\\" + input_folder + "_overview.xlsx")
df4 = pd.read_excel(r"")

df3[""] = df3[""].fillna('a')

def extract_checker(element):
    for index, row in df3.iterrows():
        if(df3.loc[index, ""] == element and df3.loc[index, ""] == float(pw)):
            if(df3.loc[index, ""] != 'a'):
                return "SUCCESS"

df4[""] = ""
for index, row in df4.iterrows():
    element = df4.loc[index, df4.columns[5]]
    df4.loc[index, ""] = extract_checker(element)

df4.to_excel(r"", index=False)

df8 = pd.read_excel(r"")

df8['deviation'] = df8[''] - df8['']

df8.to_excel(r"",index=False)