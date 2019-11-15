import pandas as pd
import numpy as np
import os
import xlrd
from xlrd.formatting import Font
from os import listdir

def converter (file):

    sheets = pd.ExcelFile(file)
    sheet_list = sheets.sheet_names
    doc_names = [name.replace(" ","") for name in sheet_list]

    for c,  sheet_name in enumerate( sheet_list ):
        if c == 0:
            print('first page')
        else:
            try:
                os.mkdir(str(file)+'_folder')
            except: 
                print('-----------------------')
            struck_outs = struck_out_finder(file,c,1)
            conf_list = struck_out_finder(file,c,2)

            print(sheet_name)
            dfs = data_extractor(file, sheet_name, struck_outs, conf_list)
            try:
                dfs.to_csv(path_or_buf=file+'_folder/'+str(doc_names[c]) + '.csv')
            except: print("Nothing Saved")

def struck_out_finder(file, sheet_no, conf):
    book = xlrd.open_workbook(file, formatting_info=True)
    sheets = book.sheet_names()

    sheet = book.sheet_by_name(sheets[sheet_no])
    struck_out = []
    try:
        for row in range(0,sheet.nrows):
            font = book.font_list
            cell_val = sheet.cell_value(row,conf)
            cell_xf = book.xf_list[sheet.cell_xf_index(row,conf)]
            struck_out.append([ cell_val, font[cell_xf.font_index].struck_out ])
    except:
        print("Out of range Error")
    try : 
        del struck_out[0]
        return struck_out
    except:
        print("Out of range error")
    


def data_extractor(file, sheet_name, strike_list, conf_list):

    dfs = pd.read_excel(file, sheet_name= sheet_name)
    df = pd.DataFrame(dfs)

    if len(df.columns) < 5:
        return print("empty sheet")

    try:
        rows = df[df.columns[0]].str.contains('etwork')
    except:
        return print("No IP Adresses Found")

    for c, pair in enumerate( strike_list ):
        if pair[1] == 1:
            try:
                df = df.drop(c)
            except:
                print("nothing deleted")
        elif conf_list[c][1] == 1:
            try:
                df = df.drop(c)   
            except:
                print("Nothing deleted")
  
    df = df.reset_index(drop= True)
    DHCPs2 = []
    DHCPs = df[df.columns[1]].str.contains('DHCP')
    for c , DHCP in enumerate( DHCPs ):
        if DHCP == True:
            DHCPs2.append(c)

    counter = 0
    try:
        diff = DHCPs2[-1]-DHCPs2[0]
        for DHCP in range(diff+1):
            df = df.drop(DHCPs2[0] + counter)
            counter += 1

    except: print("No DHCP info found")


    for c , row in enumerate( rows ):
        if row == True:
            header_line = c
    try: 
        df.columns = df.loc[header_line, :]
        local_host_data = df.loc[(header_line + 2) :, :]
    except:
        return print("Empty Sheet")

    if len(local_host_data.columns) < 8:
        local_host_data["C"] = np.nan
        local_host_data["D"] = np.nan
        local_host_data["E"] = np.nan

    data = local_host_data.loc[:,[local_host_data.columns[1],local_host_data.columns[0],local_host_data.columns[2],local_host_data.columns[4],local_host_data.columns[5],local_host_data.columns[6],local_host_data.columns[7]]]
    try:
        data.columns = ['name','address','description','user', 'comments','null', 'CTASK','null2', 'CTASK2']
    except:
        try:
            data.columns = ['name','address','description','user', 'comments', 'CTASK', 'CTASK2']
        except: 
            return "WRONG COLUMNS--------------" + str(len(data.columns))
        


    names = data['name']
    ips = data['address']
    ips2 = ips
    nulls = names.isnull()
    names = [name for name in names]
    try:
        ips2 = [ip.replace('.','-') for ip in ips2]
        ips2 = [ 'Host-'+ip for ip in ips2]
        for a, null in enumerate( nulls ):
            if (null == True) or (len(names[a]) < 3):
                names[a] = ips2[a]
    except:
        try:
            ips2 = [str(ip) for ip in ips2]
            ips2 = [ip.replace('.','-') for ip in ips2]
            ips2 = [ 'Host-'+ip for ip in ips2]
            for a, null in enumerate( nulls ):
                if (null == True) or (len(names[a]) < 3):
                    names[a] = ips2[a]
        except: 
            print("ERROR")
    

    

    CTASK = data['CTASK']
    CTASK = [t for t in CTASK]
    C_ALT = data['CTASK2']
    C_NULL = C_ALT.isnull()
    C_ALT = [c for c in C_ALT]
    C_NULL = [c for c in C_NULL]

    for c , task in enumerate(C_NULL):
        if task == True:
            CTASK = C_ALT

    data_final = {'name':names,'address':ips,'descriptions':data['description'],'user':data['user'],'comments':data['comments'], 'CTASK':CTASK}
    data_final_df = pd.DataFrame(data_final)

    # inv_chars = data_final_df[data_final_df.columns[1]].str.contains('"')
    # for c, inv in enumerate( inv_chars):
    #     if inv == True:
    #         print(str(c) + " GOT ONE!")
    #         data_final_df = data_final_df.drop([c-2])

    data_final_df['Header-hostrecord'] = np.nan
    data_final_df['Header-hostrecord'] = ['hostrecord' for i in data_final_df['Header-hostrecord']]
    data_final_df = data_final_df.set_index('Header-hostrecord')
    return data_final_df

# A few of the names have invalid characters. The only legal character are a-z, A-Z, 0-09, 
# and '-'. Any other characters in ther name field should be removed. 
# If a name consists of only invalid characters, it should be a Host-x-x-x-x entry instead.
# Bodine/10.16.0.x Troy sheet/ row 239 is a DHCP address and should be removed
# Bodine/10.16.0.x Troy sheet/ row 32 is only '??'.  It should be a Host-x-x-x-x instead.
# Bodine/10.16.0.x TroySwitchMgmt sheet/ row 18 has no Name, just spaces. 
# It should be a Host-x-x-x-x instead.
# For the header and the first column, the value is word is 'hostrecord' (no space). 

# file_names = listdir("T_Data_Raw")
# for file_name in file_names:
#     converter("T_data_raw/" + file_name)




converter('T_data_raw/Bodine.xls')