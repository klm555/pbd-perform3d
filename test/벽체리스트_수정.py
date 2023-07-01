# -*- coding: utf-8 -*-
"""
Created on Fri May 26 10:30:27 2023

@author: hwlee
"""

# -*- coding: utf-8 -*-
"""
Created on Fri May 26 08:03:25 2023

@author: hwlee
"""
import pandas as pd
import numpy as np
import win32com.client
import pythoncom

input_file_name = r'K:\2105-이형우\성능기반 내진설계\KHSM\107\KHSM_107_Data Conversion_Ver.1.3M_내진상세_변경.xlsx'
input_sheet_name = 'Results_Wall_보강 (2)'

output_file_name = r'K:\1110-김학범\김해신문1지구 A17-1BL\내진성능설계\230525-2차_보강안(102D,105D,106D,107D,108D)\벽체_배근(전동)\230526-GIMHAE-107D-WALL LIST-0004.xlsx'
output_sheet_name = 'BatchWall'

#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%

data = pd.read_excel(input_file_name, sheet_name=input_sheet_name, skiprows=[0,1,2]
                     , usecols=[0,5,6,7,8,9])
data.columns = ['Name', 'V.Rebar Type', 'V.Rebar Spacing', 'V.Rebar EA'
                , 'H.Rebar Type', 'H.Rebar Spacing']

data = data.replace(np.nan, '', regex=True)

new_v_rebar = []
for i, j in zip(data['V.Rebar Type'], data['V.Rebar EA']):
    if j == '':
        new_v_rebar.append(i)
    else:
        new_v_rebar.append(str(int(j)) + '-' + i)
        
data['V.Rebar Type'] = new_v_rebar


excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize())
excel.Visible = True

wb = excel.Workbooks.Open(output_file_name)
ws = wb.Sheets(output_sheet_name)


output_name = ws.Range('A%s:A%s' %(1, 5000)).Value
output_name_df = pd.DataFrame(output_name)
output_name_df_sliced = output_name_df[output_name_df.iloc[:,0].isna() == False]
output_name_df_sliced.reset_index(inplace=True, drop=False)
output_name_df_sliced.columns = ['index', 'original value']


# Story가 있는 셀의 인덱스 찾기
story_index = output_name_df_sliced[output_name_df_sliced['original value'] == 'Story'].index
wall_list = []
for i in story_index:
    wall_name = output_name_df_sliced.iloc[i-1, 1]
    wall_list.append(wall_name)
wall_list = pd.DataFrame(wall_list)

wall_list.columns = ['added_col']


output_name_df_sliced = pd.merge(output_name_df_sliced, wall_list, how='left', left_on='original value', right_on='added_col')
output_name_df_sliced.ffill(axis = 0, inplace=True)

output_name_df_sliced['new name'] = output_name_df_sliced['added_col'] + '_1_' + output_name_df_sliced['original value']



# merge
output_name_df_sliced = pd.merge(output_name_df_sliced, data, how='left', left_on='new name', right_on='Name')
output_name_df_sliced = output_name_df_sliced.dropna(subset=['Name'])



for idx, row in output_name_df_sliced.iterrows():
    ws.Range('E%s:E%s' %(str(row['index']+1), str(row['index']+1))).Value = row['V.Rebar Type']
    ws.Range('F%s:F%s' %(str(row['index']+1), str(row['index']+1))).Value = row['V.Rebar Spacing']
    ws.Range('H%s:H%s' %(str(row['index']+1), str(row['index']+1))).Value = row['H.Rebar Type']
    ws.Range('I%s:I%s' %(str(row['index']+1), str(row['index']+1))).Value = row['H.Rebar Spacing']
    
wb.Save()  