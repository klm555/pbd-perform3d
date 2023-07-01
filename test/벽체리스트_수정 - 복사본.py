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

input_file_name = r'K:\2105-이형우\성능기반 내진설계\KHSM\105\KHSM_105_Data Conversion_Ver.1.3M_내진상세_변경.xlsx'
input_sheet_name = 'Results_Wall_보강 (7)'

output_file_name = r'K:\1110-김학범\김해신문1지구 A17-1BL\내진성능설계\230525-2차_보강안(102D,105D,106D,107D,108D)\벽체_배근(전동)\230526-GIMHAE-105D-WALL LIST-0004.xlsx'
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
    ws.Range('E%s:E%s' %(str(row['index']+1), str(row['index']+1))).Value = [output_name_df_sliced.iloc[idx,[5]]]
    ws.Range('F%s:F%s' %(str(row['index']+1), str(row['index']+1))).Value = [output_name_df_sliced.iloc[idx,[6]]]
    ws.Range('H%s:H%s' %(str(row['index']+1), str(row['index']+1))).Value = [output_name_df_sliced.iloc[idx,[8]]]
    ws.Range('I%s:I%s' %(str(row['index']+1), str(row['index']+1))).Value = [output_name_df_sliced.iloc[idx,[9]]]  
        
wb.Save()  






#%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

# = list(SF_output.iloc[:,[10,11,12,13,14,15,16,17,18]].itertuples(index=False, name=None))

# startrow = 5

# while True:
#     # 엑셀 읽기
#     # H. Rebar 정보 읽기
#     h_rebar_space = ws_retrofit.Range('J%s:J%s' %(startrow, startrow+element_info.shape[0]-1)).Value
#     h_rebar_space_array = np.array(h_rebar_space)[:,0] # list of tuples -> np.array
#     # 전단보강 가능여부 읽기
#     avail = ws_retrofit.Range('W%s:W%s' %(startrow, startrow+element_info.shape[0]-1)).Value
#     # DCR 읽기
#     dcr = ws_retrofit.Range('AK%s:AK%s' %(startrow, startrow+element_info.shape[0]-1)).Value
#     # DCR 값에 따른 np,array 생성 (NG가 있는 경우 = 1, NG가 없는 경우 = 0)
#     dcr_array = np.array([1 if 'N.G' in row else 0 for row in dcr])

#     # (NG) & (수평철근간격이 최소철근간격에 도달하지 않은) 부재들의 철근 간격 down
#     h_rebar_space_array_updated = np.where(((dcr_array == 1) & (h_rebar_space_array-10 >= rebar_limit[1]))
#                                            , h_rebar_space_array-10, h_rebar_space_array)

#     # 수평철근간격 before & updated가 동일한 경우(철근간격이 update되지 않는 경우) break
#     if np.array_equal(h_rebar_space_array, h_rebar_space_array_updated):
#         break            

#     # Horizontal Rebar 간격의 변경된 값을 Excel에 다시 입력
#     ws_retrofit.Range('J%s:J%s' %(startrow, startrow+element_info.shape[0]-1)).Value\
#     = [[i] for i in h_rebar_space_array_updated]        
    

#     # Horizontal Diameter 직경/간격이 변경된(DCR == NG) 경우, 색 변경하기
#     h_rebar_space_diff_idx = np.where(h_rebar_space_array != h_rebar_space_array_updated)
#     for j in h_rebar_space_diff_idx[0]:
#         ws_retrofit.Range('J%s' %str(startrow+int(j))).Font.ColorIndex = 3 # 3 : 빨간색