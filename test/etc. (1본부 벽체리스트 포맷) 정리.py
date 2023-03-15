# -*- coding: utf-8 -*-
"""
Created on Wed May 11 09:38:16 2022

@author: hwlee
"""

import pandas as pd
import numpy as np
import io

#%% 사용자 입력

data_path = r'D:\이형우\내진성능평가\광명 4R\접수자료\220905-103동 보, 기둥 변경\04_벽체' # data 폴더 경로
data_xlsx = '220905-광명4R-103동-WALL LIST-0004(전이층 변경).xlsx' # 파일명

output_xlsx = '220905-WALL_LIST.xlsx'

#%% File Load

wall_list_raw = pd.ExcelFile(data_path + '\\' + data_xlsx)
wall_list_sheets = pd.read_excel(wall_list_raw, None, skiprows=28)

pd.options.mode.chained_assignment = None # SettingWithCopyWarning 안뜨게 하기

count = 1
for wall_list_value in wall_list_sheets.values():
   
    wall_list = wall_list_value.iloc[5:, 2:]

    for i in range(5):
        globals()['wall_data_{}'.format(i+1)] = wall_list.iloc[:, [0,4+7*i,5+7*i,6+7*i,7+7*i,8+7*i]]
        
        # globals()['wall_data_{}'.format(i+1)].iloc[:, 1].replace('', np.nan, inplace=True)
        # globals()['wall_data_{}'.format(i+1)] = globals()['wall_data_{}'.format(i+1)][globals()['wall_data_{}'.format(i+1)].iloc[:, 1].notna()]
        globals()['wall_data_{}'.format(i+1)]['Name'] = wall_list_value.iloc[0, 4+2+7*i]
        globals()['wall_data_{}'.format(i+1)].columns = ['Story', 'Thickness', 'V. Bar', 'V. Spacing', 'H. Bar', 'H. Spacing', 'Name']
        
    wall_data = pd.concat([wall_data_1, wall_data_2, wall_data_3, wall_data_4, wall_data_5], ignore_index=True)
    wall_data = wall_data[wall_data.iloc[:,1].notna()]
    
    if count == 1:
        wall_data_total = wall_data.copy()
        
    if count >= 2:
        wall_data_total = wall_data_total.append(wall_data, ignore_index=True)
        
    count += 1

#%% 입력용 포맷으로 다시 정리
wall_list_output = wall_data_total.iloc[:,[6,0,0,1,2,3,4,5]]

# @ 제거, D 제거
wall_list_output = wall_list_output.replace('@', '', regex=True)
wall_list_output['H. Bar'] = wall_list_output['H. Bar'].replace('D', '', regex=True)

# D 일괄 붙이기
wall_list_output = wall_list_output.astype(str)

new_v_bar = []
new_h_bar = []

for idx, row in wall_list_output.iterrows():
    if not '-' in row[4]:
        new_v_bar.append('D' + row[4])
    else: new_v_bar.append(row[4])
    
    if not '-' in row[6]:
        new_h_bar.append('D' + row[6])
    else: new_h_bar.append(row[6])

wall_list_output['H. Bar'] = new_h_bar
wall_list_output['V. Bar'] = new_v_bar

# nan(str) to np.nan
wall_list_output['V. Spacing'] = np.where(wall_list_output['V. Spacing'] == 'nan', np.nan, wall_list_output['V. Spacing'])
wall_list_output['H. Spacing'] = np.where(wall_list_output['H. Spacing'] == 'nan', np.nan, wall_list_output['H. Spacing'])

# Rebar Type 열 만들기
wall_list_output['Rebar Type'] = '일반용'

#%% 필요없는 부재 빼기
wall_list_output = wall_list_output[wall_list_output.iloc[:,1].str.contains('PH') == False]

#%% 부재 사이에 빈 줄 추가 (어케하는겨)

mask = wall_list_output['Name'].ne(wall_list_output['Name'].shift(-1))
wall_list_output_temp = pd.DataFrame('', index=mask.index[mask] + .5,\
                                     columns=wall_list_output.columns)

wall_list_output = pd.concat([wall_list_output, wall_list_output_temp]).sort_index()\
                   .reset_index(drop=True).iloc[:-1]
    
#%% 엑셀로 만들기~

wall_list_output.to_excel(data_path + '\\'+ output_xlsx, sheet_name = 'Wall List', index = False)
