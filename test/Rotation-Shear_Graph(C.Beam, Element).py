# -*- coding: utf-8 -*-
"""
Created on Mon May 22 13:31:11 2023

@author: hwlee
"""

import pandas as pd
import matplotlib.pyplot as plt

file_path = r'D:\이형우\성능기반 내진설계\22-RM-200 창원 신월2구역 재건축 아파트 신축공사 성능기반 내진설계\08. Analysis Results\108D\PB2-17_Analysis_Result_MCE6.xlsx'

rotation_data = pd.read_excel(file_path, sheet_name='Frame Results - Bending Deform', skiprows=[0,2], usecols=[5,7,8,9,14,15])
force_data = pd.read_excel(file_path, sheet_name='Frame Results - End Forces', skiprows=[0,2], usecols=[5,7,8,19])

rotation_data = rotation_data[rotation_data['Load Case'].str.contains('MCE61')]
force_data = force_data[force_data['Load Case'].str.contains('MCE61')]

rotation_data = rotation_data[rotation_data['Step Type'] == 'Time']
rotation_data = rotation_data[rotation_data['Point ID'] == 5] # or 5 ????????????????
rotation_data = rotation_data.drop_duplicates(subset=['Load Case', 'Step Number'])

force_data = force_data[force_data['Step Type'] == 'Time']
force_data = force_data.drop_duplicates(subset=['Load Case', 'Step Number'])

force_data.sort_values(by=['Load Case', 'Step Number'], inplace=True)
force_data.reset_index(inplace=True, drop=True)

rotation_data.sort_values(by=['Load Case', 'Step Number'], inplace=True)
rotation_data.reset_index(inplace=True, drop=True)

# 초반 100개만 slick
# rotation_data = rotation_data.iloc[0:1001, :]
# force_data = force_data.iloc[0:1001, :]

#%% 지진파 이름 list 만들기
load_name_list = []
for i in rotation_data['Load Case'].drop_duplicates():
    new_i = i.split('+')[1]
    new_i = new_i.strip()
    load_name_list.append(new_i)

gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]

seismic_load_name_list.sort()

DE_load_name_list = [x for x in load_name_list if 'DE' in x]
MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]

#%% GRAPHr
fig, ax = plt.subplots(1,1,figsize=(6,5))

for load_name in MCE_load_name_list:
    rotation = rotation_data[rotation_data['Load Case'].str.contains(load_name)]
    force = force_data[force_data['Load Case'].str.contains(load_name)]
    
    ax.plot(rotation['R3'], force['M3 J-End'], '-o', label=load_name, markersize=2, linewidth=0.1)
    
ax.axvline(x=0, color='k', linestyle='-')
ax.axhline(y=0, color='k', linestyle='-')    
ax.grid(linestyle='-.')
ax.set_xlabel('Rotation(rad)')
ax.set_ylabel('M(kN-mm)')
ax.set_title('PB2-17_1_25F')
ax.legend()      


