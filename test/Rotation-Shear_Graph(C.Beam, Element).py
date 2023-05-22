# -*- coding: utf-8 -*-
"""
Created on Mon May 22 13:31:11 2023

@author: hwlee
"""

import pandas as pd
import matplotlib.pyplot as plt

file_path = r'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\Results\107\LB2_1_RF_Analysis_Result.xlsx'

rotation_data = pd.read_excel(file_path, sheet_name='Frame Results - Bending Deform', skiprows=[0,2], usecols=[5,7,8,11,14,15])
force_data = pd.read_excel(file_path, sheet_name='Frame Results - End Forces', skiprows=[0,2], usecols=[5,7,8,11])

rotation_data = rotation_data[rotation_data['Distance from I-End'] == 0]
rotation_data = rotation_data.drop_duplicates(subset=['Load Case', 'Step Number'])
force_data = force_data.drop_duplicates(subset=['Load Case', 'Step Number'])

force_data.sort_values(by=['Load Case', 'Step Number'], inplace=True)
force_data.reset_index(inplace=True, drop=True)

rotation_data.sort_values(by=['Load Case', 'Step Number'], inplace=True)
rotation_data.reset_index(inplace=True, drop=True)

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

#%% GRAPH
fig, ax = plt.subplots(1,1,figsize=(6,5))

for load_name in MCE_load_name_list:
    rotation = rotation_data[rotation_data['Load Case'].str.contains(load_name)]
    force = force_data[force_data['Load Case'].str.contains(load_name)]
    
    ax.plot(rotation['R3'], force['V2 I-End'], '-o', label=load_name, markersize=2, linewidth=0.1)
    
ax.axvline(x=0, color='k', linestyle='-')
ax.axhline(y=0, color='k', linestyle='-')    
ax.grid(linestyle='-.')
ax.set_xlabel('Rotation(rad)')
ax.set_ylabel('V(kN)')
ax.set_title('LB2_1_RF')
ax.legend()      


