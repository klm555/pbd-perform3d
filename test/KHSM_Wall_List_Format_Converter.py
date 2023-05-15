# -*- coding: utf-8 -*-
"""
Created on Wed Apr 26 13:47:50 2023

@author: hwlee

For Gimhae Sinmoon

"""

import pandas as pd
import numpy as np
import os
from collections import deque

# =======  User input  ========================================================
data_path = r'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\230510-김해신문1지구 벽체 수평배근 변경'
data_file = '230509-GIMHAE-108D-WALL LIST-0004.xlsx'
data_sheet = 'BatchWall' # Sheet about Wall Information
data_sheet2 = 'Length Info' # Sheet about Length Information
# =============================================================================

# Load data
wall_data = pd.read_excel(os.path.join(data_path, data_file), sheet_name=data_sheet, usecols=[0,2,4,5,7,8])
# Change the name of columns in data
wall_data.columns=['Story', 'Thickness', 'V.Bar', 'V.Space(mm)', 'H.Bar', 'H.Space(mm)']

# Drop(delete) unnecessary rows
wall_data = wall_data.drop(wall_data[(wall_data['Story'].isnull()) # row which has 'nan' in [Story] column
                                     | (wall_data['Story'] == 'Story')].index) # row which has 'Story' in [Story] column
wall_data.reset_index(inplace=True, drop=True)

# Create [Name] column(w/ empty cells)
wall_data['Name'] = ''

# Fill in [Name] column
wall_name = deque() # create empty deque

# Iterrate rows of wall_data
for idx, row in wall_data.iterrows():
    if pd.isnull(row['Thickness']):
        wall_name.append(row['Story'])
        # 빈칸 열 만들기
        row['Story'] = np.nan
        row['Name'] = np.nan
    else:
        if len(wall_name) == 1:
            row['Name'] = wall_name[0]
            wall_name.pop()

# Create empty column        
wall_data['dummy'] = np.nan

# Change the order of columns
wall_data = wall_data.iloc[1:,[6,0,7,1,2,3,4,5]]

# (optional) Wall name list (Use wherever you want)
# wall_name = pd.Series([i for i in wall_data['Name'] if str(i) != 'nan' if i != ''])

#%% Add Length(total, element) Column

length_data = pd.read_excel(os.path.join(data_path, data_file), sheet_name=data_sheet2, usecols=[1,2,3,7], skiprows=2)
length_data.columns = ['Name', 'Lo', 'Le', 'Type']

wall_data = pd.merge(wall_data, length_data, how='left', on='Name')
wall_data = wall_data.iloc[:,np.r_[0:8,10,8,9]]
