import pandas as pd
import numpy as np
from collections import deque


##### 사용자 입력 ##############################################################
input_file_path = r'D:\이형우\성능기반 내진설계\22-RM-200 창원 신월2구역 재건축 아파트 신축공사 성능기반 내진설계\02. 접수자료\230717 신월2구역 실시설계 배근\230711-신월2구역-103동-벽체 일람표 (수정) - 0002.xlsx'
sheet_num = [1,2,3,4]
rebar_type = '일반용'
###############################################################################

# Read EXCEL
wall = pd.read_excel(input_file_path, sheet_name=sheet_num, usecols='A,E:I,L:P,S:W,Z:AD,AG:AK', skiprows=2)
wall_df = pd.concat(wall.values(), ignore_index=True)

wall_df = wall_df.dropna(how='all')
wall_df.reset_index(inplace=True, drop=True)

# Add Story Column to every Column Chunks
wall_df = wall_df.iloc[:,[0,1,2,3,4,5, 0,6,7,8,9,10, 0,11,12,13,14,15, 0,16,17,18,19,20, 0,21,22,23,24,25]]

# Create Column Name List
wall_df_name = []
for i in list(range(1,6)):
    wall_df_name.append('story'+str(i))
    wall_df_name.append('thk'+str(i))
    wall_df_name.append('v_rebar'+str(i))
    wall_df_name.append('v_space'+str(i))
    wall_df_name.append('h_rebar'+str(i))
    wall_df_name.append('h_space'+str(i))    
wall_df.columns = wall_df_name

# Get Indices of the rows which match 'FLOOR' in the first column
wall_split_idx = wall_df.index[wall_df.iloc[:,0] == 'FLOOR'].tolist()
wall_split_idx.append(len(wall_df)-1+1) # Append the last index of wall_df (if not, the last chunk is not included)

# Split Whole df by Indices
wall_df_split = []
for i in range(len(wall_split_idx)):

    df = wall_df.iloc[wall_split_idx[i]:wall_split_idx[i+1],:]
    # Create "index" column in df to get i argument value while using pd.wide_to_long
    df.reset_index(inplace=True, drop=False)
    # Reshape df to Combine the Information of every element into Single Common Columns
    df_reshape = pd.wide_to_long(df, stubnames=['story', 'thk', 'v_rebar', 'v_space', 'h_rebar', 'h_space']
                                  , i = 'index', j='suffix') # i & j should be fed to run this function   
    wall_df_split.append(df_reshape)
    
    # Break if i+1 meets the end of the range
    if i+2 == len(wall_split_idx):
        break

# Concatenate Split Dataframes
wall_df_edited = pd.concat(wall_df_split, ignore_index=True)

# Drop Unnecessary Rows
wall_df_edited = wall_df_edited.dropna(subset=['thk'])
wall_df_edited = wall_df_edited.dropna(subset=['story'])
wall_df_edited.reset_index(inplace=True, drop=True)

# Create Blank 'name' Column
wall_df_edited['name'] = 'W'

# Fill in [Name] column
wall_name = deque() # create empty deque

# Iterrate rows of wall_data
for idx, row in wall_df_edited.iterrows():
    if row['story'] == 'FLOOR':
        wall_name.append(row['thk'])
        # 빈칸 열 만들기
        row['story'] = np.nan
        row['thk'] = np.nan
    else:
        if len(wall_name) == 1:
            row['name'] = row['name'] + str(wall_name[0])
            wall_name.pop()

# Create empty column        
wall_df_edited['dummy'] = np.nan

# Get the Index where name == 'W1'
W1_idx = wall_df_edited.index[wall_df_edited['name'] == 'W1'] 

# Change second 'W1' to 'CW1'
wall_df_edited.iloc[W1_idx[1]:,6] = 'C' + wall_df_edited.iloc[W1_idx[1]:,6]

# Make Blank if name is 'W'
new_wall_name = []
for i in wall_df_edited['name']:
    if (i == 'CW') | (i == 'W'):
        new_wall_name.append('')
    else: new_wall_name.append(i)
wall_df_edited['name'] = new_wall_name

# Make '일반용' or '내진용' if name != ''
rebar_type_list = []
for i in wall_df_edited['name']:
    if i != '':
        rebar_type_list.append(rebar_type)
    else: rebar_type_list.append('')
wall_df_edited['type'] = rebar_type_list

# Replace np.nana to ''
wall_df_edited = wall_df_edited.replace(np.nan, '', regex=True)

# Remove 'D' if 'D' is attached on 간격배근
wall_df_edited['v_space'] = wall_df_edited['v_space'].astype(str).str.replace('D', '')

         
# Change the order of columns
wall_output = wall_df_edited.iloc[1:,[6,0,7,1,2,3,4,5,8]]