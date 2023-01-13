import pandas as pd
import numpy as np
import os
from collections import deque  # Double-ended Queue : 자료의 앞, 뒤 양 방향에서 자료를 추가하거나 제거가능
import matplotlib.pyplot as plt
import matplotlib as mpl
import win32com.client

pd.options.mode.chained_assignment = None # SettingWithCopyWarning 안뜨게 하기

# Analysis Result
result_path = r'L:\33. 내진성능평가 & 성능설계\02. 진행 중\21-GR-222 광명 4R구역 주택재개발사업 성능기반내진설계\결과정리\102'
result_xlsx = 'Analysis Result' # 해석결과에 공통으로 포함되는 이름 (확장자X)

# Data Conversion Sheet, Column Sheet, Beam Sheet
input_path = r'L:\33. 내진성능평가 & 성능설계\02. 진행 중\21-GR-222 광명 4R구역 주택재개발사업 성능기반내진설계\결과정리\102'
input_xlsx = '102D_Data Conversion_Shear Wall Type.xlsx'

#%%
story_info = pd.DataFrame()
deformation_cap = pd.DataFrame()

input_data_raw = pd.ExcelFile(input_path + '\\' + input_xlsx)
input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', 'Output_C.Beam Properties'], skiprows=3)
input_data_raw.close()
    
story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
beam_info = input_data_sheets['Output_C.Beam Properties'].iloc[:,np.r_[0:6,9,10,12,13,14,15,56,60,62]]
deformation_cap = input_data_sheets['Output_C.Beam Properties'].iloc[:,[0,47,48,49]]
    
    
story_info.columns = ['Index', 'Story Name', 'Height(mm)']
beam_info.columns = ['Name', 'Length(mm)', 'b(mm)', 'h(mm)', 'D(mm)'
                     , 'Concrete Grade', 'Main Rebar Type', 'Stirrup Rebar Type'
                     , 'Top1 EA', 'Top2 EA', 'Stirrup EA', 'Stirrup Spacing(mm)'
                     , 'Vy(kN)', 'Vn(kN)', 'Vy<=Vn']
deformation_cap.columns = ['Name', 'IO', 'LS', 'CP']

#%% Analysis Result 불러오기

to_load_list = []
file_names = os.listdir(result_path)
for file_name in file_names:
    if (result_xlsx in file_name) and ('~$' not in file_name):
        to_load_list.append(file_name)

beam_rot_data = pd.DataFrame()

for i in to_load_list:
    result_data_raw = pd.ExcelFile(result_path + '\\' + i)
    result_data_sheets = pd.read_excel(result_data_raw, ['Frame Results - Bending Deform', 'Node Coordinate Data',\
                                                     'Element Data - Frame Types'], skiprows=[0,2])
    
    columne_name_to_choose = ['Group Name', 'Element Name', 'Load Case'
                              , 'Step Type', 'Distance from I-End', 'R2', 'R3']        
    beam_rot_data_temp = result_data_sheets['Frame Results - Bending Deform'].loc[:,columne_name_to_choose]
    beam_rot_data = pd.concat([beam_rot_data, beam_rot_data_temp])
    
node_data = result_data_sheets['Node Coordinate Data'].iloc[:,[1,4]]
element_data = result_data_sheets['Element Data - Frame Types'].iloc[:,[2,5,7]] # beam의 양 nodes중 한 node에서의 rotation * 2
            
beam_rot_data.rename(columns={'R2':'H2 Rotation(rad)', 'R3':'H3 Rotation(rad)'}, inplace=True)
node_data.columns = ['Node ID', 'V(mm)']
element_data.columns = ['Element Name', 'Property Name', 'I-Node ID']

#%% temporary ((L), (R) 등 지우기)
element_data.loc[:, 'Property Name'] = element_data.loc[:, 'Property Name'].str.split('(').str[0]

#%% 필요없는 부재 빼기, 필요한 부재만 추출

beam_rot_data = beam_rot_data[beam_rot_data['Distance from I-End'] == 0]

element_data = element_data.drop_duplicates()
node_data = node_data.drop_duplicates()

beam_rot_data = pd.merge(beam_rot_data, element_data, how='left')
beam_rot_data = pd.merge(beam_rot_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID')

beam_rot_data = beam_rot_data[beam_rot_data['Property Name'].notna()]

beam_rot_data.reset_index(inplace=True, drop=True)

#%% beam_rot_data의 값 수정(H1, H2 방향 중 major한 방향의 rotation값만 추출, 그리고 2배)
major_rot = []
for i, j in zip(beam_rot_data['H2 Rotation(rad)'], beam_rot_data['H3 Rotation(rad)']):
    if abs(i) >= abs(j):
        major_rot.append(i)
    else: major_rot.append(j)

beam_rot_data['Major Rotation(rad)'] = major_rot
 
# 필요한 정보들만 다시 모아서 new dataframe
beam_rot_data = beam_rot_data.iloc[:, [0,1,7,10,2,3,11]]
    
#%% 지진파 이름 list 만들기

load_name_list = []
for i in beam_rot_data['Load Case'].drop_duplicates():
    new_i = i.split('+')[1]
    new_i = new_i.strip()
    load_name_list.append(new_i)

gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]

seismic_load_name_list.sort()

DE_load_name_list = [x for x in load_name_list if 'DE' in x] # base shear로 사용할 지진파 개수 산정을 위함
MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]

#%% 성능기준(LS, CP) 정리해서 merge
    
beam_rot_data = pd.merge(beam_rot_data, deformation_cap, how='left', left_on='Property Name', right_on='Name')

beam_rot_data['DE Rotation(rad)'] = beam_rot_data['Major Rotation(rad)'].abs()
beam_rot_data['MCE Rotation(rad)'] = beam_rot_data['Major Rotation(rad)'].abs()

beam_rot_data = beam_rot_data[beam_rot_data['Name'].notna()]

#%% DE 결과
if len(DE_load_name_list) != 0:
    
    beam_rot_data_total_DE = pd.DataFrame()    
    
    for load_name in DE_load_name_list:
    
        temp_df = beam_rot_data[beam_rot_data['Load Case'].str.contains('{}'.format(load_name))]\
                  .groupby(['Element Name'])['DE Rotation(rad)']\
                  .agg(**{'Rotation avg':'mean'})['Rotation avg']
                      
        beam_rot_data_total_DE['{}'.format(load_name)] = temp_df.tolist()
        
    beam_rot_data_total_DE['Element Name'] = temp_df.index 
    beam_rot_data_total_DE.reset_index(inplace=True, drop=True)
    
    beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, element_data, how='left')
    beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
    beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, story_info, how='left', left_on='V(mm)', right_on='Height(mm)')
    beam_rot_data_total_DE.sort_values('Height(mm)', inplace=True)
    
# 평균 열 생성
    beam_rot_data_total_DE['DE avg'] = beam_rot_data_total_DE.iloc[:,list(range(0,len(DE_load_name_list)))].mean(axis=1)
    beam_rot_data_total_DE = beam_rot_data_total_DE.loc[:,['Property Name', 'DE avg']]
    beam_rot_data_total_DE.columns = ['Name', 'DE avg']
        
#%% MCE 결과 Plot
    
if len(MCE_load_name_list) != 0:
    
    beam_rot_data_total_MCE = pd.DataFrame()    
    
    for load_name in MCE_load_name_list:
    
        temp_df = beam_rot_data[beam_rot_data['Load Case'].str.contains('{}'.format(load_name))]\
                  .groupby(['Element Name'])['MCE Rotation(rad)']\
                  .agg(**{'Rotation avg':'mean'})['Rotation avg']
                      
        beam_rot_data_total_MCE['{}'.format(load_name)] = temp_df.tolist()
        
    beam_rot_data_total_MCE['Element Name'] = temp_df.index    
    beam_rot_data_total_MCE.reset_index(inplace=True, drop=True)
    
    beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, element_data, how='left')
    beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
    beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, story_info, how='left', left_on='V(mm)', right_on='Height(mm)')
    beam_rot_data_total_MCE.sort_values('Height(mm)', inplace=True)
    
# 평균 열 생성
    beam_rot_data_total_MCE['MCE avg'] = beam_rot_data_total_MCE.iloc[:,list(range(0,len(MCE_load_name_list)))].mean(axis=1)
    beam_rot_data_total_MCE = beam_rot_data_total_MCE.loc[:,['Property Name', 'MCE avg']]
    beam_rot_data_total_MCE.columns = ['Name', 'MCE avg']

beam_rot_data_total = pd.merge(beam_rot_data_total_DE, beam_rot_data_total_MCE, how='left')

#%% 조작용 코드
# 없애고 싶은 부재의 이름 입력(error_beam 확인 후!, DE, MCE에서 다 없어짐)
# beam_rot_data_total = beam_rot_data_total.drop(beam_rot_data_total
#                                                [(beam_rot_data_total['Name']
#                                                  .str.contains('LB4_'))].index) # 부재 이름에 따라
# beam_rot_data_total = beam_rot_data_total.drop(beam_rot_data_total
#                                                [(beam_rot_data_total['Name']
#                                                  .str.contains('B31a_1_'))].index) # 부재 이름에 따라(반복)

# beam_rot_data_total = beam_rot_data_total.drop([12,13]) #index에 따라

# beam_rot_data_total = beam_rot_data_total.drop(beam_rot_data_total
#                                                [(beam_rot_data_total['MCE avg']>0.05)].index) # 값에 따라

#%%

beam_rot_data_total_avg = beam_rot_data_total.groupby(['Name']).mean()
beam_rot_data_total_avg.reset_index(inplace=True, drop=False)

#%%
beam_info = pd.concat([beam_info, deformation_cap.iloc[:,[1,2,3]]], axis=1)

# beam_rot_data_total_DE.dropna(subset='Property Name', inplace=True)
# beam_rot_data_total_MCE.dropna(subset='Property Name', inplace=True)

#### OLD VERSION ####    
# 이전 버전의 네이밍에 맞게 merge하는 방법
# new_name_list = []
# for i in SWR_criteria['Name']:
#     if i.count('_') != 2:
#         new_name_list.append(i.split('_')[0] + '_' + i.split('_')[2])

# SWR_criteria['Name'] = new_name_list
# deformation_cap['Name'] = new_name_list
# results_wall['gage_name'] = new_name_list

# new_name_list_2 = []
# for i in SWR_avg_total['gage_name']:
#     new_name_2 = i.split('_')[0] + '_1_' + i.split('_')[1]
#     new_name_list_2.append(new_name_2)

# SWR_avg_total['gage_name'] = new_name_list_2
#####################

### SWR avg total에 SWR criteria join(wall name 기준)
BR_avg_total = pd.merge(beam_info, beam_rot_data_total_avg, how='left')
BR_avg_total.dropna(subset='IO', inplace=True)

#%% IO,LS,CP 판별
def determine_cap(rotation, cri_df): # 길이가 같은 열들
    rotation_arr = np.array(rotation)
    cri_arr = np.array(cri_df) 
    combined_arr = np.column_stack((rotation_arr, cri_arr))
    
    det_result = []
    for row in combined_arr:
        if row[0]<=row[1]:
            det_result.append('IO')
        elif (row[1]<row[0]) & (row[0]<=row[2]):
            det_result.append('LS')
        elif (row[2]<row[0]) & (row[0]<=row[3]):
            det_result.append('CP')
        else: det_result.append('NG')
        
    return det_result


BR_avg_total['BR_DCR_DE'] = determine_cap(BR_avg_total['DE avg']
                                            , BR_avg_total.iloc[:,[15,16,17]])
BR_avg_total['BR_DCR_MCE'] = determine_cap(BR_avg_total['MCE avg']
                                             , BR_avg_total.iloc[:,[15,16,17]])

BR_avg_total.dropna(subset='Name', inplace=True)
BR_avg_total.dropna(subset='DE avg', inplace=True)

#%% 배근 표기 combine
v_rebar_combined = []
h_rebar_combined = []
for idx, row in BR_avg_total.iloc[:,[6,7,8,9,10,11]].iterrows():
    # vertical rebar
    v_rebar_combined.append(str(int(row[2]+row[3]))+'-'+row[0])
    
    # horizontal rebar
    h_rebar_combined.append(str(int(row[4]))+'-'+row[1]+'@'+str(int(row[5])))

BR_avg_total['V.Rebar_Combined'] = v_rebar_combined
BR_avg_total['H.Rebar_Combined'] = h_rebar_combined

BR_output = BR_avg_total.iloc[:,np.r_[0:6,22,23,15,16,17,18,20,19,21,12,13,14]]

#%% Output

# 소수점 자리 설정
BR_output.iloc[:,[8,9,10,11,13]] = BR_output.iloc[:,[8,9,10,11,13]].round(4)
BR_output.iloc[:,[15,16]] = BR_output.iloc[:,[15,16]].round(2)


# 111동 부재명에서 R 떼기
# remove_R_list = []
# for i in BR_output.iloc[:,0]:
#     remove_R = i[1:]
#     remove_R_list.append(remove_R)
    
# BR_output.iloc[:,0] = remove_R_list