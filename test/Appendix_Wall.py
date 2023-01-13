import pandas as pd
import numpy as np
import os
from collections import deque  # Double-ended Queue : 자료의 앞, 뒤 양 방향에서 자료를 추가하거나 제거가능
import matplotlib.pyplot as plt
import matplotlib as mpl
import win32com.client

pd.options.mode.chained_assignment = None # SettingWithCopyWarning 안뜨게 하기

# Analysis Result
result_path = r'L:\33. 내진성능평가 & 성능설계\02. 진행 중\21-GR-222 광명 4R구역 주택재개발사업 성능기반내진설계\결과정리\107'
result_xlsx = 'Analysis Result' # 해석결과에 공통으로 포함되는 이름 (확장자X)

# Data Conversion Sheet, Column Sheet, Beam Sheet
input_path = r'L:\33. 내진성능평가 & 성능설계\02. 진행 중\21-GR-222 광명 4R구역 주택재개발사업 성능기반내진설계\결과정리\107'
input_xlsx = '107D_Data Conversion_Shear Wall Type_Ver.1.9_11.xlsx'


#%%
story_info = pd.DataFrame()
deformation_cap = pd.DataFrame()

input_data_raw = pd.ExcelFile(input_path + '\\' + input_xlsx)
input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', 'Results_Wall', 'Output_Wall Properties'], skiprows=3)
input_data_raw.close()

story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
results_wall = input_data_sheets['Results_Wall'].iloc[:,np.r_[0:15,22]]
deformation_cap = input_data_sheets['Results_Wall'].iloc[:,[0,11,12,13,14,38,39,44,45,50,51]]
name_cri = input_data_sheets['Output_Wall Properties'].iloc[:,0]

story_info.columns = ['Index', 'Story Name', 'Height(mm)']
deformation_cap.columns = ['Name', 'Vu_DE_H1', 'Vu_DE_H2', 'Vu_MCE_H1', 'Vu_MCE_H2'\
                           , 'IO(H1)', 'IO(H2)', 'LS(H1)', 'LS(H2)', 'CP(H1)', 'CP(H2)']
results_wall.columns = ['gage_name', 'Length(mm)', 'Thickness(mm)', 'Concrete Grade'
                        , 'Steel Type', 'V.Rebar Type', 'V.Rebar Spacing', 'V.Rebar EA'
                        , 'H.Rebar Type', 'H.Rebar Spacing', 'Nu(kN)', 'Vu_DE_H1(kN)'
                        , 'Vu_DE_H2(kN)', 'Vu_MCE_H1(kN)', 'Vu_MCE_H2(kN)', 'Vn(kN)']
    
    
story_name = story_info.loc[:, 'Story Name']

name_cri = pd.concat([pd.Series(name_cri.index), name_cri], axis=1)
name_cri.columns = ['Index', 'gage_name']

#%% Analysis Result 불러오기

to_load_list = []
file_names = os.listdir(result_path)
for file_name in file_names:
    if (result_xlsx in file_name) and ('~$' not in file_name):
        to_load_list.append(file_name)

wall_rot_data = pd.DataFrame()

for i in to_load_list:
    result_data_raw = pd.ExcelFile(result_path + '\\' + i)
    result_data_sheets = pd.read_excel(result_data_raw,\
                                        ['Gage Results - Wall Type', 'Node Coordinate Data',\
                                        'Gage Data - Wall Type', 'Element Data - Shear Wall']\
                                        ,skiprows=[0,2])
    columne_name_to_choose = ['Group Name', 'Element Name', 'Load Case'
                              , 'Step Type', 'Rotation', 'Performance Level']
    wall_rot_data_temp = result_data_sheets['Gage Results - Wall Type'].loc[:,columne_name_to_choose]
    wall_rot_data = pd.concat([wall_rot_data, wall_rot_data_temp])
    
node_data = result_data_sheets['Node Coordinate Data'].iloc[:,[1,2,3,4]]
gage_data = result_data_sheets['Gage Data - Wall Type'].iloc[:,[2,7,9,11,13]] # beam의 양 nodes중 한 node에서의 rotation * 2
element_data = result_data_sheets['Element Data - Shear Wall'].iloc[:,[2,5,7,9,11,13]] # beam의 양 nodes중 한 node에서의 rotation * 2

#%% Gage Data & Result에 Node 정보 매칭
    
gage_data = gage_data.drop_duplicates()
node_data = node_data.drop_duplicates()

#%% 지진파 이름 list 만들기

load_name_list = []
for i in wall_rot_data['Load Case'].drop_duplicates():
    new_i = i.split('+')[1]
    new_i = new_i.strip()
    load_name_list.append(new_i)

gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]

seismic_load_name_list.sort()

DE_load_name_list = [x for x in load_name_list if 'DE' in x] # base shear로 사용할 지진파 개수 산정을 위함
MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]

#%% 데이터 매칭 후 결과뽑기

### Gage data에서 Element Name, I-Node ID 불러와서 v좌표 match하기
gage_num = len(gage_data) # gage 개수 얻기

# Gage data의 i, j-node 좌표
gage_data = gage_data.join(node_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')
gage_data.rename({'H1' : 'I_H1', 'H2' : 'I_H2', 'V' : 'I_V'}, axis=1, inplace = True) # I node 의 H1, H2, V 좌표 가져오기
gage_data = gage_data.join(node_data.set_index('Node ID')[['H1', 'H2']], on='J-Node ID')
gage_data.rename({'H1' : 'J_H1', 'H2' : 'J_H2'}, axis=1, inplace=True) # J node 의 H1 좌표 가져오기

# vector ij의 x, y방향 성분 구하기
gage_data['J_H1-I_H1'] = gage_data.apply(lambda x: x['J_H1']- x['I_H1'], axis=1)
gage_data['J_H2-I_H2'] = gage_data.apply(lambda x: x['J_H2']- x['I_H2'], axis=1)
gage_data['I_H1-J_H1'] = gage_data.apply(lambda x: x['I_H1']- x['J_H1'], axis=1)
gage_data['I_H2-J_H2'] = gage_data.apply(lambda x: x['I_H2']- x['J_H2'], axis=1)

# gage 벡터, (1,0)벡터 만들기(array)
gage_vector_ij = gage_data.iloc[:,[10,11]].values
gage_vector_ji = gage_data.iloc[:,[12,13]].values
e1_vector = np.array([1,0])
e2_vector = np.array([0,1])


# Vector ij와 (1,0)의 Cosine Similarity 구하기
def cos_sim(arr, unit_arr):
    result = np.dot(arr, unit_arr) / (np.linalg.norm(arr, axis=1)*np.linalg.norm(unit_arr))
    return result
       
gage_data['Similarity ij-e1'] = cos_sim(gage_vector_ij, e1_vector)
gage_data['Similarity ij-e2'] = cos_sim(gage_vector_ij, e2_vector)
gage_data['Similarity ji-e1'] = cos_sim(gage_vector_ji, e1_vector)
gage_data['Similarity ji-e2'] = cos_sim(gage_vector_ji, e2_vector)


# Wall element data의 i, j-node 좌표
element_data = element_data.join(node_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')
element_data.rename({'H1' : 'I_H1', 'H2' : 'I_H2', 'V' : 'I_V'}, axis=1, inplace=True)
element_data = element_data.join(node_data.set_index('Node ID')[['H1', 'H2']], on='J-Node ID')
element_data.rename({'H1' : 'J_H1', 'H2' : 'J_H2'}, axis=1, inplace=True)

# vector ij의 x, y방향 성분 구하기
element_data['J_H1-I_H1'] = element_data.apply(lambda x: x['J_H1']- x['I_H1'], axis=1)
element_data['J_H2-I_H2'] = element_data.apply(lambda x: x['J_H2']- x['I_H2'], axis=1)
element_data['I_H1-J_H1'] = element_data.apply(lambda x: x['I_H1']- x['J_H1'], axis=1)
element_data['I_H2-J_H2'] = element_data.apply(lambda x: x['I_H2']- x['J_H2'], axis=1)

# element 벡터, (1,0)벡터 만들기(array)
element_vector_ij = element_data.iloc[:,[11,12]].values
element_vector_ji = element_data.iloc[:,[13,14]].values

# Vector ij와 (1,0)의 Cosine Similarity 구하기
element_data['Similarity ij-e1'] = cos_sim(element_vector_ij, e1_vector)
element_data['Similarity ij-e2'] = cos_sim(element_vector_ij, e2_vector)
element_data['Similarity ji-e1'] = cos_sim(element_vector_ji, e1_vector)
element_data['Similarity ji-e2'] = cos_sim(element_vector_ji, e2_vector)

### wall element data 와 SWR gage data 연결하기(wall 이름)
gage_data = gage_data.join(element_data.set_index(['I-Node ID', 'Similarity ij-e1', 'Similarity ij-e2'])\
                           ['Property Name'], on=['I-Node ID', 'Similarity ij-e1', 'Similarity ij-e2'])
gage_data.rename({'Property Name' : 'gage_name'}, axis=1, inplace=True)

# i, j 노드가 반대로 설정된 경우
gage_data = gage_data.join(element_data.set_index(['I-Node ID', 'Similarity ij-e1', 'Similarity ij-e2'])\
                           ['Property Name'], on=['J-Node ID', 'Similarity ji-e1', 'Similarity ji-e2'])
gage_data.rename({'Property Name' : 'gage_name'}, axis=1, inplace=True)

# 위에서 join한 두 가지 경우의 이름 열 합치기
for i in range(len(gage_data)):
    if pd.isnull(gage_data.iloc[i, 18]):
        gage_data.iloc[i, 18] = gage_data.iloc[i, 19]

gage_data = gage_data.iloc[:, 0:19]


wall_rot_data = wall_rot_data[wall_rot_data['Load Case']\
                              .str.contains('|'.join(seismic_load_name_list))]

### SWR gage data와 SWR result data 연결하기(Element Name 기준으로)
wall_rot_data = wall_rot_data.join(gage_data.set_index('Element Name')['gage_name'], on='Element Name')    
    
### SWR_total data 만들기
SWR_max = wall_rot_data[(wall_rot_data['Step Type'] == 'Max') & (wall_rot_data['Performance Level'] == 1)][['Rotation']].values # dataframe을 array로
SWR_max_gagename = wall_rot_data[(wall_rot_data['Step Type'] == 'Max') & (wall_rot_data['Performance Level'] == 1)][['gage_name']].values # dataframe을 array로
SWR_max = SWR_max.reshape(gage_num, len(DE_load_name_list)+len(MCE_load_name_list), order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
SWR_max_gagename = SWR_max_gagename.reshape(gage_num, len(DE_load_name_list)+len(MCE_load_name_list), order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
SWR_max = pd.DataFrame(SWR_max) # array를 다시 dataframe으로
SWR_max_gagename = pd.DataFrame(SWR_max_gagename) # array를 다시 dataframe으로

SWR_min = wall_rot_data[(wall_rot_data['Step Type'] == 'Min') & (wall_rot_data['Performance Level'] == 1)][['Rotation']].values
SWR_min_gagename = wall_rot_data[(wall_rot_data['Step Type'] == 'Min') & (wall_rot_data['Performance Level'] == 1)][['gage_name']].values
SWR_min = SWR_min.reshape(gage_num, len(DE_load_name_list)+len(MCE_load_name_list), order='F')
SWR_min_gagename = SWR_min_gagename.reshape(gage_num, len(DE_load_name_list)+len(MCE_load_name_list), order='F')
SWR_min = pd.DataFrame(SWR_min)
SWR_min_gagename = pd.DataFrame(SWR_min_gagename)

SWR_total = pd.concat([gage_data['I_V'], SWR_max_gagename.iloc[:,0], SWR_max, SWR_min], axis=1)

#SWR_total 의 column 명 만들기
SWR_total_column_max = []
for load_name in seismic_load_name_list:
    SWR_total_column_max.extend([load_name + '_max'])
    
SWR_total_column_min = []
for load_name in seismic_load_name_list:
    SWR_total_column_min.extend([load_name + '_min'])

SWR_total.columns = ['Height', 'gage_name'] + SWR_total_column_max + SWR_total_column_min

### SWR_avg_data 만들기
DE_max_avg = SWR_total.iloc[:, 2:len(DE_load_name_list)+2].mean(axis=1) # 2를 더해준 건 앞에 Height와 gage_name이 추가되었기 때문
MCE_max_avg = SWR_total.iloc[:, len(DE_load_name_list)+2 : len(DE_load_name_list) + len(MCE_load_name_list)+2].mean(axis=1)
DE_min_avg = SWR_total.iloc[:, len(DE_load_name_list)+len(MCE_load_name_list)+2 : 2*len(DE_load_name_list)+len(MCE_load_name_list)+2].mean(axis=1)
MCE_min_avg = SWR_total.iloc[:, 2*len(DE_load_name_list)+len(MCE_load_name_list)+2 : 2*len(DE_load_name_list) + 2*len(MCE_load_name_list)+2].mean(axis=1)
SWR_avg_total = pd.concat([SWR_total[['Height', 'gage_name']], DE_max_avg, DE_min_avg, MCE_max_avg, MCE_min_avg], axis=1)
SWR_avg_total.columns = ['Height', 'gage_name', 'DE_max_avg', 'DE_min_avg', 'MCE_max_avg', 'MCE_min_avg']   

# LS 기준
deformation_cap_DE = pd.DataFrame()
for i in range(len(deformation_cap)):
    if deformation_cap.iloc[i, 1] > deformation_cap.iloc[i, 2]:
        deformation_cap_DE = pd.concat([deformation_cap_DE, pd.Series(deformation_cap.iloc[i, 7])], ignore_index=True)
    else:
        deformation_cap_DE = pd.concat([deformation_cap_DE, pd.Series(deformation_cap.iloc[i, 8])], ignore_index=True)

# CP 기준
deformation_cap_MCE = pd.DataFrame()
for i in range(len(deformation_cap)):
    if deformation_cap.iloc[i, 3] > deformation_cap.iloc[i, 4]:
        deformation_cap_MCE = pd.concat([deformation_cap_MCE, pd.Series(deformation_cap.iloc[i, 9])], ignore_index=True)
    else:
        deformation_cap_MCE = pd.concat([deformation_cap_MCE, pd.Series(deformation_cap.iloc[i, 10])], ignore_index=True)

SWR_criteria = pd.concat([deformation_cap['Name'], deformation_cap['IO(H1)']
                          ,deformation_cap_DE, deformation_cap_MCE]
                         , axis = 1, ignore_index=True)
SWR_criteria.columns = ['Name','IO Criteria', 'LS criteria', 'CP criteria']

#%%
# H1, H2 방향 전단력 중 큰 값만 get
results_wall['Vu_DE(kN)'] = results_wall[['Vu_DE_H1(kN)', 'Vu_DE_H2(kN)']].max(axis=1)
results_wall['Vu_MCE(kN)'] = results_wall[['Vu_MCE_H1(kN)', 'Vu_MCE_H2(kN)']].max(axis=1)

results_wall = pd.concat([results_wall.iloc[:,[0,1,2,3,5,6,7,8,9,10,16,17,15]]
                          , SWR_criteria.iloc[:,[1,2,3]]], axis=1)

SWR_avg_total.dropna(subset='gage_name', inplace=True)

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
SWR_avg_total = pd.merge(SWR_avg_total, results_wall, how='left')

SWR_avg_total.dropna(subset='IO Criteria', inplace=True)

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

SWR_avg_total['Rotation_DE'] = np.maximum(SWR_avg_total['DE_min_avg'].abs()
                                          , SWR_avg_total['DE_max_avg'])
SWR_avg_total['SWR_DCR_DE'] = determine_cap(SWR_avg_total['Rotation_DE']
                                            , SWR_avg_total.iloc[:,[18,19,20]])

SWR_avg_total['Rotation_MCE'] = np.maximum(SWR_avg_total['MCE_min_avg'].abs()
                                          , SWR_avg_total['MCE_max_avg'])
SWR_avg_total['SWR_DCR_MCE'] = determine_cap(SWR_avg_total['Rotation_MCE']
                                             , SWR_avg_total.iloc[:,[18,19,20]])

#%% 배근 표기 combine
v_rebar_combined = []
h_rebar_combined = []
for idx, row in SWR_avg_total.iloc[:,[9,10,11,12,13]].iterrows():
    # vertical rebar
    if np.isnan(row[1]) == False:
        v_rebar_combined.append(row[0]+'@'+str(int(row[1])))
    elif np.isnan(row[1]) == True:
        v_rebar_combined.append(str(int(row[2]))+'-'+row[0])
    # horizontal rebar
    h_rebar_combined.append(row[3]+'@'+str(int(row[4])))

SWR_avg_total['V.Rebar_Combined'] = v_rebar_combined
SWR_avg_total['H.Rebar_Combined'] = h_rebar_combined

# 이름 index에 따라 정렬
SWR_avg_total = pd.merge(SWR_avg_total, name_cri, how='left')
SWR_avg_total.sort_values(by='Index', ascending=True, inplace=True)
SWR_avg_total.reset_index(inplace=True, drop=True)

SWR_output = SWR_avg_total.iloc[:,[1,6,7,8,25,26,18,19,20,21,22,23,24,14,15,16,17]]

#%% SF DCR 계산 및 입력
def determine_DCR(value_list):
    value_arr = np.array(value_list)
    DCR_list = []
    for i in value_arr:
        if i >= 1:
            DCR_list.append('NG')
        else: DCR_list.append('OK')
    return DCR_list

SWR_output['SF_DCR_DE'] = SWR_output['Vu_DE(kN)'] / SWR_output['Vn(kN)']
SWR_output['SF_Results_DE'] = determine_DCR(SWR_output['SF_DCR_DE'])

SWR_output['SF_DCR_MCE'] = SWR_output['Vu_MCE(kN)'] / SWR_output['Vn(kN)']
SWR_output['SF_Results_MCE'] = determine_DCR(SWR_output['SF_DCR_MCE'])


#%% Output

# 소수점 자리 설정
SWR_output.iloc[:,[6,7,8,9,11,17,19]] = SWR_output.iloc[:,[6,7,8,9,11,17,19]].round(4)
SWR_output.iloc[:,[13,14,15,16]] = SWR_output.iloc[:,[13,14,15,16]].round(2)

# 111동 부재명에서 R 떼기
# remove_R_list = []
# for i in SWR_output.iloc[:,0]:
#     remove_R = i[1:]
#     remove_R_list.append(remove_R)
    
# SWR_output.iloc[:,0] = remove_R_list