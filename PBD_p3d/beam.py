import pandas as pd
import numpy as np
import os
from collections import deque  # Double-ended Queue : 자료의 앞, 뒤 양 방향에서 자료를 추가하거나 제거가능
import matplotlib.pyplot as plt
import matplotlib as mpl
import win32com.client

#%% Beam Rotation

def BR(result_path, result_xlsx, input_path, input_xlsx,\
       m_hinge_group_name, **kwargs): 
    # arguments가 너무 많을 때, 함수를 사용할 때 직접 매개변수를 명시해주는 방식

#%% 변수 정리
    s_hinge_group_name = kwargs['s_hinge_group_name'] if 's_hinge_group_name' in kwargs.keys() else None
    m_cri_DE = kwargs['moment_cri_DE'] if 'moment_cri_DE' in kwargs.keys() else 0.015/1.2
    m_cri_MCE = kwargs['moment_cri_MCE'] if 'moment_cri_MCE' in kwargs.keys() else 0.03/1.2
    s_cri_DE = kwargs['shear_cri_DE'] if 'shear_cri_DE' in kwargs.keys() else 0.015/1.2
    s_cri_MCE = kwargs['shear_cri_MCE'] if 'shear_cri_MCE' in kwargs.keys() else 0.03/1.2
    yticks = kwargs['yticks'] if 'yticks' in kwargs.keys() else 3
    xlim = kwargs['xlim'] if 'xlim' in kwargs.keys() else 0.03

#%% Analysis Result 불러오기(BR)
    to_load_list = []
    file_names = os.listdir(result_path)
    for file_name in file_names:
        if (result_xlsx in file_name) and ('~$' not in file_name):
            to_load_list.append(file_name)
    
    # Gage data
    gage_data = pd.read_excel(result_path + '\\' + to_load_list[0],
                                    sheet_name='Gage Data - Beam Type', skiprows=[0, 2], header=0, usecols=[0, 2, 7, 9]) # usecols로 원하는 열만 불러오기
    
    BR_M_gage_data = gage_data[gage_data['Group Name'] == m_hinge_group_name]
    BR_S_gage_data = gage_data[gage_data['Group Name'] == s_hinge_group_name]
    
    
    # Gage result data
    result_data = pd.DataFrame()
    for i in to_load_list:
        result_data_temp = pd.read_excel(result_path + '\\' + i,
                                    sheet_name='Gage Results - Beam Type', skiprows=[0, 2], header=0, usecols=[0, 2, 5, 7, 8, 9])
        result_data = pd.concat([result_data, result_data_temp])
    
    result_data = result_data.sort_values(by=['Load Case', 'Element Name', 'Step Type']) # 지진파 순서가 섞여있을 때 sort
    
    result_data = result_data[(result_data['Load Case'].str.contains('DE')) | (result_data['Load Case'].str.contains('MCE'))]
    
    BR_M_result_data = result_data[result_data['Group Name'] == m_hinge_group_name]
    BR_S_result_data = result_data[result_data['Group Name'] == s_hinge_group_name]
    
#%% Node, Story 정보 불러오기
    
    # Node Coord data
    Node_coord_data = pd.read_excel(result_path + '\\' + to_load_list[0],
                                    sheet_name='Node Coordinate Data', skiprows=[0, 2], header=0, usecols=[1, 2, 3, 4])
    
    # Story Info data
    story_info_xlsx_sheet = 'Story Data'
    story_info = pd.read_excel(input_path + '\\' + input_xlsx, sheet_name=story_info_xlsx_sheet, skiprows=3, usecols=[0, 1, 2], keep_default_na=False)
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    story_name = story_info.loc[:, 'Story Name']

#%% 지진파 이름 list 만들기
    load_name_list = []
    for i in result_data['Load Case'].drop_duplicates():
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
    BR_M_gage_data = BR_M_gage_data[['Element Name', 'I-Node ID']]; BR_M_gage_num = len(BR_M_gage_data) # gage 개수 얻기
    BR_S_gage_data = BR_S_gage_data[['Element Name', 'I-Node ID']]; BR_S_gage_num = len(BR_S_gage_data) # gage 개수 얻기
    
    # I-Node의 v좌표 match해서 추가
    gage_data = gage_data.join(Node_coord_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')
    BR_M_gage_data = BR_M_gage_data.join(Node_coord_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')
    BR_S_gage_data = BR_S_gage_data.join(Node_coord_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')
    BR_S_gage_data.reset_index(drop=True, inplace=True)
    
    ### BR_total data 만들기
    BR_M_max = BR_M_result_data[(BR_M_result_data['Step Type'] == 'Max') & (BR_M_result_data['Performance Level'] == 1)][['Rotation']].values # dataframe을 array로
    BR_M_max = BR_M_max.reshape(BR_M_gage_num, len(DE_load_name_list)+len(MCE_load_name_list), order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
    BR_M_max = pd.DataFrame(BR_M_max) # array를 다시 dataframe으로
    BR_M_min = BR_M_result_data[(BR_M_result_data['Step Type'] == 'Min') & (BR_M_result_data['Performance Level'] == 1)][['Rotation']].values
    BR_M_min = BR_M_min.reshape(BR_M_gage_num, len(DE_load_name_list)+len(MCE_load_name_list), order='F')
    BR_M_min = pd.DataFrame(BR_M_min)
    BR_M_total = pd.concat([BR_M_max, BR_M_min], axis=1) # DE11_max~MCE72_max, DE11_min~MCE72_min 각각 28개씩

    BR_S_max = BR_S_result_data[(BR_S_result_data['Step Type'] == 'Max') & (BR_S_result_data['Performance Level'] == 1)][['Rotation']].values # dataframe을 array로
    BR_S_max = BR_S_max.reshape(BR_S_gage_num, len(DE_load_name_list)+len(MCE_load_name_list), order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
    BR_S_max = pd.DataFrame(BR_S_max) # array를 다시 dataframe으로
    BR_S_min = BR_S_result_data[(BR_S_result_data['Step Type'] == 'Min') & (BR_S_result_data['Performance Level'] == 1)][['Rotation']].values
    BR_S_min = BR_S_min.reshape(BR_S_gage_num, len(DE_load_name_list)+len(MCE_load_name_list), order='F')
    BR_S_min = pd.DataFrame(BR_S_min)
    BR_S_total = pd.concat([BR_S_max, BR_S_min], axis=1) # DE11_max~MCE72_max, DE11_min~MCE72_min 각각 28개씩
            
    ### BR_avg_data 만들기
    BR_M_DE_max_avg = BR_M_total.iloc[:, 0:len(DE_load_name_list)].mean(axis=1)
    BR_M_MCE_max_avg = BR_M_total.iloc[:, len(DE_load_name_list) : len(DE_load_name_list)+len(MCE_load_name_list)].mean(axis=1)
    BR_M_DE_min_avg = BR_M_total.iloc[:, len(DE_load_name_list)+len(MCE_load_name_list) : 2*len(DE_load_name_list)+len(MCE_load_name_list)].mean(axis=1)
    BR_M_MCE_min_avg = BR_M_total.iloc[:, 2*len(DE_load_name_list)+len(MCE_load_name_list) : 2*len(DE_load_name_list) + 2*len(MCE_load_name_list)].mean(axis=1)
    BR_M_avg_total = pd.concat([BR_M_gage_data.iloc[:,[2,3,4]], BR_M_DE_max_avg, BR_M_DE_min_avg, BR_M_MCE_max_avg, BR_M_MCE_min_avg], axis=1)
    BR_M_avg_total.columns = ['X', 'Y', 'Height', 'DE_max_avg', 'DE_min_avg', 'MCE_max_avg', 'MCE_min_avg']

    BR_S_DE_max_avg = BR_S_total.iloc[:, 0:len(DE_load_name_list)].mean(axis=1)
    BR_S_MCE_max_avg = BR_S_total.iloc[:, len(DE_load_name_list) : len(DE_load_name_list)+len(MCE_load_name_list)].mean(axis=1)
    BR_S_DE_min_avg = BR_S_total.iloc[:, len(DE_load_name_list)+len(MCE_load_name_list) : 2*len(DE_load_name_list)+len(MCE_load_name_list)].mean(axis=1)
    BR_S_MCE_min_avg = BR_S_total.iloc[:, 2*len(DE_load_name_list)+len(MCE_load_name_list) : 2*len(DE_load_name_list) + 2*len(MCE_load_name_list)].mean(axis=1)
    BR_S_avg_total = pd.concat([BR_S_gage_data.iloc[:,[2,3,4]], BR_S_DE_max_avg, BR_S_DE_min_avg, BR_S_MCE_max_avg, BR_S_MCE_min_avg], axis=1, ignore_index=True)
    BR_S_avg_total.columns = ['X', 'Y', 'Height', 'DE_max_avg', 'DE_min_avg', 'MCE_max_avg', 'MCE_min_avg']
    
#%% ***조작용 코드
    # 없애고 싶은 부재의 x좌표 입력
    # BR_M_avg_total = BR_M_avg_total.drop(BR_M_avg_total[(BR_M_avg_total['X'] == -2700)].index)
    # BR_M_avg_total = BR_M_avg_total.drop(BR_M_avg_total[(BR_M_avg_total['X'] == -6.1e-05)].index)
    # BR_M_avg_total = BR_M_avg_total.drop(BR_M_avg_total[(BR_M_avg_total['X'] == -4725)].index)
    
#%% BR (Moment Hinge) 그래프

    count = 1

    if BR_M_avg_total.shape[0] != 0:

# DE 그래프        
        if len(DE_load_name_list) != 0:
        
            fig1 = plt.figure(count, figsize=(4,5), dpi=150)  # 그래프 사이즈
            plt.xlim(-xlim, xlim)
            
            plt.scatter(BR_M_avg_total['DE_min_avg'], BR_M_avg_total['Height'], color = 'k', s=1) # s=1 : point size
            plt.scatter(BR_M_avg_total['DE_max_avg'], BR_M_avg_total['Height'], color = 'k', s=1)
            
            # height값에 대응되는 층 이름으로 y축 눈금 작성
            plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
            
            # reference line 그려서 허용치 나타내기
            plt.axvline(x= -m_cri_DE, color='r', linestyle='--')
            plt.axvline(x= m_cri_DE, color='r', linestyle='--')
            
            plt.grid(linestyle='-.')
            plt.xlabel('Rotation(rad)')
            plt.ylabel('Story')
            plt.title('DE (Moment Hinge)')
            
            plt.close()
            count += 1
            
            error_coord_M_DE = BR_M_avg_total[(BR_M_avg_total['DE_max_avg'] >= m_cri_DE)\
                                              | (BR_M_avg_total['DE_min_avg'] <= -m_cri_DE)]
                
            yield fig1
            yield error_coord_M_DE

# MCE 그래프    
        if len(MCE_load_name_list) != 0:
    
            fig2 = plt.figure(count, figsize=(4,5), dpi=150)
            plt.xlim(-xlim, xlim)
            plt.scatter(BR_M_avg_total['MCE_min_avg'], BR_M_avg_total['Height'], color = 'k', s=1)
            plt.scatter(BR_M_avg_total['MCE_max_avg'], BR_M_avg_total['Height'], color = 'k', s=1)
            
            plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
            
            plt.axvline(x= -m_cri_MCE, color='r', linestyle='--')
            plt.axvline(x= m_cri_MCE, color='r', linestyle='--')
            
            plt.grid(linestyle='-.')
            plt.xlabel('Rotation(rad)')
            plt.ylabel('Story')
            plt.title('MCE (Moment Hinge)')
            
            plt.close()
            count += 1
            
            error_coord_M_MCE = BR_M_avg_total[(BR_M_avg_total['MCE_max_avg'] >= m_cri_MCE)\
                                               | (BR_M_avg_total['MCE_min_avg'] <= -m_cri_MCE)]
    
            yield fig2
            yield error_coord_M_MCE
            
#%% BR(Shear Hinge) 그래프

    if BR_S_avg_total.shape[0] != 0:

# DE 그래프        
        if len(DE_load_name_list) != 0:

            fig3 = plt.figure(count, figsize=(4,5), dpi=150)  # 그래프 사이즈
            plt.xlim(-xlim, xlim)
            
            plt.scatter(BR_S_avg_total['DE_min_avg'], BR_S_avg_total['Height'], color = 'k', s=1) # s=1 : point size
            plt.scatter(BR_S_avg_total['DE_max_avg'], BR_S_avg_total['Height'], color = 'k', s=1)
            
            # height값에 대응되는 층 이름으로 y축 눈금 작성
            plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
            
            # reference line 그려서 허용치 나타내기
            # plt.axvline(x= -s_cri_DE, color='r', linestyle='--')
            # plt.axvline(x= s_cri_DE, color='r', linestyle='--')
            
            plt.grid(linestyle='-.')
            plt.xlabel('Rotation(rad)')
            plt.ylabel('Story')
            plt.title('DE (Shear Hinge)')
            
            plt.close()
            count += 1
            
            error_coord_S_DE = BR_S_avg_total[(BR_S_avg_total['DE_max_avg'] >= s_cri_DE)\
                                              | (BR_S_avg_total['DE_min_avg'] <= -s_cri_DE)]
    
            yield fig3
            yield error_coord_S_DE

# MCE 그래프    
        if len(MCE_load_name_list) != 0:
            
            fig4 = plt.figure(count, figsize=(4,5), dpi=150)
            plt.xlim(-xlim, xlim)
            plt.scatter(BR_S_avg_total['MCE_min_avg'], BR_S_avg_total['Height'], color = 'k', s=1)
            plt.scatter(BR_S_avg_total['MCE_max_avg'], BR_S_avg_total['Height'], color = 'k', s=1)
            
            plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
            
            # plt.axvline(x= -s_cri_MCE, color='r', linestyle='--')
            # plt.axvline(x= s_cri_MCE, color='r', linestyle='--')
            
            plt.grid(linestyle='-.')
            plt.xlabel('Rotation(rad)')
            plt.ylabel('Story')
            plt.title('MCE (Shear Hinge)')
            
            plt.close()
            count += 1
            
            error_coord_S_MCE = BR_S_avg_total[(BR_S_avg_total['MCE_max_avg'] >= s_cri_MCE)\
                                               | (BR_S_avg_total['MCE_min_avg'] <= -s_cri_MCE)]

            yield fig4
            yield error_coord_S_MCE

#%% Beam Rotation (DCR)
def BR_DCR(result_path, result_xlsx, input_path, input_xlsx
           , c_beam_group='C.Beam', DCR_criteria=1, yticks=3, xlim=3):

#%% Input Sheets 정보 load
    story_info = pd.DataFrame()
    deformation_cap = pd.DataFrame()
    
    input_data_raw = pd.ExcelFile(input_path + '\\' + input_xlsx)
    input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', 'Output_C.Beam Properties'], skiprows=3)
    input_data_raw.close()
    
    story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
    deformation_cap = input_data_sheets['Output_C.Beam Properties'].iloc[:,[0,48,49]]
    
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    deformation_cap.columns = ['Name', 'LS', 'CP']
    
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
                                                         'Element Data - Frame Types'], skiprows=2)
        
        beam_rot_data_temp = result_data_sheets['Frame Results - Bending Deform'].iloc[:,[0,2,5,7,10,13,14]]
        beam_rot_data = pd.concat([beam_rot_data, beam_rot_data_temp])
        
    node_data = result_data_sheets['Node Coordinate Data'].iloc[:,[1,4]]
    
    element_data = result_data_sheets['Element Data - Frame Types'].iloc[:,[2,5,7]] # beam의 양 nodes중 한 node에서의 rotation * 2
    
                
    beam_rot_data.columns = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'Distance from I-End', 'H2 Rotation(rad)', 'H3 Rotation(rad)']
    node_data.columns = ['Node ID', 'V(mm)']
    element_data.columns = ['Element Name', 'Property Name', 'I-Node ID']
    
#%% temporary ((L), (R) 등 지우기)
    element_data.loc[:, 'Property Name'] = element_data.loc[:, 'Property Name'].str.split('(').str[0]
    
    #%% 필요없는 부재 빼기, 필요한 부재만 추출
    beam_rot_data = beam_rot_data[beam_rot_data['Group Name'] == c_beam_group]
    beam_rot_data = beam_rot_data[beam_rot_data['Distance from I-End'] == 0]
    
#%% Analysis Result에 Element, Node 정보 매칭
    
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
    
    DE_load_name_list = [x for x in load_name_list if 'DE' in x]
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]
    
#%% 성능기준(LS, CP) 정리해서 merge
    
    beam_rot_data = pd.merge(beam_rot_data, deformation_cap, how='left', left_on='Property Name', right_on='Name')
    
    beam_rot_data['DE Rotation(rad)'] = beam_rot_data['Major Rotation(rad)'].abs() / beam_rot_data['LS']
    beam_rot_data['MCE Rotation(rad)'] = beam_rot_data['Major Rotation(rad)'].abs() / beam_rot_data['CP']
    
    beam_rot_data = beam_rot_data[beam_rot_data['Name'].notna()]
    
#%% 조작용 코드
    # 없애고 싶은 부재의 이름 입력(error_beam 확인 후!, DE, MCE에서 다 없어짐)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('LB4_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('B15_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB4A_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB4B_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB3D_'))].index)

#%% DE 결과 Plot
    count = 1
    
    if len(DE_load_name_list) != 0:
        
        beam_rot_data_total_DE = pd.DataFrame()    
        
        for load_name in DE_load_name_list:
        
            temp_df_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['DE Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']
                          
            beam_rot_data_total_DE['{}_max'.format(load_name)] = temp_df_max.tolist()
            
            temp_df_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['DE Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']
                          
            beam_rot_data_total_DE['{}_min'.format(load_name)] = temp_df_min.tolist()
            
        beam_rot_data_total_DE['Element Name'] = temp_df_max.index
        
        beam_rot_data_total_DE.reset_index(inplace=True, drop=True)
        
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, element_data, how='left')
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, story_info, how='left', left_on='V(mm)', right_on='Height(mm)')
        beam_rot_data_total_DE.sort_values('Height(mm)', inplace=True)
        # beam_rot_data_total_DE.reset_index(inplace=True, drop=True)
        
    # 평균 열 생성
        
        beam_rot_data_total_DE['DE Max avg'] = beam_rot_data_total_DE.iloc[:,list(range(0,len(DE_load_name_list)*2,2))].mean(axis=1)
        beam_rot_data_total_DE['DE Min avg'] = beam_rot_data_total_DE.iloc[:,list(range(1,len(DE_load_name_list)*2,2))].mean(axis=1)
        
    # 전체 Plot
            
        ### DE 
        fig1 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        plt.scatter(beam_rot_data_total_DE['DE Max avg'], beam_rot_data_total_DE.loc[:,'V(mm)'], color='k', s=1)
        plt.scatter(beam_rot_data_total_DE['DE Min avg'], beam_rot_data_total_DE.loc[:,'V(mm)'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Beam Rotation (DE)')
        
        plt.tight_layout()   
        plt.close()

    # 기준 넘는 점 확인
        error_beam_DE = beam_rot_data_total_DE[['Element Name', 'Property Name', 'Story Name', 'DE Max avg', 'DE Min avg']]\
                      [(beam_rot_data_total_DE['DE Max avg'] >= DCR_criteria) | (beam_rot_data_total_DE['DE Min avg'] >= DCR_criteria)]
        
        count += 1
        
        yield fig1
        yield error_beam_DE
        
#%% MCE 결과 Plot
    
    if len(MCE_load_name_list) != 0:
        
        beam_rot_data_total_MCE = pd.DataFrame()    
        
        for load_name in MCE_load_name_list:
        
            temp_df_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['MCE Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']
                          
            beam_rot_data_total_MCE['{}_max'.format(load_name)] = temp_df_max.tolist()
            
            temp_df_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['MCE Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']
                          
            beam_rot_data_total_MCE['{}_min'.format(load_name)] = temp_df_min.tolist()
            
        beam_rot_data_total_MCE['Element Name'] = temp_df_max.index
        
        beam_rot_data_total_MCE.reset_index(inplace=True, drop=True)
        
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, element_data, how='left')
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, story_info, how='left', left_on='V(mm)', right_on='Height(mm)')
        beam_rot_data_total_MCE.sort_values('Height(mm)', inplace=True)
        # beam_rot_data_total_MCE.reset_index(inplace=True, drop=True)
        
    # 평균 열 생성
        
        beam_rot_data_total_MCE['MCE Max avg'] = beam_rot_data_total_MCE.iloc[:,list(range(0,len(MCE_load_name_list)*2,2))].mean(axis=1)
        beam_rot_data_total_MCE['MCE Min avg'] = beam_rot_data_total_MCE.iloc[:,list(range(1,len(MCE_load_name_list)*2,2))].mean(axis=1)
            

        # 전체 Plot 
        ### MCE 
        fig2 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        # 평균 plot
        plt.scatter(beam_rot_data_total_MCE['MCE Max avg'], beam_rot_data_total_MCE.loc[:,'V(mm)'], color='k', s=1)
        plt.scatter(beam_rot_data_total_MCE['MCE Min avg'], beam_rot_data_total_MCE.loc[:,'V(mm)'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Beam Rotation (MCE)')
        
        plt.tight_layout()
        plt.close()
    
        # 기준 넘는 점 확인    
        error_beam_MCE = beam_rot_data_total_MCE[['Element Name', 'Property Name', 'Story Name', 'MCE Max avg', 'MCE Min avg']]\
                      [(beam_rot_data_total_MCE['MCE Max avg'] >= DCR_criteria) | (beam_rot_data_total_MCE['MCE Min avg'] >= DCR_criteria)]
        
        yield fig2
        yield error_beam_MCE
        
#%% Return! (지진파가 다 없는 경우도 고려함)
    # if 'fig1' in locals():
    #     if 'fig2' in locals():
    #         return fig1, fig2, error_beam_DE, error_beam_MCE
        
    #     elif 'fig2' not in locals():
    #         return fig1, error_beam_DE
        
    # elif 'fig1' not in locals():
    #     if 'fig2' in locals():
    #         return fig2, error_beam_MCE

#%% Transfer Beam SF (DCR)

def trans_beam_SF(result_path, result_xlsx, input_path, input_xlsx):

#%% Input Sheet 정보 load
        
    story_info = pd.DataFrame()
    transfer_element_info = pd.DataFrame()

    input_xlsx_sheet = 'Output_E.Beam Properties'
    input_data_raw = pd.ExcelFile(input_path + '\\' + input_xlsx)
    input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', input_xlsx_sheet], skiprows=3)
    input_data_raw.close()

    story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
    transfer_element_info = input_data_sheets[input_xlsx_sheet].iloc[:,0]
    story_info = story_info[::-1]
    story_info.reset_index(inplace=True, drop=True)

    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    transfer_element_info.name = 'Name'

    #%% Analysis Result 불러오기

    to_load_list = []
    file_names = os.listdir(result_path)
    for file_name in file_names:
        if (result_xlsx in file_name) and ('~$' not in file_name):
            to_load_list.append(file_name)

    # 전단력 Data
    SF_info_data = pd.DataFrame()
    for i in to_load_list:
        SF_info_data_temp = pd.read_excel(result_path + '\\' + i,
                                   sheet_name='Frame Results - End Forces', skiprows=[0, 2], header=0, usecols=[0,2,5,7,10,11,17,18]) # usecols로 원하는 열만 불러오기
        SF_info_data = pd.concat([SF_info_data, SF_info_data_temp])

    SF_info_data = SF_info_data.sort_values(by=['Load Case', 'Element Name', 'Step Type']) # 지진파 순서가 섞여있을 때 sort

    # 부재 이름 Matching을 위한 Element 정보
    element_info_data = pd.DataFrame()
    for i in to_load_list:
        element_info_data_temp = pd.read_excel(result_path + '\\' + i,
                                   sheet_name='Element Data - Frame Types', skiprows=[0, 2], header=0, usecols=[0, 2, 5, 7]) # usecols로 원하는 열만 불러오기
        element_info_data = pd.concat([element_info_data, element_info_data_temp])

    # 필요한 부재만 선별
    element_info_data = element_info_data[element_info_data['Property Name'].isin(transfer_element_info)]
    
    # 층 정보 Matching을 위한 Node 정보
    height_info_data = pd.DataFrame()    
    for i in to_load_list:
        height_info_data_temp = pd.read_excel(result_path + '\\' + i,
                                   sheet_name='Node Coordinate Data', skiprows=[0, 2], header=0, usecols=[1, 4]) # usecols로 원하는 열만 불러오기
        height_info_data = pd.concat([height_info_data, height_info_data_temp])

    element_info_data = pd.merge(element_info_data, height_info_data, how='left', left_on='I-Node ID', right_on='Node ID')

    element_info_data = element_info_data.drop_duplicates()

    # 전단력, 부재 이름 Matching (by Element Name)
    SF_ongoing = pd.merge(element_info_data.iloc[:, [1,2,5]], SF_info_data.iloc[:, 1:], how='left')

    SF_ongoing = SF_ongoing.sort_values(by=['Element Name', 'Load Case', 'Step Type'])

    SF_ongoing.reset_index(inplace=True, drop=True)

    #%% 지진파 이름 list 만들기

    load_name_list = []
    for i in SF_ongoing['Load Case'].drop_duplicates():
        new_i = i.split('+')[1]
        new_i = new_i.strip()
        load_name_list.append(new_i)

    gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
    seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]

    seismic_load_name_list.sort()

    DE_load_name_list = [x for x in load_name_list if 'DE' in x]
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]

    #%% V, M값에 절대값, 최대값, 평균값 뽑기

    # 절대값, 1.2배
    SF_ongoing.iloc[:,[5,6,7,8]] = SF_ongoing.iloc[:,[5,6,7,8]].abs() * 1.2

    # i, j 노드 중 최대값 뽑기
    SF_ongoing['V2 max'] = SF_ongoing[['V2 I-End', 'V2 J-End']].max(axis = 1)
    SF_ongoing['M3 max'] = SF_ongoing[['M3 I-End', 'M3 J-End']].max(axis = 1)

    # max, min 중 최대값 뽑기
    SF_ongoing_max = SF_ongoing.loc[SF_ongoing.groupby(SF_ongoing.index // 2)['V2 max'].idxmax()]
    SF_ongoing_max['M3 max'] = SF_ongoing.groupby(SF_ongoing.index // 2)['M3 max'].max().tolist()

    # 필요한 하중만 포함된 데이터 slice (MCE)
    SF_ongoing_max = SF_ongoing_max[SF_ongoing_max['Load Case']\
                                    .str.contains('|'.join(MCE_load_name_list))] # function equivalent of a combination of df.isin() and df.str.contains()
    
    # 부재별 평균값 뽑기
    SF_ongoing_max_avg = SF_ongoing_max.iloc[:,[0,1,2]]
    SF_ongoing_max_avg = SF_ongoing_max_avg.drop_duplicates()
    SF_ongoing_max_avg.set_index('Element Name', inplace=True)
    
    SF_ongoing_max_avg['V2 max'] = SF_ongoing_max.groupby(['Element Name'])['V2 max'].mean()
    SF_ongoing_max_avg['M3 max'] = SF_ongoing_max.groupby(['Element Name'])['M3 max'].mean()
 
    # 같은 부재(그러나 잘려있는) 경우 최대값 뽑기
    SF_ongoing_max_avg_max = SF_ongoing_max_avg.loc[SF_ongoing_max_avg.groupby(['Property Name'])['V2 max'].idxmax()]
    SF_ongoing_max_avg_max['M3 max'] = SF_ongoing_max_avg.groupby(['Property Name'])['M3 max'].max().tolist()
    
    SF_ongoing_max_avg_max.reset_index(inplace=True, drop=True)

#%% 결과값 정리 후 input sheets에 넣기
    
    SF_ongoing_max_avg_max = pd.merge(transfer_element_info.rename('Property Name'),\
                                      SF_ongoing_max_avg_max, how='left')
        
    SF_ongoing_max_avg_max = SF_ongoing_max_avg_max.dropna()
    SF_ongoing_max_avg_max.reset_index(inplace=True, drop=True)
    
    # SF_ongoing_max_avg 재정렬
    SF_output = SF_ongoing_max_avg_max.iloc[:,[0,2,3]] 
    
    # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
    SF_output = SF_output.replace(np.nan, '', regex=True)
    
    # Using win32com...
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application') # 엑셀 실행
    excel.Visible = False # 엑셀창 안보이게

    wb = excel.Workbooks.Open(input_path + '\\' + input_xlsx)
    ws = wb.Sheets('Results_E.Beam')
    
    startrow, startcol = 5, 1
    
    ws.Range(ws.Cells(startrow, startcol),\
             ws.Cells(startrow + SF_output.shape[0]-1,\
                      startcol + SF_output.shape[1]-1)).Value\
    = list(SF_output.itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능
    
    wb.Close(SaveChanges=1) # Closing the workbook
    excel.Quit() # Closing the application
    
    print('Done!')

#%% Transfer Beam SF (DCR) ------ revising

def trans_beam_SF_2(result_path, result_xlsx, input_path, input_xlsx, beam_xlsx, contour=True):

#%% Input Sheet 정보 load
        
    story_info = pd.DataFrame()
    transfer_element_info = pd.DataFrame()

    input_xlsx_sheet = 'Output_E.Beam Properties'
    input_data_raw = pd.ExcelFile(input_path + '\\' + input_xlsx)
    input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', input_xlsx_sheet], skiprows=3)
    # input_data_raw.close()

    story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
    transfer_element_info = input_data_sheets[input_xlsx_sheet].iloc[:,0]
    story_info = story_info[::-1]
    story_info.reset_index(inplace=True, drop=True)

    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    transfer_element_info.name = 'Name'

    #%% Analysis Result 불러오기

    to_load_list = []
    file_names = os.listdir(result_path)
    for file_name in file_names:
        if (result_xlsx in file_name) and ('~$' not in file_name):
            to_load_list.append(file_name)

    # 전단력 Data
    SF_info_data = pd.DataFrame()
    for i in to_load_list:
        SF_info_data_temp = pd.read_excel(result_path + '\\' + i,
                                   sheet_name='Frame Results - End Forces', skiprows=[0, 2], header=0, usecols=[0,2,5,7,10,11,17,18]) # usecols로 원하는 열만 불러오기
        SF_info_data = pd.concat([SF_info_data, SF_info_data_temp])

    SF_info_data = SF_info_data.sort_values(by=['Load Case', 'Element Name', 'Step Type']) # 지진파 순서가 섞여있을 때 sort

    # 부재 이름 Matching을 위한 Element 정보
    element_info_data = pd.DataFrame()
    for i in to_load_list:
        element_info_data_temp = pd.read_excel(result_path + '\\' + i,
                                   sheet_name='Element Data - Frame Types', skiprows=[0, 2], header=0, usecols=[0,2,5,7,9]) # usecols로 원하는 열만 불러오기
        element_info_data = pd.concat([element_info_data, element_info_data_temp])

    # 필요한 부재만 선별
    element_info_data = element_info_data[element_info_data['Property Name'].isin(transfer_element_info)]
    
    # 기둥과 겹치는 등 평가에 반영하지 않을 부재 제거
    element_to_remove = ['E880','E26229','E555','E671','E658','E525','E528','E932','E914','E1256','E1165','E585']
    element_info_data = element_info_data[~element_info_data['Element Name'].isin(element_to_remove)]
    
    # 층 정보 Matching을 위한 Node 정보
    node_info_data = pd.DataFrame()    
    for i in to_load_list:
        node_info_data_temp = pd.read_excel(result_path + '\\' + i,
                                   sheet_name='Node Coordinate Data', skiprows=[0, 2], header=0, usecols=[1,2,3,4]) # usecols로 원하는 열만 불러오기
        node_info_data = pd.concat([node_info_data, node_info_data_temp])
    
    # 나중에 element_info_data 열이름 깔끔하게 하기 위해 미리 깔끔하게
    i_node_info_data, j_node_info_data = node_info_data.copy(), node_info_data.copy()
    i_node_info_data.columns = ['I-Node ID', 'i-H1', 'i-H2', 'i-V']
    j_node_info_data.columns = ['J-Node ID', 'j-H1', 'j-H2', 'j-V']
    
    element_info_data = pd.merge(element_info_data, i_node_info_data, how='left')
    element_info_data = pd.merge(element_info_data, j_node_info_data, how='left')
    
    element_info_data = element_info_data.drop_duplicates()
    
    # 전단력, 부재 이름 Matching (by Element Name)
    SF_ongoing = pd.merge(element_info_data.iloc[:, [1,2,7]], SF_info_data.iloc[:, 1:], how='left')
    
    SF_ongoing = SF_ongoing.sort_values(by=['Element Name', 'Load Case', 'Step Type'])
    
    SF_ongoing.reset_index(inplace=True, drop=True)

    #%% 지진파 이름 list 만들기

    load_name_list = []
    for i in SF_ongoing['Load Case'].drop_duplicates():
        new_i = i.split('+')[1]
        new_i = new_i.strip()
        load_name_list.append(new_i)

    gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
    seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]

    seismic_load_name_list.sort()

    DE_load_name_list = [x for x in load_name_list if 'DE' in x]
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]

    #%% V, M값에 절대값, 최대값, 평균값 뽑기

    # 절대값, 1.2배
    SF_ongoing.iloc[:,[5,6,7,8]] = SF_ongoing.iloc[:,[5,6,7,8]].abs() * 1.2

    # i, j 노드 중 최대값 뽑기
    SF_ongoing['V2 max'] = SF_ongoing[['V2 I-End', 'V2 J-End']].max(axis = 1)
    SF_ongoing['M3 max'] = SF_ongoing[['M3 I-End', 'M3 J-End']].max(axis = 1)

    # max, min 중 최대값 뽑기
    SF_ongoing_max = SF_ongoing.loc[SF_ongoing.groupby(SF_ongoing.index // 2)['V2 max'].idxmax()]
    SF_ongoing_max['M3 max'] = SF_ongoing.groupby(SF_ongoing.index // 2)['M3 max'].max().tolist()

    # 필요한 하중만 포함된 데이터 slice (MCE)
    SF_ongoing_max = SF_ongoing_max[SF_ongoing_max['Load Case']\
                                    .str.contains('|'.join(MCE_load_name_list))] # function equivalent of a combination of df.isin() and df.str.contains()
    
    # 부재별 평균값 뽑기
    SF_ongoing_max_avg = SF_ongoing_max.iloc[:,[0,1,2]]
    SF_ongoing_max_avg = SF_ongoing_max_avg.drop_duplicates()
    SF_ongoing_max_avg.set_index('Element Name', inplace=True)
    
    SF_ongoing_max_avg['V2 max'] = SF_ongoing_max.groupby(['Element Name'])['V2 max'].mean()
    SF_ongoing_max_avg['M3 max'] = SF_ongoing_max.groupby(['Element Name'])['M3 max'].mean()
 
    # 같은 부재(그러나 잘려있는) 경우 최대값 뽑기
    SF_ongoing_max_avg_max = SF_ongoing_max_avg.loc[SF_ongoing_max_avg.groupby(['Property Name'])['V2 max'].idxmax()]
    SF_ongoing_max_avg_max['M3 max'] = SF_ongoing_max_avg.groupby(['Property Name'])['M3 max'].max().tolist()
    
    SF_ongoing_max_avg_max.reset_index(inplace=True, drop=True)

#%% 결과값 정리 후 input sheets에 넣기
    
    SF_ongoing_max_avg_max = pd.merge(transfer_element_info.rename('Property Name'),\
                                      SF_ongoing_max_avg_max, how='left')
        
    SF_ongoing_max_avg_max = SF_ongoing_max_avg_max.dropna()
    SF_ongoing_max_avg_max.reset_index(inplace=True, drop=True)
    
    # SF_ongoing_max_avg 재정렬
    SF_output = SF_ongoing_max_avg_max.iloc[:,[0,2,3]] 
    
    # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
    SF_output = SF_output.replace(np.nan, '', regex=True)
    
    # Using win32com...
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application') # 엑셀 실행
    excel.Visible = False # 엑셀창 안보이게

    wb = excel.Workbooks.Open(input_path + '\\' + beam_xlsx)
    ws = wb.Sheets('Results_T.Beam')
    
    startrow, startcol = 5, 1
    
    # 이름 열 입력
    ws.Range(ws.Cells(startrow, startcol),\
             ws.Cells(startrow + SF_output.shape[0]-1,\
                      startcol)).Value\
    = [[i] for i in SF_output.iloc[:,0]] # series -> list 형식만 입력가능
    
    # V, M열 입력
    ws.Range(ws.Cells(startrow, startcol+12),\
             ws.Cells(startrow + SF_output.shape[0]-1,\
                      startcol + 12 + 2 - 1)).Value\
    = list(SF_output.iloc[:,[1,2]].itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능
    
    wb.Close(SaveChanges=1) # Closing the workbook
    excel.Quit() # Closing the application
    
#%% 부재의 위치별  V, M 값 확인을 위한 도면 작성
    
    # 도면을 그리기 위한 Node List 만들기
    node_map_z = SF_ongoing_max_avg['i-V'].drop_duplicates()
    node_map_z.sort_values(ascending=False, inplace=True)
    node_map_list = node_info_data[node_info_data['V'].isin(node_map_z)]
    
    # 도면을 그리기 위한 Element List 만들기
    element_map_list = pd.merge(SF_ongoing_max_avg, element_info_data.iloc[:,[1,5,6,8,9]]
                                , how='left', left_index=True, right_on='Element Name')
    
    # V, M 크기에 따른 Color 지정
    cmap_V = plt.get_cmap('Reds')
    cmap_M = plt.get_cmap('YlOrBr')
    
    # 층별 Loop
    count = 1
    for i in node_map_z:   
        # 해당 층에 해당하는 Nodes와 Elements만 Extract
        node_map_list_extracted = node_map_list[node_map_list['V'] == i]
        element_map_list_extracted = element_map_list[element_map_list['i-V'] == i]
        element_map_list_extracted.reset_index(inplace=True, drop=True)
        
        # Colorbar, 그래프 Coloring을 위한 설정
        norm_V = plt.Normalize(vmin = element_map_list_extracted['V2 max'].min()\
                             , vmax = element_map_list_extracted['V2 max'].max())
        cmap_V_elem = cmap_V(norm_V(element_map_list_extracted['V2 max']))
        scalar_map_V = mpl.cm.ScalarMappable(norm_V, cmap_V)
        
        norm_M = plt.Normalize(vmin = element_map_list_extracted['M3 max'].min()\
                             , vmax = element_map_list_extracted['M3 max'].max())
        cmap_M_elem = cmap_M(norm_M(element_map_list_extracted['M3 max']))
        scalar_map_M = mpl.cm.ScalarMappable(norm_M, cmap_M)
        
        ## V(전단)     
        # Graph    
        fig1 = plt.figure(count, dpi=150)
        
        plt.scatter(node_map_list_extracted['H1'], node_map_list_extracted['H2'], color='k', s=1)
        
        for idx, row in element_map_list_extracted.iterrows():
            
            element_map_x = [row['i-H1'], row['j-H1']]
            element_map_y = [row['i-H2'], row['j-H2']]
            
            plt.plot(element_map_x, element_map_y, c = cmap_V_elem[idx])
        
        # Colorbar 만들기
        plt.colorbar(scalar_map_V, shrink=0.7, label='V(kN)')
    
        # 기타
        plt.axis('off')
        plt.title(story_info['Story Name'][story_info['Height(mm)'] == i].iloc[0])

        plt.tight_layout()   
        plt.close()
        count += 1
        yield fig1
        
        ## M(모멘트)     
        # Graph    
        fig2 = plt.figure(count, dpi=150)
        
        plt.scatter(node_map_list_extracted['H1'], node_map_list_extracted['H2'], color='k', s=1)
        
        for idx, row in element_map_list_extracted.iterrows():
            
            element_map_x = [row['i-H1'], row['j-H1']]
            element_map_y = [row['i-H2'], row['j-H2']]
            
            plt.plot(element_map_x, element_map_y, c = cmap_M_elem[idx])
        
        # Colorbar 만들기
        plt.colorbar(scalar_map_M, shrink=0.7, label='M(kN-mm)')
    
        # 기타
        plt.axis('off')
        plt.title(story_info['Story Name'][story_info['Height(mm)'] == i].iloc[0])

        plt.tight_layout()   
        plt.close()
        count += 1
        yield fig2
        
#%% Beam Rotation (N) GAGE)

def BR_no_gage(result_path, result_xlsx, input_path, input_xlsx\
               , cri_DE=0.01, cri_MCE=0.025/1.2, yticks=2, xlim=0.04):

#%% Input Sheets 정보 load
    
    story_info = pd.DataFrame()
    deformation_cap = pd.DataFrame()
    
    input_data_raw = pd.ExcelFile(input_path + '\\' + input_xlsx)
    input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', 'Output_C.Beam Properties'], skiprows=3)
    input_data_raw.close()
    
    story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
    deformation_cap = input_data_sheets['Output_C.Beam Properties'].iloc[:,[0,48,49]]
    
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    deformation_cap.columns = ['Name', 'LS', 'CP']
    
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
                                                         'Element Data - Frame Types'], skiprows=2)
        
        beam_rot_data_temp = result_data_sheets['Frame Results - Bending Deform'].iloc[:,[0,2,5,7,10,13,14]]
        beam_rot_data = pd.concat([beam_rot_data, beam_rot_data_temp])
        
    node_data = result_data_sheets['Node Coordinate Data'].iloc[:,[1,4]]
    
    element_data = result_data_sheets['Element Data - Frame Types'].iloc[:,[2,5,7]] # beam의 양 nodes중 한 node에서의 rotation * 2
    
                
    beam_rot_data.columns = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'Distance from I-End', 'H2 Rotation(rad)', 'H3 Rotation(rad)']
    node_data.columns = ['Node ID', 'V(mm)']
    element_data.columns = ['Element Name', 'Property Name', 'I-Node ID']
    
#%% temporary ((L), (R) 등 지우기)
    element_data.loc[:, 'Property Name'] = element_data.loc[:, 'Property Name'].str.split('(').str[0]
    
    #%% 필요없는 부재 빼기, 필요한 부재만 추출
    
    beam_rot_data = beam_rot_data[beam_rot_data['Distance from I-End'] == 0]
    
#%% Analysis Result에 Element, Node 정보 매칭
    
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
    
    DE_load_name_list = [x for x in load_name_list if 'DE' in x]
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]
    
#%% 성능기준(LS, CP) 정리해서 merge
    
    beam_rot_data = pd.merge(beam_rot_data, deformation_cap, how='left', left_on='Property Name', right_on='Name')
    
    beam_rot_data['DE Rotation(rad)'] = beam_rot_data['Major Rotation(rad)']
    beam_rot_data['MCE Rotation(rad)'] = beam_rot_data['Major Rotation(rad)']
    
    beam_rot_data = beam_rot_data[beam_rot_data['Name'].notna()]
    
#%% 조작용 코드
    # 없애고 싶은 부재의 이름 입력(error_beam 확인 후!, DE, MCE에서 다 없어짐)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('B11_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('B15_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB4A_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB4B_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB3D_'))].index)

#%% DE 결과 Plot
    count = 1
    
    if len(DE_load_name_list) != 0:
        
        beam_rot_data_total_DE = pd.DataFrame()    
        
        for load_name in DE_load_name_list:
        
            temp_df_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['DE Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']
                          
            beam_rot_data_total_DE['{}_max'.format(load_name)] = temp_df_max.tolist()
            
            temp_df_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['DE Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']
                          
            beam_rot_data_total_DE['{}_min'.format(load_name)] = temp_df_min.tolist()
            
        beam_rot_data_total_DE['Element Name'] = temp_df_max.index
        
        beam_rot_data_total_DE.reset_index(inplace=True, drop=True)
        
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, element_data, how='left')
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, story_info, how='left', left_on='V(mm)', right_on='Height(mm)')
        beam_rot_data_total_DE.sort_values('Height(mm)', inplace=True)
        # beam_rot_data_total_DE.reset_index(inplace=True, drop=True)
        
    # 평균 열 생성
        
        beam_rot_data_total_DE['DE Max avg'] = beam_rot_data_total_DE.iloc[:,list(range(0,len(DE_load_name_list)*2,2))].mean(axis=1)
        beam_rot_data_total_DE['DE Min avg'] = beam_rot_data_total_DE.iloc[:,list(range(1,len(DE_load_name_list)*2,2))].mean(axis=1)
        
    # 전체 Plot
            
        ### DE 
        fig1 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(-xlim, xlim)
        
        plt.scatter(beam_rot_data_total_DE['DE Max avg'], beam_rot_data_total_DE.loc[:,'V(mm)'], color='k', s=1)
        plt.scatter(beam_rot_data_total_DE['DE Min avg'], beam_rot_data_total_DE.loc[:,'V(mm)'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x= cri_DE, color='r', linestyle='--')
        plt.axvline(x= -cri_DE, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('Rotation(rad)')
        plt.ylabel('Story')
        plt.title('Beam Rotation (DE)')
        
        plt.tight_layout()   
        plt.close()

    # 기준 넘는 점 확인
        error_beam_DE = beam_rot_data_total_DE[['Element Name', 'Property Name', 'Story Name', 'DE Max avg', 'DE Min avg']]\
                      [(beam_rot_data_total_DE['DE Max avg'].abs() >= cri_DE) | (beam_rot_data_total_DE['DE Min avg'].abs() >= cri_DE)]
        
        count += 1
        
        yield fig1
        yield error_beam_DE
        
#%% MCE 결과 Plot
    
    if len(MCE_load_name_list) != 0:
        
        beam_rot_data_total_MCE = pd.DataFrame()    
        
        for load_name in MCE_load_name_list:
        
            temp_df_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['MCE Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']
                          
            beam_rot_data_total_MCE['{}_max'.format(load_name)] = temp_df_max.tolist()
            
            temp_df_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['MCE Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']
                          
            beam_rot_data_total_MCE['{}_min'.format(load_name)] = temp_df_min.tolist()
            
        beam_rot_data_total_MCE['Element Name'] = temp_df_max.index
        
        beam_rot_data_total_MCE.reset_index(inplace=True, drop=True)
        
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, element_data, how='left')
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, story_info, how='left', left_on='V(mm)', right_on='Height(mm)')
        beam_rot_data_total_MCE.sort_values('Height(mm)', inplace=True)
        # beam_rot_data_total_MCE.reset_index(inplace=True, drop=True)
        
    # 평균 열 생성
        
        beam_rot_data_total_MCE['MCE Max avg'] = beam_rot_data_total_MCE.iloc[:,list(range(0,len(MCE_load_name_list)*2,2))].mean(axis=1)
        beam_rot_data_total_MCE['MCE Min avg'] = beam_rot_data_total_MCE.iloc[:,list(range(1,len(MCE_load_name_list)*2,2))].mean(axis=1)
            

        # 전체 Plot 
        ### MCE 
        fig2 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(xlim, -xlim)
        
        # 평균 plot
        plt.scatter(beam_rot_data_total_MCE['MCE Max avg'], beam_rot_data_total_MCE.loc[:,'V(mm)'], color='k', s=1)
        plt.scatter(beam_rot_data_total_MCE['MCE Min avg'], beam_rot_data_total_MCE.loc[:,'V(mm)'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x=cri_MCE, color='r', linestyle='--')
        plt.axvline(x=-cri_MCE, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('Rotation(rad)')
        plt.ylabel('Story')
        plt.title('Beam Rotation (MCE)')
        
        plt.tight_layout()
        plt.close()
    
        # 기준 넘는 점 확인    
        error_beam_MCE = beam_rot_data_total_MCE[['Element Name', 'Property Name', 'Story Name', 'MCE Max avg', 'MCE Min avg']]\
                      [(beam_rot_data_total_MCE['MCE Max avg'].abs() >= cri_MCE) | (beam_rot_data_total_MCE['MCE Min avg'].abs() >= cri_MCE)]
        
        yield fig2
        yield error_beam_MCE