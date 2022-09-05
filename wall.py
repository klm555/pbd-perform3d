import pandas as pd
import numpy as np
import os
from collections import deque  # Double-ended Queue : 자료의 앞, 뒤 양 방향에서 자료를 추가하거나 제거가능
import matplotlib.pyplot as plt
import matplotlib as mpl
import win32com.client

#%% Wall Axial Strain

def AS(input_path, input_xlsx, result_path, result_xlsx='Analysis Result' \
       , max_criteria=0.04, min_criteria=-0.002, yticks=2):
    ''' 

    각각의 벽체의 압축, 인장변형률을 산포도 그래프 형식으로 출력.
    
    Parameters
    ----------
    input_path : str
                 Data Conversion 엑셀 파일의 경로.
                 
    input_xlsx : str
                 Data Conversion 엑셀 파일의 이름. result_xlsx와는 달리 확장자명(.xlsx)까지 기입해줘야한다. 하나의 파일만 불러온다.
    
    result_path : str
                  Perform-3D에서 나온 해석 파일의 경로.
                  
    result_xlsx : str, optional, default='Analysis Result'
                  Perform-3D에서 나온 해석 파일의 이름. 해당 파일 이름이 포함된 파일들을 모두 불러온다.
                               
    max_criteria : float, optional, default=0.04
                   인장변형률 허용기준. (+)부호=압축.
    
    min_criteria : float, optional, default=-0.002
                   압축변형률 허용기준. (-)부호=압축.
    
    yticks : int, optional, default=2
             그래프의 y축 눈금 간격(층간격). 층이 너무 높으면 y축에 너무 많은 층이 표기되기 때문에, 층간격을 조절해서 정돈된 그래프를 표기할 수 있다.
        
    Yields
    -------
    Min, Max값 모두 출력됨. (Min = Red, Max = Black)
    
    fig1 : matplotlib.pyplot.figure or None
           DE(설계지진) 발생 시 벽체 인장, 압축변형률 그래프 (-:math:`\\infty`, 0]
    
    fig2 : matplotlib.pyplot.figure or None
           DE(설계지진) 발생 시 벽체 인장, 압축변형률 그래프 [0, :math:`\\infty`)
    
    fig3 : matplotlib.pyplot.figure or None
           MCE(최대고려지진) 발생 시 벽체 인장, 압축변형률 그래프 (-:math:`\\infty`, 0]
    
    fig4 : matplotlib.pyplot.figure or None
           MCE(최대고려지진) 발생 시 벽체 인장, 압축변형률 그래프 [0, :math:`\\infty`)
                                              
    error_coord_DE : pandas.core.frame.DataFrame or None
                     DE(설계지진) 발생 시 벽체 인장, 압축변형률 NG 부재의 좌표
                     
    error_coord_MCE : pandas.core.frame.DataFrame or None
                     MCE(최대고려지진) 발생 시 벽체 인장, 압축변형률 NG 부재의 좌표
    
    Raises
    -------
    
    References
    -------
    .. [1] "철근콘크리트 건축구조물의 성능기반 내진설계 지침", 대한건축학회, p.44, 2021
    
    '''
#%% Analysis Result 불러오기

    to_load_list = []
    file_names = os.listdir(result_path)
    for file_name in file_names:
        if (result_xlsx in file_name) and ('~$' not in file_name):
            to_load_list.append(file_name)
    
    # Gage data
    AS_gage_data = pd.read_excel(result_path + '\\' + to_load_list[0],
                                   sheet_name='Gage Data - Bar Type', skiprows=[0, 2], header=0, usecols=[0, 2, 7, 9])
    
    # Gage result data
    AS_result_data = pd.DataFrame()
    for i in to_load_list:
        AS_result_data_temp = pd.read_excel(result_path + '\\' + i,
                                   sheet_name='Gage Results - Bar Type', skiprows=[0, 2], header=0, usecols=[0, 2, 5, 7, 8, 9])
        AS_result_data = pd.concat([AS_result_data, AS_result_data_temp])
    
    AS_result_data = AS_result_data.sort_values(by= ['Load Case', 'Element Name', 'Step Type']) # 여러개로 나눠돌릴 경우 순서가 섞여있을 수 있어 DE11~MCE72 순으로 정렬
    
    # Node Coord data
    node_data = pd.read_excel(result_path + '\\' + to_load_list[0],
                                   sheet_name='Node Coordinate Data', skiprows=[0, 2], header=0, usecols=[1, 2, 3, 4])
    
    # Story Info data
    story_info_xlsx_sheet = 'Story Data'
    story_info = pd.read_excel(input_path + '\\' + input_xlsx, sheet_name=story_info_xlsx_sheet, skiprows=3, usecols=[0, 1, 2], keep_default_na=False)
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']

#%% 지진파 이름 list 만들기
    load_name_list = []
    for i in AS_result_data['Load Case'].drop_duplicates():
        new_i = i.split('+')[1]
        new_i = new_i.strip()
        load_name_list.append(new_i)
    
    gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
    seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]
    
    seismic_load_name_list.sort()
    
    DE_load_name_list = [x for x in load_name_list if 'DE' in x] # base shear로 사용할 지진파 개수 산정을 위함
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]
    
#%% 데이터 매칭 후 결과뽑기

    AS_result_data = AS_result_data[AS_result_data['Load Case']\
                                    .str.contains('|'.join(seismic_load_name_list))]

    # 층분할된 곳의 Axial strain gage는 max(abs(분할된 두 값))로 assign하기
    
        
    ### Gage data에서 Element Name, I-Node ID 불러와서 v좌표 match하기
    AS_gage_data = AS_gage_data[['Element Name', 'I-Node ID']]; 
    
    gage_num = len(AS_gage_data) # gage 개수 얻기
    
    # I-Node의 v좌표 match해서 추가
    AS_gage_data = AS_gage_data.join(node_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')
    
    ### AS_total data 만들기
    AS_max = AS_result_data[(AS_result_data['Step Type'] == 'Max') & (AS_result_data['Performance Level'] == 1)][['Axial Strain']].values # dataframe을 array로
    AS_max = AS_max.reshape(gage_num, len(seismic_load_name_list), order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
    AS_max = pd.DataFrame(AS_max) # array를 다시 dataframe으로
    AS_min = AS_result_data[(AS_result_data['Step Type'] == 'Min') & (AS_result_data['Performance Level'] == 1)][['Axial Strain']].values
    AS_min = AS_min.reshape(gage_num, len(seismic_load_name_list), order='F')
    AS_min = pd.DataFrame(AS_min)
    AS_total = pd.concat([AS_max, AS_min], axis=1)
    
    ### AS_avg_data 만들기
    DE_max_avg = AS_total.iloc[:, 0:len(DE_load_name_list)].mean(axis=1)
    MCE_max_avg = AS_total.iloc[:, len(DE_load_name_list) : len(DE_load_name_list)+len(MCE_load_name_list)].mean(axis=1)
    DE_min_avg = AS_total.iloc[:, len(DE_load_name_list)+len(MCE_load_name_list) : 2*len(DE_load_name_list)+len(MCE_load_name_list)].mean(axis=1)
    MCE_min_avg = AS_total.iloc[:, 2*len(DE_load_name_list)+len(MCE_load_name_list) : 2*len(DE_load_name_list) + 2*len(MCE_load_name_list)].mean(axis=1)
    AS_avg_total = pd.concat([AS_gage_data.loc[:, ['H1', 'H2', 'V']], DE_max_avg, DE_min_avg, MCE_max_avg, MCE_min_avg], axis=1)
    AS_avg_total.columns = ['X(mm)', 'Y(mm)', 'Z(mm)', 'DE_max_avg', 'DE_min_avg', 'MCE_max_avg', 'MCE_min_avg']

#%% ***조작용 코드
    # 데이터 없애기 위한 기준값 입력
    # AS_avg_total = AS_avg_total.drop(AS_avg_total[(AS_avg_total.loc[:,'DE_min_avg'] < -0.002)].index)
    # AS_avg_total = AS_avg_total.drop(AS_avg_total[(AS_avg_total.loc[:,'MCE_min_avg'] < -0.002)].index)
    # .....위와 같은 포맷으로 계속

#%% 그래프
    count = 1    

    # DE 그래프
    if len(DE_load_name_list) != 0:
            
        # AS_DE_1
        fig1 = plt.figure(count, dpi=150, figsize=(5,4))  # 그래프 사이즈
        plt.xlim(-0.003, 0)
        
        plt.scatter(AS_avg_total['DE_min_avg'], AS_avg_total['Z(mm)'], color = 'r', s=5) # s=1 : point size
        plt.scatter(AS_avg_total['DE_max_avg'], AS_avg_total['Z(mm)'], color = 'k', s=5)
        
        # height값에 대응되는 층 이름으로 y축 눈금 작성
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        
        # reference line 그려서 허용치 나타내기
        plt.axvline(x= min_criteria, color='r', linestyle='--')
        plt.axvline(x= max_criteria, color='r', linestyle='--')
        
        plt.grid(linestyle='-.')
        plt.xlabel('Axial Strain(m/m)')
        plt.ylabel('Story')
        plt.title('DE (Compressive)')
        
        plt.tight_layout()
        plt.style.use('fast')
        plt.close()
        count += 1
        
        # AS_DE_2
        fig2 = plt.figure(count, dpi=150, figsize=(5,4))  # 그래프 사이즈
        plt.xlim(0, 0.013)
        plt.scatter(AS_avg_total['DE_min_avg'], AS_avg_total['Z(mm)'], color = 'r', s=5) # s=1 : point size
        plt.scatter(AS_avg_total['DE_max_avg'], AS_avg_total['Z(mm)'], color = 'k', s=5)
        
        # height값에 대응되는 층 이름으로 y축 눈금 작성
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        
        plt.axvline(x= min_criteria, color='r', linestyle='--')
        plt.axvline(x= max_criteria, color='r', linestyle='--')
        
        plt.grid(linestyle='-.')
        plt.xlabel('Axial Strain(m/m)')
        plt.ylabel('Story')
        plt.title('DE (Tensile)')
        
        plt.tight_layout()
        plt.style.use('fast')
        plt.close()
        count += 1
        
        error_coord_DE = AS_avg_total[(AS_avg_total['DE_max_avg'] >= max_criteria)\
                                      | (AS_avg_total['DE_min_avg'] <= min_criteria)]
        
        yield fig1
        yield fig2
        yield error_coord_DE
    
    # MCE 그래프
    if len(MCE_load_name_list) != 0:
            
        # AS_MCE_1
        fig3 = plt.figure(count, dpi=150, figsize=(5,4))
        plt.xlim(-0.003, 0)
        plt.scatter(AS_avg_total['MCE_min_avg'], AS_avg_total['Z(mm)'], color = 'r', s=5)
        plt.scatter(AS_avg_total['MCE_max_avg'], AS_avg_total['Z(mm)'], color = 'k', s=5)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        
        plt.axvline(x= min_criteria, color='r', linestyle='--')
        plt.axvline(x= max_criteria, color='r', linestyle='--')
        
        plt.grid(linestyle='-.')
        plt.xlabel('Axial Strain(m/m)')
        plt.ylabel('Story')
        plt.title('MCE (Compressive)')
        
        plt.tight_layout()
        plt.style.use('fast')
        plt.close()
        count += 1
        
        # AS_MCE_2
        fig4 = plt.figure(count, dpi=150, figsize=(5,4))
        plt.xlim(0, 0.013)
        plt.scatter(AS_avg_total['MCE_min_avg'], AS_avg_total['Z(mm)'], color = 'r', s=5)
        plt.scatter(AS_avg_total['MCE_max_avg'], AS_avg_total['Z(mm)'], color = 'k', s=5)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        
        plt.axvline(x= min_criteria, color='r', linestyle='--')
        plt.axvline(x= max_criteria, color='r', linestyle='--')
        
        plt.grid(linestyle='-.')
        plt.xlabel('Axial Strain(m/m)')
        plt.ylabel('Story')
        plt.title('MCE (Tensile)')
        
        plt.tight_layout()
        plt.style.use('fast')
        plt.close()
        count += 1
        
        error_coord_MCE = AS_avg_total[(AS_avg_total['MCE_max_avg'] >= max_criteria)\
                                       | (AS_avg_total['MCE_min_avg'] <= min_criteria)]      
        
        yield fig3
        yield fig4
        yield error_coord_MCE


#%% Shear Wall Rotation

def SWR(input_path, input_xlsx, result_path, result_xlsx='Analysis Result'\
        , DE_criteria=0.002, MCE_criteria=0.004/1.2, yticks=2, xlim=0.005):
    ''' 

    각각의 벽체의 회전각을 산포도 그래프 형식으로 출력.
    
    Parameters
    ----------
    input_path : str
                 Data Conversion 엑셀 파일의 경로.
                 
    input_xlsx : str
                 Data Conversion 엑셀 파일의 이름. result_xlsx와는 달리 확장자명(.xlsx)까지 기입해줘야한다. 하나의 파일만 불러온다.
                 
    result_path : str
                  Perform-3D에서 나온 해석 파일의 경로.
                  
    result_xlsx : str, optional, default='Analysis Result'
                  Perform-3D에서 나온 해석 파일의 이름. 해당 파일 이름이 포함된 파일들을 모두 불러온다.
                 
    DE_criteria : float, optional, default=0.002/1.2
                  LS(인명안전)에 대한 벽체 회전각 허용기준. default=가장 보수적인 값
    
    MCE_criteria : float, optional, default=-0.004/1.2
                   CP(붕괴방지)에 대한 벽체 회전각 허용기준. default=가장 보수적인 값
                   
    yticks : int, optional, default=2
             그래프의 y축 눈금 간격(층간격). 층이 너무 높으면 y축에 너무 많은 층이 표기되기 때문에, 층간격을 조절해서 정돈된 그래프를 표기할 수 있다.

    xlim : int, optional, default=0.005
           그래프의 x축 limit 값. x축 limit 안의 값만 표기되므로, limit를 넘어가는 값을 확인하고 싶을 시에는 더 큰 xlim 값을 사용하면 된다.

    Yields
    -------
    Min, Max값 모두 출력됨. 
    
    fig1 : matplotlib.pyplot.figure or None
           DE(설계지진) 발생 시 벽체 회전각 그래프
    
    fig2 : matplotlib.pyplot.figure or None
           MCE(최대고려지진) 발생 시 벽체 회전각 그래프
    
    error_coord_DE : pandas.core.frame.DataFrame or None
                     DE(설계지진) 발생 시 기준값을 초과하는 벽체의 좌표
                     
    error_coord_MCE : pandas.core.frame.DataFrame or None
                     MCE(최대고려지진) 발생 시 기준값을 초과하는 벽체의 좌표                                          
    
    Raises
    -------
    
    References
    -------
    .. [1] "철근콘크리트 건축구조물의 성능기반 내진설계 지침", 대한건축학회, p.79, 2021    
    
    '''
    #%% Analysis Result 불러오기
    to_load_list = []
    file_names = os.listdir(result_path)
    for file_name in file_names:
        if (result_xlsx in file_name) and ('~$' not in file_name):
            to_load_list.append(file_name)

    # Gage data
    gage_data = pd.read_excel(result_path + '\\' + to_load_list[0],
                                   sheet_name='Gage Data - Wall Type', skiprows=[0, 2], header=0, usecols=[0, 2, 7, 9, 11, 13]) # usecols로 원하는 열만 불러오기

    # Gage result data
    wall_rot_data = pd.DataFrame()
    for i in to_load_list:
        wall_rot_data_temp = pd.read_excel(result_path + '\\' + i,
                                   sheet_name='Gage Results - Wall Type', skiprows=[0, 2], header=0, usecols=[0, 2, 5, 7, 9, 11])
        wall_rot_data = pd.concat([wall_rot_data, wall_rot_data_temp])

    wall_rot_data.sort_values(['Load Case', 'Element Name'] , inplace=True)

    # Node Coord data
    node_data = pd.read_excel(result_path + '\\' + to_load_list[0],
                                   sheet_name='Node Coordinate Data', skiprows=[0, 2], header=0, usecols=[1, 2, 3, 4])

    # Story Info data
    story_info_xlsx_sheet = 'Story Data'
    story_info = pd.read_excel(input_path + '\\' + input_xlsx, sheet_name=story_info_xlsx_sheet, skiprows=3, usecols=[0, 1, 2], keep_default_na=False)
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    story_name = story_info.loc[:, 'Story Name']

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
    
    wall_rot_data = wall_rot_data[wall_rot_data['Load Case']\
                                      .str.contains('|'.join(seismic_load_name_list))]
    
    ### Gage data에서 Element Name, I-Node ID 불러와서 v좌표 match하기
    gage_data = gage_data[['Element Name', 'I-Node ID']]; gage_num = len(gage_data) # gage 개수 얻기
    node_data_V = node_data[['Node ID', 'V']]

    # I-Node의 v좌표 match해서 추가
    gage_data = gage_data.join(node_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')

    ### SWR_total data 만들기

    SWR_max = wall_rot_data[(wall_rot_data['Step Type'] == 'Max') & (wall_rot_data['Performance Level'] == 1)][['Rotation']].values # dataframe을 array로
    SWR_max = SWR_max.reshape(gage_num, len(seismic_load_name_list), order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
    SWR_max = pd.DataFrame(SWR_max) # array를 다시 dataframe으로
    SWR_min = wall_rot_data[(wall_rot_data['Step Type'] == 'Min') & (wall_rot_data['Performance Level'] == 1)][['Rotation']].values
    SWR_min = SWR_min.reshape(gage_num, len(seismic_load_name_list), order='F')
    SWR_min = pd.DataFrame(SWR_min)
    SWR_total = pd.concat([SWR_max, SWR_min], axis=1) # DE11_max~MCE72_max, DE11_min~MCE72_min 각각 28개씩

    ### SWR_avg_data 만들기
    DE_max_avg = SWR_total.iloc[:, 0:len(DE_load_name_list)].mean(axis=1)
    MCE_max_avg = SWR_total.iloc[:, len(DE_load_name_list) : len(DE_load_name_list)+ len(MCE_load_name_list)].mean(axis=1)
    DE_min_avg = SWR_total.iloc[:, len(DE_load_name_list)+len(MCE_load_name_list) : 2*len(DE_load_name_list)+len(MCE_load_name_list)].mean(axis=1)
    MCE_min_avg = SWR_total.iloc[:, 2*len(DE_load_name_list)+len(MCE_load_name_list) : 2*len(DE_load_name_list) + 2*len(MCE_load_name_list)].mean(axis=1)
    SWR_avg_total = pd.concat([gage_data.loc[:,['H1', 'H2', 'V']], DE_max_avg, DE_min_avg, MCE_max_avg, MCE_min_avg], axis=1)
    SWR_avg_total.columns = ['X(mm)', 'Y(mm)', 'Height(mm)', 'DE_max_avg', 'DE_min_avg', 'MCE_max_avg', 'MCE_min_avg']

    #%% ***조작용 코드
    # SWR_avg_total = SWR_avg_total.drop(SWR_avg_total[(SWR_avg_total.iloc[:,2] < -0.0038) | (SWR_avg_total.iloc[:,1] > 0.0038)].index) # DE
    # SWR_avg_total = SWR_avg_total.drop(SWR_avg_total[(SWR_avg_total.iloc[:,4] < -0.0035) | (SWR_avg_total.iloc[:,3] > 0.0035)].index) # MCE

    #%% 그래프
    count = 1
    
    # DE 그래프
    if len(DE_load_name_list) != 0:
    
        fig1 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(-xlim, xlim)
        plt.scatter(SWR_avg_total['DE_min_avg'], SWR_avg_total['Height(mm)'], color = 'k', s=1) # s=1 : point size
        plt.scatter(SWR_avg_total['DE_max_avg'], SWR_avg_total['Height(mm)'], color = 'k', s=1)
    
        # height값에 대응되는 층 이름으로 y축 눈금 작성
        plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
    
        # reference line 그려서 허용치 나타내기
        plt.axvline(x= -DE_criteria, color='r', linestyle='--')
        plt.axvline(x= DE_criteria, color='r', linestyle='--')
    
        plt.grid(linestyle='-.')
        plt.xlabel('Rotation(rad)')
        plt.ylabel('Story')
        plt.title('Wall Rotation (DE)')
    
        plt.tight_layout()
        plt.close()
        count += 1
        
        error_coord_DE = SWR_avg_total[(SWR_avg_total['DE_max_avg'] >= DE_criteria)\
                                       | (SWR_avg_total['DE_min_avg'] <= -DE_criteria)]          
        
        yield fig1
        yield error_coord_DE

    # MCE 그래프
    if len(MCE_load_name_list) != 0:
        
        fig2 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(-xlim, xlim)
        plt.scatter(SWR_avg_total['MCE_min_avg'], SWR_avg_total['Height(mm)'], color = 'k', s=1)
        plt.scatter(SWR_avg_total['MCE_max_avg'], SWR_avg_total['Height(mm)'], color = 'k', s=1)
    
        plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
    
        plt.axvline(x= -MCE_criteria, color='r', linestyle='--')
        plt.axvline(x= MCE_criteria, color='r', linestyle='--')
    
        plt.grid(linestyle='-.')
        plt.xlabel('Rotation(rad)')
        plt.ylabel('Story')
        plt.title('Wall Rotation (MCE)')
    
        plt.tight_layout()
        plt.close()
        count += 1
        
        error_coord_MCE = SWR_avg_total[(SWR_avg_total['MCE_max_avg'] >= MCE_criteria)\
                                        | (SWR_avg_total['MCE_min_avg'] <= -MCE_criteria)]  
        
        yield fig2
        yield error_coord_MCE

#%% Shear Wall Rotation (DCR)

def SWR_DCR(input_path, input_xlsx, result_path, result_xlsx='Analysis Result', DCR_criteria=1, yticks=2, xlim=3):
    '''
    벽체 회전각과 기준에서 계산한 허용기준을 각각의 벽체에 대해 비교하여 DCR 방식으로 산포도 그래프를 출력.
    
    Parameters
    ----------
    input_path : str
                 Data Conversion 엑셀 파일의 경로.
                 
    input_xlsx : str
                 Data Conversion 엑셀 파일의 이름. result_xlsx와는 달리 확장자명(.xlsx)까지 기입해줘야한다. 하나의 파일만 불러온다.
                 
    result_path : str
                  Perform-3D에서 나온 해석 파일의 경로.
                  
    result_xlsx : str, optional, default='Analysis Result'
                  Perform-3D에서 나온 해석 파일의 이름. 해당 파일 이름이 포함된 파일들을 모두 불러온다.                  

    DCR_criteria : float, optional, default=1
                   DCR 기준값.
                   
    yticks : int, optional, default=2
             그래프의 y축 눈금 간격(층간격). 층이 너무 높으면 y축에 너무 많은 층이 표기되기 때문에, 층간격을 조절해서 정돈된 그래프를 표기할 수 있다.

    xlim : int, optional, default=3
           그래프의 x축 limit 값. x축 limit 안의 값만 표기되므로, limit를 넘어가는 값을 확인하고 싶을 시에는 더 큰 xlim 값을 사용하면 된다.

    Yields
    -------
    Min, Max값 모두 출력됨. 
    
    fig1 : matplotlib.pyplot.figure or None
           DE(설계지진) 발생 시 벽체 회전각 DCR 그래프
    
    fig2 : matplotlib.pyplot.figure or None
           MCE(최대고려지진) 발생 시 벽체 회전각 DCR 그래프
    
    error_wall_DE : pandas.core.frame.DataFrame or None
                    DE(설계지진) 발생 시 DCR 기준값을 초과하는 벽체의 정보
                     
    error_wall_MCE : pandas.core.frame.DataFrame or None
                     MCE(최대고려지진) 발생 시 DCR 기준값을 초과하는 벽체의 정보                                          
    
    Raises
    -------
    
    References
    -------
    .. [1] "철근콘크리트 건축구조물의 성능기반 내진설계 지침", 대한건축학회, p.79, 2021
    
    '''    
#%% Input Sheets 정보 load

    story_info = pd.DataFrame()
    deformation_cap = pd.DataFrame()
    
    input_data_raw = pd.ExcelFile(input_path + '\\' + input_xlsx)
    input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', 'Results_Wall'], skiprows=3)
    input_data_raw.close()
    
    story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
    deformation_cap = input_data_sheets['Results_Wall'].iloc[:,[0,11,12,13,14,44,45,50,51]]
    
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    deformation_cap.columns = ['Name', 'Vu_DE_H1', 'Vu_DE_H2', 'Vu_MCE_H1', 'Vu_MCE_H2'\
                               , 'LS(H1)', 'LS(H2)', 'CP(H1)', 'CP(H2)']
    
    story_name = story_info.loc[:, 'Story Name']
    
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
        
        wall_rot_data_temp = result_data_sheets['Gage Results - Wall Type'].iloc[:,[0,2,5,7,9,11]]
        wall_rot_data = pd.concat([wall_rot_data, wall_rot_data_temp])
        
    node_data = result_data_sheets['Node Coordinate Data'].iloc[:,[1,2,3,4]]
    gage_data = result_data_sheets['Gage Data - Wall Type'].iloc[:,[2,7,9,11,13]] # beam의 양 nodes중 한 node에서의 rotation * 2
    element_data = result_data_sheets['Element Data - Shear Wall'].iloc[:,[2,5,7,9,11,13]] # beam의 양 nodes중 한 node에서의 rotation * 2
    
#%% Gage Data & Result에 Node 정보 매칭
    
    gage_data = gage_data.drop_duplicates()
    node_data = node_data.drop_duplicates()
    
#     wall_rot_data = pd.merge(wall_rot_data, gage_data, how='left')
#     wall_rot_data = pd.merge(wall_rot_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
#     wall_rot_data = pd.merge(wall_rot_data, node_data, how='left', left_on='J-Node ID', right_on='Node ID')
#     wall_rot_data = pd.merge(wall_rot_data, node_data, how='left', left_on='K-Node ID', right_on='Node ID')
#     wall_rot_data = pd.merge(wall_rot_data, node_data, how='left', left_on='L-Node ID', right_on='Node ID')

#     wall_rot_data = wall_rot_data.iloc[:, np.r_[0:9, 10:13, 14:17, 18:21, 22:25]]
        
#     wall_rot_data.columns.values[9] = 'X(I-node)'
#     wall_rot_data.columns.values[10] = 'Y(I-node)'
#     wall_rot_data.columns.values[11] = 'Z(I-node)'
#     wall_rot_data.columns.values[12] = 'X(J-node)'
#     wall_rot_data.columns.values[13] = 'Y(J-node)'
#     wall_rot_data.columns.values[14] = 'Z(J-node)'
#     wall_rot_data.columns.values[15] = 'X(K-node)'
#     wall_rot_data.columns.values[16] = 'Y(K-node)'
#     wall_rot_data.columns.values[17] = 'Z(K-node)'
#     wall_rot_data.columns.values[18] = 'X(L-node)'
#     wall_rot_data.columns.values[19] = 'Y(L-node)'
#     wall_rot_data.columns.values[20] = 'Z(L-node)'

#     wall_rot_data.reset_index(inplace=True, drop=True)
    
# #%% Element Data에 Node정보 매칭

#     element_data = element_data.drop_duplicates()
    
#     element_data = pd.merge(element_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
#     element_data = pd.merge(element_data, node_data, how='left', left_on='J-Node ID', right_on='Node ID')
#     element_data = pd.merge(element_data, node_data, how='left', left_on='K-Node ID', right_on='Node ID')
#     element_data = pd.merge(element_data, node_data, how='left', left_on='L-Node ID', right_on='Node ID')

#     element_data = element_data.iloc[:, np.r_[0:6, 7:10, 11:14, 15:18, 19:22]]
    
#     element_data.columns = ['Element Name', 'Property Name', 'I-Node ID', 'J-Node ID'\
#                             , 'K-Node ID', 'L-Node ID', 'X(I-node)', 'Y(I-node)', 'Z(I-node)'\
#                             , 'X(J-node)', 'Y(J-node)', 'Z(J-node)', 'X(K-node)', 'Y(K-node)'\
#                             , 'Z(K-node)', 'X(L-node)', 'Y(L-node)', 'Z(L-node)']

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
            deformation_cap_DE = pd.concat([deformation_cap_DE, pd.Series(deformation_cap.iloc[i, 5])], ignore_index=True)
        else:
            deformation_cap_DE = pd.concat([deformation_cap_DE, pd.Series(deformation_cap.iloc[i, 6])], ignore_index=True)
    
    # CP 기준
    deformation_cap_MCE = pd.DataFrame()
    for i in range(len(deformation_cap)):
        if deformation_cap.iloc[i, 3] > deformation_cap.iloc[i, 4]:
            deformation_cap_MCE = pd.concat([deformation_cap_MCE, pd.Series(deformation_cap.iloc[i, 7])], ignore_index=True)
        else:
            deformation_cap_MCE = pd.concat([deformation_cap_MCE, pd.Series(deformation_cap.iloc[i, 8])], ignore_index=True)
    
    SWR_criteria = pd.concat([deformation_cap['Name'], deformation_cap_DE, deformation_cap_MCE], axis = 1, ignore_index=True)
    SWR_criteria.columns = ['Name', 'DE criteria', 'MCE criteria']
        
    #### OLD VERSION ####    
    # 이전 버전의 네이밍에 맞게 merge하는 방법

    new_name = []
    for i in SWR_criteria['Name']:
        if i.count('_') == 2:
            new_name.append(i.split('_')[0] + '_' + i.split('_')[2])
    
    SWR_criteria['Name'] = new_name    
    #####################
    
    ### SWR avg total에 SWR criteria join(wall name 기준)
    SWR_avg_total = pd.merge(SWR_avg_total, SWR_criteria, how='left'\
                             , left_on='gage_name', right_on='Name')
    
    #SWR_avg_total.dropna(inplace=True)
    SWR_avg_total['DCR_DE_min'] = SWR_avg_total['DE_min_avg'].abs()/SWR_avg_total['DE criteria']
    SWR_avg_total['DCR_DE_max'] = SWR_avg_total['DE_max_avg']/SWR_avg_total['DE criteria']
    SWR_avg_total['DCR_MCE_min'] = SWR_avg_total['MCE_min_avg'].abs()/SWR_avg_total['MCE criteria']
    SWR_avg_total['DCR_MCE_max'] = SWR_avg_total['MCE_max_avg']/SWR_avg_total['MCE criteria']
    
    #%% ***조작용 코드
    # SWR_avg_total = SWR_avg_total.drop(SWR_avg_total[(SWR_avg_total.iloc[:,2] < -0.0038) | (SWR_avg_total.iloc[:,1] > 0.0038)].index) # DE
    # SWR_avg_total = SWR_avg_total.drop(SWR_avg_total[(SWR_avg_total.iloc[:,4] < -0.0035) | (SWR_avg_total.iloc[:,3] > 0.0035)].index) # MCE
    
    #%% 그래프
    count = 1

    ### DE 그래프
    if len(DE_load_name_list) != 0:
        
        fig1 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        plt.scatter(SWR_avg_total['DCR_DE_min'], SWR_avg_total['Height'], color='k', s=1)
        plt.scatter(SWR_avg_total['DCR_DE_max'], SWR_avg_total['Height'], color='k', s=1)
        plt.yticks(story_info['Height(mm)'][::-3], story_info['Story Name'][::-3])
        plt.axvline(x = DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Wall Rotation (DE)')
        
        plt.close()
        count += 1
    
        # 기준 넘는 벽체 확인
        error_wall_DE = SWR_avg_total[['gage_name', 'DCR_DE_min', 'DCR_DE_max']]\
                        [(SWR_avg_total['DCR_DE_min']>= DCR_criteria) | \
                         (SWR_avg_total['DCR_DE_max']>= DCR_criteria)]
                            
        yield fig1
        yield error_wall_DE    
        
    ### MCE 그래프
    if len(MCE_load_name_list) != 0:
        
        fig2 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        plt.scatter(SWR_avg_total['DCR_MCE_min'], SWR_avg_total['Height'], color='k', s=1)
        plt.scatter(SWR_avg_total['DCR_MCE_max'], SWR_avg_total['Height'], color='k', s=1)
        plt.yticks(story_info['Height(mm)'][::-3], story_info['Story Name'][::-3])
        plt.axvline(x = DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Wall Rotation (MCE)')
        
        plt.close()
        count += 1
        
        # 기준 넘는 벽체 확인
        error_wall_MCE = SWR_avg_total[['gage_name', 'DCR_MCE_min', 'DCR_MCE_max']]\
                        [(SWR_avg_total['DCR_MCE_min']>= DCR_criteria) | \
                         (SWR_avg_total['DCR_MCE_max']>= DCR_criteria)]
        
        
        yield fig2
        yield error_wall_MCE
        
#%% Wall_SF
# 오류없는 또는 정확한 결과를 위해서는 MCE11, MCE12와 같이 짝이되는 지진파가 함께 있어야 함.

def wall_SF(input_path, input_xlsx, result_path, result_xlsx='Analysis Result', graph=True,  DCR_criteria=1, yticks=2, xlim=3): 
    ''' 

    Perform-3D 해석 결과에서 벽체의 축력, 전단력(DE, MCE)을 불러와 Data Conversion 엑셀파일의 Results_Wall 시트를 작성하고, 벽체 전단력 DCR 그래프를 출력(optional). \n
    
       
    벽체 회전각과 기준에서 계산한 허용기준을 각각의 벽체에 대해 비교하여 DCR 방식으로 산포도 그래프를 출력.
    
    Parameters
    ----------
    input_path : str
                 Data Conversion 엑셀 파일의 경로.
                 
    input_xlsx : str
                 Data Conversion 엑셀 파일의 이름. result_xlsx와는 달리 확장자명(.xlsx)까지 기입해줘야한다. 하나의 파일만 불러온다.
                 
    result_path : str
                  Perform-3D에서 나온 해석 파일의 경로.
                  
    result_xlsx : str, optional, default='Analysis Result'
                  Perform-3D에서 나온 해석 파일의 이름. 해당 파일 이름이 포함된 파일들을 모두 불러온다.                  

    graph : bool, optional, default=True
            True = Data Conversion 엑셀 파일에 입력된 값으로 DCR 그래프를 그릴지 설정.
            False = Data Conversion 엑셀파일만 작성. (그래프 X)    
    
    DCR_criteria : float, optional, default=1
                   DCR 기준값.
                   
    yticks : int, optional, default=2
             그래프의 y축 눈금 간격(층간격). 층이 너무 높으면 y축에 너무 많은 층이 표기되기 때문에, 층간격을 조절해서 정돈된 그래프를 표기할 수 있다.

    xlim : int, optional, default=3
           그래프의 x축 limit 값. x축 limit 안의 값만 표기되므로, limit를 넘어가는 값을 확인하고 싶을 시에는 더 큰 xlim 값을 사용하면 된다.

    Yields
    -------
    Min, Max값 모두 출력됨. 
    
    fig1 : matplotlib.pyplot.figure or None
           DE(설계지진) 발생 시 벽체 회전각 DCR 그래프
    
    fig2 : matplotlib.pyplot.figure or None
           MCE(최대고려지진) 발생 시 벽체 회전각 DCR 그래프
    
    error_wall_DE : pandas.core.frame.DataFrame or None
                    DE(설계지진) 발생 시 DCR 기준값을 초과하는 벽체의 정보
                     
    error_wall_MCE : pandas.core.frame.DataFrame or None
                     MCE(최대고려지진) 발생 시 DCR 기준값을 초과하는 벽체의 정보                                          
    
    Raises
    -------
    
    References
    -------
    .. [1] "철근콘크리트 건축구조물의 성능기반 내진설계 지침", 대한건축학회, p.79, 2021
    
    '''
#%% Input Sheet 정보 load
        
    story_info = pd.DataFrame()
    transfer_element_info = pd.DataFrame()

    input_xlsx_sheet = 'Output_Wall Properties'
    input_data_raw = pd.ExcelFile(input_path + '\\' + input_xlsx)
    input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', input_xlsx_sheet], skiprows=3)
    input_data_raw.close()

    story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
    transfer_element_info = input_data_sheets[input_xlsx_sheet].iloc[:,0:10]
    story_info = story_info[::-1]
    story_info.reset_index(inplace=True, drop=True)

    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    transfer_element_info.columns = ['Name', 'Length(mm)', 'Thickness(mm)', 'Concrete Grade', 'Rebar Type', 'V.Rebar Type',\
                                     'V.Rebar Spacing(mm)', 'V.Rebar EA', 'H.Rebar Type', 'H.Rebar Spacing(mm)']

    transfer_element_info.reset_index(inplace=True, drop=True)

#%% Analysis Result 불러오기

    to_load_list = []
    file_names = os.listdir(result_path)
    for file_name in file_names:
        if (result_xlsx in file_name) and ('~$' not in file_name):
            to_load_list.append(file_name)

    # 전단력 불러오기
    wall_SF_data = pd.DataFrame()

    for i in to_load_list:
        result_data_raw = pd.ExcelFile(result_path + '\\' + i)
        result_data_sheets = pd.read_excel(result_data_raw, ['Structure Section Forces', 'Frame Results - End Forces'], skiprows=2)
        
        wall_SF_data_temp = result_data_sheets['Structure Section Forces'].iloc[:,[0,3,5,6,7,8]]
        wall_SF_data = pd.concat([wall_SF_data, wall_SF_data_temp])

    wall_SF_data.columns = ['Name', 'Load Case', 'Step Type', 'H1(kN)', 'H2(kN)', 'V(kN)']

    # 필요없는 전단력 제거(층전단력)
    wall_SF_data = wall_SF_data[wall_SF_data['Name'].str.count('_') == 2] # underbar가 두개 들어간 행만 선택
        
    wall_SF_data.reset_index(inplace=True, drop=True)

#%% 부재명, H1, H2 값 뽑기

    # 지진파 이름 list 만들기
    # gravity_load_name = [x.split('+',1)[1].strip() \
    #                      for x in wall_SF_data['Load Case'].drop_duplicates() if '[0]' in x]
    
    load_name_list = []
    for i in wall_SF_data['Load Case'].drop_duplicates():
        new_i = i.split('+')[1]
        new_i = new_i.strip()
        load_name_list.append(new_i)

    gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
    seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]

    seismic_load_name_list.sort()
    
    DE_load_name_list = [x for x in load_name_list if 'DE' in x]
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]


#%% 데이터 Grouping

    shear_force_H1_DE_data_grouped = pd.DataFrame()
    shear_force_H2_DE_data_grouped = pd.DataFrame()
    shear_force_H1_MCE_data_grouped = pd.DataFrame()
    shear_force_H2_MCE_data_grouped = pd.DataFrame()

    # DE를 max, min으로 grouping
    for load_name in DE_load_name_list:
        shear_force_H1_DE_data_grouped['{}_H1_max'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (wall_SF_data['Step Type'] == 'Max')]['H1(kN)'].values
            
        shear_force_H1_DE_data_grouped['{}_H1_min'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (wall_SF_data['Step Type'] == 'Min')]['H1(kN)'].values

        shear_force_H2_DE_data_grouped['{}_H2_max'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (wall_SF_data['Step Type'] == 'Max')]['H2(kN)'].values
            
        shear_force_H2_DE_data_grouped['{}_H2_min'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (wall_SF_data['Step Type'] == 'Min')]['H2(kN)'].values   

    # MCE를 max, min으로 grouping
    for load_name in MCE_load_name_list:
        shear_force_H1_MCE_data_grouped['{}_H1_max'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (wall_SF_data['Step Type'] == 'Max')]['H1(kN)'].values
            
        shear_force_H1_MCE_data_grouped['{}_H1_min'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (wall_SF_data['Step Type'] == 'Min')]['H1(kN)'].values

        shear_force_H2_MCE_data_grouped['{}_H2_max'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (wall_SF_data['Step Type'] == 'Max')]['H2(kN)'].values
            
        shear_force_H2_MCE_data_grouped['{}_H2_min'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (wall_SF_data['Step Type'] == 'Min')]['H2(kN)'].values   

    if len(DE_load_name_list) != 0:

        # all 절대값
        shear_force_H1_DE_abs = shear_force_H1_DE_data_grouped.abs()
        shear_force_H2_DE_abs = shear_force_H2_DE_data_grouped.abs()
        
        # 최대값 every 4 columns
        shear_force_H1_DE_max = shear_force_H1_DE_abs.groupby([[i//4 for i in range(0,2*len(DE_load_name_list))]], axis=1).max()
        shear_force_H2_DE_max = shear_force_H2_DE_abs.groupby([[i//4 for i in range(0,2*len(DE_load_name_list))]], axis=1).max()

        # 1.2 * 평균값
        shear_force_H1_DE_avg = 1.2 * shear_force_H1_DE_max.mean(axis=1)
        shear_force_H2_DE_avg = 1.2 * shear_force_H2_DE_max.mean(axis=1)
        
    else : 
        shear_force_H1_DE_avg = ''
        shear_force_H2_DE_avg = ''

    if len(MCE_load_name_list) != 0:

        # all 절대값
        shear_force_H1_MCE_abs = shear_force_H1_MCE_data_grouped.abs()
        shear_force_H2_MCE_abs = shear_force_H2_MCE_data_grouped.abs()
        
        # 최대값 every 4 columns
        shear_force_H1_MCE_max = shear_force_H1_MCE_abs.groupby([[i//4 for i in range(0,2*len(MCE_load_name_list))]], axis=1).max()
        shear_force_H2_MCE_max = shear_force_H2_MCE_abs.groupby([[i//4 for i in range(0,2*len(MCE_load_name_list))]], axis=1).max()

        # 1.2 * 평균값
        shear_force_H1_MCE_avg = 1.2 * shear_force_H1_MCE_max.mean(axis=1)
        shear_force_H2_MCE_avg = 1.2 * shear_force_H2_MCE_max.mean(axis=1)
        
    else : 
        shear_force_H1_MCE_avg = ''
        shear_force_H2_MCE_avg = ''

#%% V(축력) 값 뽑기

    # 축력 불러와서 Grouping
    axial_force_data = wall_SF_data[wall_SF_data['Load Case'].str.contains(gravity_load_name[0])]['V(kN)']

    # 절대값
    axial_force_abs = axial_force_data.abs()

    # result
    axial_force_abs.reset_index(inplace=True, drop=True)
    axial_force = axial_force_abs.groupby([[i//2 for i in range(0, len(axial_force_abs))]], axis=0).max()

#%% 결과 정리 후 Input Sheets에 넣기

# 출력용 Dataframe 만들기
    SF_output = pd.DataFrame()
    SF_output['Name'] = wall_SF_data['Name'].drop_duplicates()
    SF_output.reset_index(inplace=True, drop=True)

    SF_output['Nu'] = axial_force
    SF_output['1.2_DE_H1'] = shear_force_H1_DE_avg
    SF_output['1.2_DE_H2'] = shear_force_H2_DE_avg
    SF_output['1.2_MCE_H1'] = shear_force_H1_MCE_avg
    SF_output['1.2_MCE_H2'] = shear_force_H2_MCE_avg
       
    SF_output = pd.merge(SF_output, transfer_element_info, how='left')

    SF_output = SF_output.iloc[:,[0,6,7,8,9,10,11,12,13,14,1,2,3,4,5]] # SF_output 재정렬
    
# nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)

    SF_output = SF_output.replace(np.nan, '', regex=True)
    
# 엑셀로 출력(Using win32com)
    
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application') # 엑셀 실행
    excel.Visible = False # 엑셀창 안보이게
    
    wb = excel.Workbooks.Open(input_path + '\\' + input_xlsx)
    ws = wb.Sheets('Results_Wall')
    
    startrow, startcol = 5, 1
    
    ws.Range(ws.Cells(startrow, startcol),\
              ws.Cells(startrow + SF_output.shape[0]-1,\
                      startcol + SF_output.shape[1]-1)).Value\
    = list(SF_output.itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능
    
    wb.Close(SaveChanges=1) # Closing the workbook
    excel.Quit() # Closing the application 

#%% 그래프 process

    if graph == True:

#%% 파일 load

        # Wall 정보 load
        # wall_result_wb = openpyxl.load_workbook(input_path + '\\' + input_xlsx)
        wall_result = pd.read_excel(input_path +'\\' + input_xlsx,
                              sheet_name='Results_Wall', skiprows=3, header=0)
        
        wall_result = wall_result.iloc[:, [0, 25, 27, 29, 31]]
        wall_result.columns = ['Name', 'DE_H1', 'DE_H2', 'MCE_H1', 'MCE_H2']
        wall_result.reset_index(inplace=True, drop=True)
        
        # Story 정보에서 층이름만 뽑아내기
        story_name = story_info.iloc[:, 1]
        story_name.reset_index(drop=True, inplace=True)

#%% ***조작용 코드
        # wall_name_to_delete = ['84A-W1_1','84A-W3_1_40F'] 
        # # 지우고싶은 층들을 대괄호 안에 입력(벽 이름만 입력하면 벽 전체 다 없어짐, 벽+층 이름 입력하면 특정 층의 벽만 없어짐)
        
        # for i in wall_name_to_delete:
        #     wall_result = wall_result[wall_result['Name'].str.contains(i) == False]
        
#%% 벽체 해당하는 층 높이 할당
        floor = []
        for i in wall_result['Name']:
            floor.append(i.split('_')[-1])
        
        wall_result['Story Name'] = floor
        
        wall_result_output = pd.merge(wall_result, story_info.iloc[:,[1,2]], how='left')
        
#%% 그래프
        count = 1
        
        ### H1 DE 그래프 ###
        if len(DE_load_name_list) != 0:
        
            fig1 = plt.figure(count, dpi=150, figsize=(5,6))
            plt.xlim(0, xlim)
            plt.scatter(wall_result_output['DE_H1'], wall_result_output['Height(mm)'], color = 'k', s=1) # s=1 : point size
            
            # height값에 대응되는 층 이름으로 y축 눈금 작성
            plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
            
            plt.axvline(x= DCR_criteria, color='r', linestyle='--')
            plt.grid(linestyle='-.')
            plt.xlabel('D/C Ratios')
            plt.ylabel('Story')
            plt.title('Shear Strength (H1 DE)')
            
            plt.tight_layout()
            plt.close()
            count += 1
            
            yield fig1
            
            ### H2 DE 그래프 ###
            fig2 = plt.figure(count, dpi=150, figsize=(5,6))
            plt.xlim(0, xlim)
            plt.scatter(wall_result_output['DE_H2'], wall_result_output['Height(mm)'], color = 'k', s=1) # s=1 : point size
            
            # height값에 대응되는 층 이름으로 y축 눈금 작성
            plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
            
            plt.axvline(x= DCR_criteria, color='r', linestyle='--')
            plt.grid(linestyle='-.')
            plt.xlabel('D/C Ratios')
            plt.ylabel('Story')
            plt.title('Shear Strength (H2 DE)')
            
            plt.tight_layout()
            plt.close()  
            count += 1
            
            yield fig2
        
        ### H1 MCE 그래프 ###
        if len(MCE_load_name_list) != 0:
        
            fig3 = plt.figure(count, dpi=150, figsize=(5,6))
            plt.xlim(0, xlim)
            plt.scatter(wall_result_output['MCE_H1'], wall_result_output['Height(mm)'], color = 'k', s=1) # s=1 : point size
            
            # height값에 대응되는 층 이름으로 y축 눈금 작성
            plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
            
            plt.axvline(x= DCR_criteria, color='r', linestyle='--')
            plt.grid(linestyle='-.')
            plt.xlabel('D/C Ratios')
            plt.ylabel('Story')
            plt.title('Shear Strength (H1 MCE)')    
            
            plt.tight_layout()
            plt.close()
            count += 1
            
            yield fig3
            
            ### H2 MCE 그래프 ###
            fig4 = plt.figure(count, dpi=150, figsize=(5,6))
            plt.xlim(0, xlim)
            plt.scatter(wall_result_output['MCE_H2'], wall_result_output['Height(mm)'], color = 'k', s=1) # s=1 : point size
            
            # height값에 대응되는 층 이름으로 y축 눈금 작성
            plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
            
            plt.axvline(x= DCR_criteria, color='r', linestyle='--')
            plt.grid(linestyle='-.')
            plt.xlabel('D/C Ratios')
            plt.ylabel('Story')
            plt.title('Shear Strength (H2 MCE)')
            
            plt.tight_layout()
            plt.close()
            count += 1
        
            yield fig4
            
# #%% 부재의 위치별  V, M 값 확인을 위한 도면 작성
    
#     # 도면을 그리기 위한 Node List 만들기
#     node_map_z = SF_ongoing_max_avg['i-V'].drop_duplicates()
#     node_map_z.sort_values(ascending=False, inplace=True)
#     node_map_list = node_info_data[node_info_data['V'].isin(node_map_z)]
    
#     # 도면을 그리기 위한 Element List 만들기
#     element_map_list = pd.merge(SF_ongoing_max_avg, element_info_data.iloc[:,[1,5,6,8,9]]
#                                 , how='left', left_index=True, right_on='Element Name')
    
#     # V, M 크기에 따른 Color 지정
#     cmap_V = plt.get_cmap('Reds')
#     cmap_M = plt.get_cmap('YlOrBr')
    
#     # 층별 Loop
#     count = 1
#     for i in node_map_z:   
#         # 해당 층에 해당하는 Nodes와 Elements만 Extract
#         node_map_list_extracted = node_map_list[node_map_list['V'] == i]
#         element_map_list_extracted = element_map_list[element_map_list['i-V'] == i]
#         element_map_list_extracted.reset_index(inplace=True, drop=True)
        
#         # Colorbar, 그래프 Coloring을 위한 설정
#         norm_V = plt.Normalize(vmin = element_map_list_extracted['V2 max'].min()\
#                              , vmax = element_map_list_extracted['V2 max'].max())
#         cmap_V_elem = cmap_V(norm_V(element_map_list_extracted['V2 max']))
#         scalar_map_V = mpl.cm.ScalarMappable(norm_V, cmap_V)
        
#         norm_M = plt.Normalize(vmin = element_map_list_extracted['M3 max'].min()\
#                              , vmax = element_map_list_extracted['M3 max'].max())
#         cmap_M_elem = cmap_M(norm_M(element_map_list_extracted['M3 max']))
#         scalar_map_M = mpl.cm.ScalarMappable(norm_M, cmap_M)
        
#         ## V(전단)     
#         # Graph    
#         fig1 = plt.figure(count, dpi=150)
        
#         plt.scatter(node_map_list_extracted['H1'], node_map_list_extracted['H2'], color='k', s=1)
        
#         for idx, row in element_map_list_extracted.iterrows():
            
#             element_map_x = [row['i-H1'], row['j-H1']]
#             element_map_y = [row['i-H2'], row['j-H2']]
            
#             plt.plot(element_map_x, element_map_y, c = cmap_V_elem[idx])
        
#         # Colorbar 만들기
#         plt.colorbar(scalar_map_V, shrink=0.7, label='V(kN)')
    
#         # 기타
#         plt.axis('off')
#         plt.title(story_info['Story Name'][story_info['Height(mm)'] == i].iloc[0])

#         plt.tight_layout()   
#         plt.close()
#         count += 1
#         yield fig1
        
#         ## M(모멘트)     
#         # Graph    
#         fig2 = plt.figure(count, dpi=150)
        
#         plt.scatter(node_map_list_extracted['H1'], node_map_list_extracted['H2'], color='k', s=1)
        
#         for idx, row in element_map_list_extracted.iterrows():
            
#             element_map_x = [row['i-H1'], row['j-H1']]
#             element_map_y = [row['i-H2'], row['j-H2']]
            
#             plt.plot(element_map_x, element_map_y, c = cmap_M_elem[idx])
        
#         # Colorbar 만들기
#         plt.colorbar(scalar_map_M, shrink=0.7, label='M(kN-mm)')
    
#         # 기타
#         plt.axis('off')
#         plt.title(story_info['Story Name'][story_info['Height(mm)'] == i].iloc[0])

#         plt.tight_layout()   
#         plt.close()
#         count += 1
#         yield fig2    
        
#%% wall_SF (Graph)

def wall_SF_graph(input_path, input_xlsx, DCR_criteria=1, yticks=2, xlim=3):

#%% 파일 load

    # Wall 정보 load
    # wall_result = pd.read_excel(input_path +'\\' + input_xlsx,
    #                       sheet_name='Results_Wall', skiprows=3, header=0)
    
    # wall_result = wall_result.iloc[:, [0, 25, 27, 29, 31]]
    # wall_result.columns = ['Name', 'DE_H1', 'DE_H2', 'MCE_H1', 'MCE_H2']
    # wall_result.reset_index(inplace=True, drop=True)
    
    wall_result = pd.DataFrame()
    story_info = pd.DataFrame()

    input_xlsx_sheet = 'Results_Wall'
    input_data_raw = pd.ExcelFile(input_path + '\\' + input_xlsx)
    input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', input_xlsx_sheet], skiprows=3)

    story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
    wall_result = input_data_sheets[input_xlsx_sheet].iloc[:,[0,25,27,29,31]]
    story_info = story_info[::-1]
    story_info.reset_index(inplace=True, drop=True)

    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    wall_result.columns = ['Name', 'DE_H1', 'DE_H2', 'MCE_H1', 'MCE_H2']
    
    # Story 정보에서 층이름만 뽑아내기
    story_name = story_info.iloc[:, 1]
    story_name.reset_index(drop=True, inplace=True)

#%% ***조작용 코드
    # wall_name_to_delete = ['84A-W1_1','84A-W3_1_40F'] 
    # # 지우고싶은 층들을 대괄호 안에 입력(벽 이름만 입력하면 벽 전체 다 없어짐, 벽+층 이름 입력하면 특정 층의 벽만 없어짐)
    
    # for i in wall_name_to_delete:
    #     wall_result = wall_result[wall_result['Name'].str.contains(i) == False]
    
#%% 벽체 해당하는 층 높이 할당
    floor = []
    for i in wall_result['Name']:
        floor.append(i.split('_')[-1])
    
    wall_result.loc[:, 'Story Name'] = floor
    
    wall_result_output = pd.merge(wall_result, story_info.iloc[:,[1,2]], how='left')
    
#%% 그래프
    count = 1

    ### H1 DE 그래프 ###
    if wall_result['DE_H1'].isnull().all() == False:
    
        fig1 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        plt.scatter(wall_result_output['DE_H1'], wall_result_output['Height(mm)'], color = 'k', s=1) # s=1 : point size
        
        # height값에 대응되는 층 이름으로 y축 눈금 작성
        plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Shear Strength (H1 DE)')
        
        plt.tight_layout()
        plt.close()
        count += 1
        
        yield fig1
        
        ### H2 DE 그래프 ###
        fig2 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        plt.scatter(wall_result_output['DE_H2'], wall_result_output['Height(mm)'], color = 'k', s=1) # s=1 : point size
        
        # height값에 대응되는 층 이름으로 y축 눈금 작성
        plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Shear Strength (H2 DE)')
        
        plt.tight_layout()
        plt.close()  
        count += 1
        
        yield fig2
    
    ### H1 MCE 그래프 ###
    if wall_result['MCE_H1'].isnull().all() == False:
    
        fig3 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        plt.scatter(wall_result_output['MCE_H1'], wall_result_output['Height(mm)'], color = 'k', s=1) # s=1 : point size
        
        # height값에 대응되는 층 이름으로 y축 눈금 작성
        plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Shear Strength (H1 MCE)')    
        
        plt.tight_layout()
        plt.close()
        count += 1
        
        yield fig3
        
        ### H2 MCE 그래프 ###
        fig4 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        plt.scatter(wall_result_output['MCE_H2'], wall_result_output['Height(mm)'], color = 'k', s=1) # s=1 : point size
        
        # height값에 대응되는 층 이름으로 y축 눈금 작성
        plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Shear Strength (H2 MCE)')
        
        plt.tight_layout()
        plt.close()
        count += 1
    
        yield fig4

