import pandas as pd
import numpy as np
import os
import pickle
from collections import deque  # Double-ended Queue : 자료의 앞, 뒤 양 방향에서 자료를 추가하거나 제거가능
import matplotlib.pyplot as plt
import matplotlib as mpl
import win32com.client
import pythoncom
from decimal import *

#%% Wall Axial Strain

def WAS(self, wall_design_xlsx_path, max_criteria=0.04, min_criteria=-0.002, yticks=2, WAS_gage_group='AS', graph=True):
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
    ------
    
    References
    ----------
    .. [1] "철근콘크리트 건축구조물의 성능기반 내진설계 지침", 대한건축학회, p.44, 2021
    
    '''
#%% Load Data
    # Data Conversion Sheets
    story_info = self.story_info
    wall_info = self.wall_info
    rebar_info = self.rebar_info
    
    # Analysis Result Sheets
    AS_gage_data = self.wall_as_gage_data
    AS_result_data = self.wall_as_result_data
    node_data = self.node_data
    element_data = self.wall_data

    # Seismic Loads List
    load_name_list = self.load_name_list
    gravity_load_name = self.gravity_load_name
    seismic_load_name_list = self.seismic_load_name_list
    DE_load_name_list = self.DE_load_name_list
    MCE_load_name_list = self.MCE_load_name_list

    AS_result_data = AS_result_data.sort_values(by= ['Load Case', 'Element Name', 'Step Type']) # 여러개로 나눠돌릴 경우 순서가 섞여있을 수 있어 DE11~MCE72 순으로 정렬
    
    #%% (Only for SW2R Project) AS Gage가 분할층에서 나뉘어서 모델링 된 경우, 한 개의 게이지 값만 고려하기

    # Merge로 Node 번호에 맞는 좌표를 결합
    AS_gage_node_coord = pd.merge(AS_gage_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID', suffixes=(None, '_1'))
    AS_gage_node_coord = pd.merge(AS_gage_node_coord, node_data, how='left', left_on='J-Node ID', right_on='Node ID', suffixes=(None, '_2'))
    
    ### WAS gage가 분할층에서 나눠지지 않게 만들기 
    # 분할층 노드가 포함되지 않은 부재 slice
    AS_gage_node_coord_no_div = AS_gage_node_coord[(AS_gage_node_coord['V'].isin(story_info['Height(mm)']))\
                                                & (AS_gage_node_coord['V_2'].isin(story_info['Height(mm)']))]
    
    # 분할층 노드가 상부에만(j-node) 포함되는 부재 slice
    AS_gage_node_coord_div = AS_gage_node_coord[(AS_gage_node_coord['V'].isin(story_info['Height(mm)']))\
                                             & (~AS_gage_node_coord['V_2'].isin(story_info['Height(mm)']))]
        
    AS_gage_node_coord = AS_gage_node_coord.iloc[:,[0,1,2,3,7,11]]
    
    # AS_gage_node_coord_div 노드들의 상부 노드(j-node)의 z좌표를 다음 측으로 격상
    next_level_list = []
    for i in AS_gage_node_coord_div['V_2']:
        level_bigger = story_info['Height(mm)'][story_info['Height(mm)']-i >= 0]
        next_level = level_bigger.sort_values(ignore_index=True)[0]

        next_level_list.append(next_level)
    AS_gage_node_coord_div.loc[:, 'V_2'] = next_level_list
    
    next_node_list = []
    for idx, row in AS_gage_node_coord_div[['H1_2', 'H2_2', 'V_2']].iterrows():
        new_node = node_data[(node_data['H1'] == row[0]) 
                             & (node_data['H2'] == row[1]) 
                             & (node_data['V'] == row[2])]['Node ID']
        next_node_list.append(new_node)
    AS_gage_node_coord_div.loc[:, 'J-Node ID'] = next_node_list        
    
    AS_gage_data = pd.concat([AS_gage_node_coord_no_div, AS_gage_node_coord_div]\
                                , ignore_index=True)
    
    AS_gage_data = AS_gage_data.iloc[:,0:4]        
        
#%% Gage Data & Result에 Node 정보 매칭   
    # element_data 중, data conversion sheet에 있는 부재만 선별
    prop_name = wall_info.iloc[:,0]
    prop_name.name = 'Property Name'
    element_data = element_data[element_data['Property Name'].isin(prop_name)]
    
    ### 여러개로 나뉜 Wall Elements의 양 끝단 점(i,j,k,l) 알아내기 
    # 부재의 orientation 맞추기 (안맞추면 missing data 생기는 경우 발생)
    # Merge로 Node 번호에 맞는 좌표를 결합
    element_data = pd.merge(element_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID', suffixes=(None, '_i'))
    element_data = pd.merge(element_data, node_data, how='left', left_on='J-Node ID', right_on='Node ID', suffixes=(None, '_j'))

    # 허용 오차
    tolerance = 5 # mm
    element_data_pos = element_data[element_data['H1_j'] - element_data['H1'] > tolerance]
    element_data_neg = element_data[element_data['H1_j'] - element_data['H1'] < -tolerance]
    element_data_zero = element_data[(element_data['H1_j'] - element_data['H1'] >= -tolerance) 
                               & (element_data['H1_j'] - element_data['H1'] <= tolerance)]
    
    # node i <-> node j, node k <-> node l
    element_data_neg = element_data_neg.iloc[:,[0,1,3,2,5,4,6,7,8,9,10,11,12,13]]
    element_data_neg.columns = element_data_pos.columns.values
    
    # (X좌표가 같다면) Y 좌표가 더 작은 노드를 i-node로!
    element_data_zero_pos = element_data_zero[element_data_zero['H2_j'] >= element_data_zero['H2']]
    element_data_zero_neg = element_data_zero[element_data_zero['H2_j'] < element_data_zero['H2']]
    
    element_data_zero_neg = element_data_zero_neg.iloc[:,[0,1,3,2,5,4,6,7,8,9,10,11,12,13]]
    element_data_zero_neg.columns = element_data_zero_pos.columns.values
    
    # pos, neg 합치기
    element_data = pd.concat([element_data_pos, element_data_neg, element_data_zero_pos, element_data_zero_neg]\
                          , ignore_index=True)

    # 필요한 열 뽑고 재정렬
    element_data = element_data.iloc[:,[0,1,2,3,4,5]]
        
    # 같은 Property Name에 따라 Sorting
    element_data_sorted = element_data.set_index('Property Name')
    element_data_sorted = element_data_sorted.sort_values('Property Name')

    # Wall의 4개의 node를 하나의 리스트로 합치기
    element_data_sorted['Node List'] = element_data_sorted\
        .loc[:,['I-Node ID', 'J-Node ID', 'K-Node ID', 'L-Node ID']].values.tolist()
    
    # For Loop 돌리면서 Property Name에 따라 Node 리스트 업데이트 (겹치는거 제거하면서)
    count = 0
    element_data_updated = pd.DataFrame(columns=['Property Name','Node List'])
    for idx, elem_node_data in element_data_sorted.groupby('Property Name')['Node List']:

        # series -> list
        elem_node_data = list(elem_node_data)
        # deque 생성
        elem_node_dq = deque()
        
        # 노드를 위치 순서대로 deque에 insert
        for i in range(0, len(elem_node_data)):
            elem_node_dq.insert(int(i*1 + 0), elem_node_data[i][0])
            elem_node_dq.insert(int(i*2 + 1), elem_node_data[i][1])
            elem_node_dq.insert(int(i*3 + 2), elem_node_data[i][2])
            elem_node_dq.insert(int(i*4 + 3), elem_node_data[i][3])
        elem_node_dq = list(elem_node_dq) 
        
        # count = 1인 노드만 추출(중복되는 노드들 제거하는 법 몰라서 우회함)
        elem_node_dq_flat = []
        for i in elem_node_dq:
            if elem_node_dq.count(i) == 1:
                elem_node_dq_flat.append(i)

        # 합쳐져 있는 노드들 분류
        for i in range(0, len(elem_node_dq_flat)//4):
            temp = [elem_node_dq_flat[i + len(elem_node_dq_flat)//4*0]
                    , elem_node_dq_flat[i + len(elem_node_dq_flat)//4*1]
                    , elem_node_dq_flat[i + len(elem_node_dq_flat)//4*2]
                    , elem_node_dq_flat[i + len(elem_node_dq_flat)//4*3]]
            # pd.at = Access a single value for a row/column label pair
            element_data_updated.at[count,'Node List'] = temp
        element_data_updated.at[count,'Property Name'] = idx
        count += 1
    
    # list로 되어있는 i,j,k,l node들 각 column으로 나누기
    element_data_updated[['i Node', 'j Node', 'k Node', 'l Node']]\
        = pd.DataFrame(element_data_updated['Node List'].tolist())
    
    ### Gage 데이터 정리하기
    # Group Name에 따라서 데이터 추출
    AS_result_data = AS_result_data[AS_result_data['Group Name'] == WAS_gage_group]
    AS_gage_data = AS_gage_data[AS_gage_data['Group Name'] == WAS_gage_group]

    # 혹시 데이터 중복으로 들어간 경우, drop
    AS_gage_data = AS_gage_data.drop_duplicates()
    node_data = node_data.drop_duplicates()
    
    # 지진하중 결과 & Performance Level=1만 포함시키기
    AS_result_data = AS_result_data[AS_result_data['Load Case']\
                                    .str.contains('|'.join(seismic_load_name_list + gravity_load_name))]
    AS_result_data = AS_result_data[AS_result_data['Performance Level'] == 1]
    
    # Max나 Min값이 여러개인 경우(performance level이 여러개일 때로 추정), 큰 값만 뽑기
    AS_result_data['Axial Strain(abs)'] = AS_result_data['Axial Strain'].copy().abs()
    AS_result_data = AS_result_data.sort_values(by='Axial Strain(abs)')
    AS_result_data = AS_result_data.drop_duplicates(subset=['Element Name', 'Load Case', 'Step Type'], keep='last')
    
    # 필요한 정보만 추출
    AS_result_data = AS_result_data.iloc[:,[1,2,3,4]]
    AS_result_data.reset_index(inplace=True, drop=True)
    
    # Element Name별로 Grouping
    AS_result_grouped_list = list(AS_result_data.groupby(['Load Case', 'Step Type']))
    
    ### 벽체 정보에 Gage Result를 붙이기 위해, 우선 Gage Result data 정리하기   
    # 해석 결과 상관없이 Full 지진하중 이름 list 만들기
    full_DE_load_name_list = 'DE' + pd.Series([11,12,21,22,31,32,41,42,51,52,61,62,71,72]).astype(str)
    full_MCE_load_name_list = 'MCE' + pd.Series([11,12,21,22,31,32,41,42,51,52,61,62,71,72]).astype(str)
    full_load_name_list = pd.concat([full_DE_load_name_list, full_MCE_load_name_list])
    
    # 지진하중, Max/Min loop 돌리면서 reshape된 df만들기
    elem_name = AS_result_data['Element Name'].drop_duplicates().sort_values()
    elem_name.name = 'Element Name'    
    elem_name.reset_index(inplace=True, drop=True)
    AS_result_reshaped = pd.DataFrame(elem_name)
    for load_name in full_load_name_list:
        for max_min in ['Max', 'Min']:
            # 만들어진 Group List loop 돌리기
            for AS_result_grouped in AS_result_grouped_list:
                if (load_name in AS_result_grouped[0][0]) &  (AS_result_grouped[0][1] == max_min):
                    # Element Name 순서대로 재정렬
                    # AS_result_grouped_df = AS_result_grouped.sort_values(by='Element Name')
                    AS_result_grouped_df = pd.merge(elem_name, AS_result_grouped[1], how='left')
                    AS_result_grouped_df.reset_index(inplace=True, drop=True)
                    AS_result_reshaped = pd.concat([AS_result_reshaped, AS_result_grouped_df['Axial Strain']], axis=1)
                   
            # 해당 지진하중의 해석결과가 없는 경우 Blank Column 생성
            if load_name not in seismic_load_name_list: 
                blank_col = pd.Series([''] * len(elem_name))
                AS_result_reshaped = pd.concat([AS_result_reshaped, blank_col], axis=1)    
    
    # 중력하중, Max/Min reshape된 df 합치기
    for max_min in ['Max', 'Min']:
        # 중력하중의 해석결과가 없는 경우 Blank Column 생성
        if len(gravity_load_name) == 0:
            blank_col = pd.Series([''] * len(elem_name))
            AS_result_reshaped = pd.concat([AS_result_reshaped, blank_col], axis=1) 
        else: # 만들어진 Group List loop 돌리기            
            for AS_result_grouped in AS_result_grouped_list:
                if (gravity_load_name[0] in AS_result_grouped[0][0]) &  (AS_result_grouped[0][1] == max_min):
                    # Element Name 순서대로 재정렬
                    # AS_result_grouped_df = AS_result_grouped.sort_values(by='Element Name')
                    AS_result_grouped_df = pd.merge(elem_name, AS_result_grouped[1], how='left')
                    AS_result_grouped_df.reset_index(inplace=True, drop=True)
                    AS_result_reshaped = pd.concat([AS_result_reshaped, AS_result_grouped_df['Axial Strain']], axis=1)
    
    ### element_data_updated에 AS_gage_data 합치기
    # 벽체의 i,l 노드와 일치하는 AS gage 합치기
    AS_result_merged = pd.merge(element_data_updated, AS_gage_data, how='left'
                                , left_on=['i Node', 'l Node'], right_on=['I-Node ID', 'J-Node ID'])
    AS_result_merged = AS_result_merged.iloc[:,[0,2,3,4,5,7]]
    # 벽체의 j,k 노드와 일치하는 AS gage 합치기
    AS_result_merged = pd.merge(AS_result_merged, AS_gage_data, how='left'
                                , left_on=['j Node', 'k Node'], right_on=['I-Node ID', 'J-Node ID'])
    AS_result_merged = AS_result_merged.iloc[:,np.r_[0:6,7]]
    AS_result_merged.columns = ['Property Name', 'i Node', 'j Node', 'k Node', 'l Node', 'i-l Element', 'j-k Element']
    
    ### 만들어진 AS_output에 AS_result_data 합치기
    AS_i_result_merged = pd.merge(AS_result_merged, AS_result_reshaped, how='left'
                           , left_on='i-l Element', right_on='Element Name')
    AS_j_result_merged = pd.merge(AS_result_merged, AS_result_reshaped, how='left'
                           , left_on='j-k Element', right_on='Element Name')
    AS_i_output = AS_i_result_merged.iloc[:,8:]
    AS_j_output = AS_j_result_merged.iloc[:,8:]
    
    # 최종 dataframe을 우선 이름만 넣고 생성
    AS_output = pd.DataFrame(AS_result_merged['Property Name'])
    
    # 최종 dataframe에 i-end와 j-end의 결과를 번갈아가면서 넣어 최종 df 생성
    for i in range(int(AS_i_output.shape[1] / 2)):
        max_col = int(2*i)
        min_col = int(2*i + 1)
        AS_output = pd.concat([AS_output, AS_i_output.iloc[:,[max_col, min_col]]], axis=1)
        AS_output = pd.concat([AS_output, AS_j_output.iloc[:,[max_col, min_col]]], axis=1)
    
    # Input 시트의 부재 순서대로 재정렬
    AS_output = pd.merge(prop_name, AS_output, how='left')
    
#%% 결과 생성 후 Seismic Design 시트에 넣기

    # 출력용 Dataframe 만들기
    # Design_S.Wall 시트
    steel_design_df = wall_info.iloc[:,[5,6,7,9,10]]
    wall_output = pd.concat([wall_info.iloc[:,0:11], steel_design_df], axis=1)
    
    # Table_S.Wall_DE 시트
    # Ground Level(0mm, 1F)에 가장 가까운 층의 row index get
    ground_level_idx = story_info['Height(mm)'].abs().idxmin()
    # story_info의 Index열을 1부터 시작하도록 재지정
    story_info['Index'] = range(story_info.shape[0], 0, -1)
    # Ground Level(0mm, 1F)에 가장 가까운 층을 index 5에 배정
    add_num_new_story = 5 - story_info.iloc[ground_level_idx, 0]
    story_info['Index'] = story_info['Index'] + add_num_new_story
    
    # Wall 이름 split    
    wall_info[['Wall Name', 'Wall Number', 'Story Name']] = wall_info['Name'].str.split('_', expand=True)
    wall_info = pd.merge(wall_info, story_info, how='left')
    # 결과값 없는 부재 제거
    idx_to_slice = AS_output.iloc[:,1:].dropna().index # dropna로 결과값 있는 부재만 남긴 후 idx 추출
    idx_to_slice2 = wall_info['Name'].iloc[idx_to_slice].index # 결과값 있는 부재만 slice 후 idx 추출
    wall_info = wall_info.iloc[idx_to_slice2,:]
    wall_info.reset_index(inplace=True, drop=True)
    # 벽체 이름, 번호에 따라 grouping
    wall_name_list = list(wall_info.groupby(['Wall Name', 'Wall Number'], sort=False))
    # 55 row짜리 empty dataframe 만들기
    name_empty = pd.DataFrame(np.nan, index=range(55), columns=range(len(wall_name_list)))
    # dataframe에 이름 채워넣기
    count = 0
    while True:
        name_iter = wall_name_list[count][0][0]
        num_iter = wall_name_list[count][0][1]
        total_iter = wall_info['Name'][(wall_info['Wall Name'] == name_iter) 
                                       & (wall_info['Wall Number'] == num_iter)]
        idx_range = wall_info['Index'][(wall_info['Wall Name'] == name_iter) 
                                       & (wall_info['Wall Number'] == num_iter)]
        name_empty.iloc[idx_range, count] = total_iter
        
        count += 1
        if count == len(wall_name_list):
            break
    # dataframe을 1열로 만들기
    name_output_arr = np.array(name_empty)
    name_output_arr = np.reshape(name_output_arr, (-1, 1), order='F')
    name_output = pd.DataFrame(name_output_arr)    
    
    # ETC 시트
    rebar_output = rebar_info.iloc[:,1:]
    
    # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
    AS_output = AS_output.replace(np.nan, '', regex=True)
    AS_G_output = AS_output.iloc[:, [113,114,115,116]]
    AS_output = AS_output.iloc[:,0:113]
    wall_output = wall_output.replace(np.nan, '', regex=True)
    name_output = name_output.replace(np.nan, '', regex=True)
    rebar_output = rebar_output.replace(np.nan, '', regex=True)
    
    # 엑셀로 출력(Using win32com)
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게
    
    wb = excel.Workbooks.Open(wall_design_xlsx_path)
    ws1 = wb.Sheets('Results_S.Wall_Strain')
    ws2 = wb.Sheets('Design_S.Wall')
    ws3 = wb.Sheets('Table_S.Wall_DE')
    ws4 = wb.Sheets('ETC')
    
    startrow, startcol = 5, 1
    
    # Results_S.Wall_Strain 시트 입력
    # 값을 입력하기 전에, 우선 해당 셀에 있는 값 지우기
    ws1.Range('A%s:DI%s' %(startrow, 5000)).ClearContents()
    ws1.Range('A%s:DI%s' %(startrow, startrow + AS_output.shape[0] - 1)).Value\
        = list(AS_output.itertuples(index=False, name=None))
        
    # Results_S.Wall_Strain 시트 입력 (중력하중)
    # 값을 입력하기 전에, 우선 해당 셀에 있는 값 지우기
    ws1.Range('FZ%s:GC%s' %(startrow, 5000)).ClearContents()
    ws1.Range('FZ%s:GC%s' %(startrow, startrow + AS_G_output.shape[0] - 1)).Value\
        = list(AS_G_output.itertuples(index=False, name=None))
    
    # Design_S.Wall 시트 입력
    ws2.Range('A%s:P%s' %(startrow, 5000)).ClearContents()
    ws2.Range('A%s:P%s' %(startrow, startrow + wall_output.shape[0] - 1)).Value\
        = list(wall_output.itertuples(index=False, name=None))
    
    # Table_S.Wall_DE 시트 입력
    ws3.Range('B%s:B%s' %(startrow, 5000)).ClearContents()
    ws3.Range('B%s:B%s' %(startrow, startrow + name_output.shape[0] - 1)).Value\
        = [[i] for i in name_output[0]] # series -> list 형식만 입력가능
    ws3.Range('A4:A4').Value\
        = len(wall_name_list) # series -> list 형식만 입력가능
    
    # Design_S.Wall 시트 입력
    ws4.Range('D%s:L%s' %(startrow, 5000)).ClearContents()
    ws4.Range('D%s:L%s' %(startrow, startrow + rebar_output.shape[0] - 1)).Value\
        = list(rebar_output.itertuples(index=False, name=None))
        
    wb.Save()
     
#%% ***조작용 코드
    # 데이터 없애기 위한 기준값 입력
    # AS_output = AS_output.drop(AS_output[(AS_output.loc[:,'DE_min_avg'] < -0.002)].index)
    # AS_output = AS_output.drop(AS_output[(AS_output.loc[:,'MCE_min_avg'] < -0.002)].index)
    # .....위와 같은 포맷으로 계속

#%% 그래프
    if graph == True:
        # Wall 정보 load
        ws_DE = wb.Sheets('Table_S.Wall_DE')
        ws_MCE = wb.Sheets('Table_S.Wall_MCE')

        DE_result = ws_DE.Range('U%s:V%s' %(startrow, startrow + name_output.shape[0] - 1)).Value
        DE_result_arr = np.array(DE_result)[:,[0,1]]
        MCE_result = ws_MCE.Range('U%s:V%s' %(startrow, startrow + name_output.shape[0] - 1)).Value
        MCE_result_arr = np.array(MCE_result)[:,[0,1]]
        
        WAS_plot = name_output.copy()
        WAS_plot[['DE(Compressive)', 'DE(Tensile)']] = DE_result_arr
        WAS_plot[['MCE(Compressive)', 'MCE(Tensile)']] = MCE_result_arr
        WAS_plot.columns = ['Name', 'DE(Compressive)', 'DE(Tensile)', 'MCE(Compressive)', 'MCE(Tensile)']
        
        # 벽체 해당하는 층 높이 할당
        story = []
        for i in WAS_plot['Name']:
            if i == '':
                story.append(np.nan)
            else:
                story.append(i.split('_')[-1])        
        WAS_plot['Story Name'] = story
        
        WAS_plot = pd.merge(WAS_plot, story_info.iloc[:,[1,2]], how='left')
        
        # Change non-numeric objects(e.g. str) into int or float as appropriate.
        WAS_plot['DE(Compressive)'] = pd.to_numeric(WAS_plot['DE(Compressive)'])
        WAS_plot['DE(Tensile)'] = pd.to_numeric(WAS_plot['DE(Tensile)'])
        WAS_plot['MCE(Compressive)'] = pd.to_numeric(WAS_plot['MCE(Compressive)'])
        WAS_plot['MCE(Tensile)'] = pd.to_numeric(WAS_plot['MCE(Tensile)'])
        
        # Delete rows with missing name
        WAS_plot = WAS_plot[WAS_plot['Name'] != '']        
        
        # 결과 dataframe -> pickle
        WAS_result = []
        WAS_result.append(WAS_plot)
        WAS_result.append(story_info)
        WAS_result.append(DE_load_name_list)
        WAS_result.append(MCE_load_name_list)
        with open('pkl/WAS.pkl', 'wb') as f:
            pickle.dump(WAS_result, f)
        
#%% Shear Wall Rotation (DCR)

def WR(self, input_xlsx_path, wall_design_xlsx_path, graph=True, DCR_criteria=1, yticks=2, xlim=3):
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
#%% Load Data
    # Data Conversion Sheets
    story_info = self.story_info
    wall_info = self.wall_info
    rebar_info = self.rebar_info

    story_info.reset_index(inplace=True, drop=True)
    wall_info.reset_index(inplace=True, drop=True)

    # Analysis Result Sheets
    node_data = self.node_data
    element_data = self.wall_data
    wall_SF_data = self.shear_force_data
    gage_data = self.wall_rot_gage_data
    wall_rot_data = self.wall_rot_result_data

    # Seismic Loads List
    load_name_list = self.load_name_list
    gravity_load_name = self.gravity_load_name
    seismic_load_name_list = self.seismic_load_name_list
    DE_load_name_list = self.DE_load_name_list
    MCE_load_name_list = self.MCE_load_name_list
    
    # story_info = result.story_info
    # wall_info = result.wall_info
    # rebar_info = result.rebar_info

    # story_info.reset_index(inplace=True, drop=True)
    # wall_info.reset_index(inplace=True, drop=True)

    # # Analysis Result Sheets
    # node_data = result.node_data
    # element_data = result.wall_data
    # wall_SF_data = result.shear_force_data
    # gage_data = result.wall_rot_gage_data
    # wall_rot_data = result.wall_rot_result_data

    # # Seismic Loads List
    # load_name_list = result.load_name_list
    # gravity_load_name = result.gravity_load_name
    # seismic_load_name_list = result.seismic_load_name_list
    # DE_load_name_list = result.DE_load_name_list
    # MCE_load_name_list = result.MCE_load_name_list

    # 필요없는 전단력 제거(층전단력)
    wall_SF_data = wall_SF_data[wall_SF_data['Name'].str.count('_') == 2] # underbar가 두개 들어간 행만 선택        
    wall_SF_data.reset_index(inplace=True, drop=True)

########### wall_SF와 동일 #####################################################
#%% 중력하중에 대한 전단력 데이터 grouping

    shear_force_G_data_grouped = pd.DataFrame()
    
    # G를 max, min으로 grouping
    for load_name in gravity_load_name:
        shear_force_G_data_grouped['G_H1_max'] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                 (wall_SF_data['Step Type'] == 'Max')]['H1(kN)'].values
            
        shear_force_G_data_grouped['G_H1_min'] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                 (wall_SF_data['Step Type'] == 'Min')]['H1(kN)'].values

        shear_force_G_data_grouped['G_H2_max'] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                 (wall_SF_data['Step Type'] == 'Max')]['H2(kN)'].values
            
        shear_force_G_data_grouped['G_H2_min'] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                 (wall_SF_data['Step Type'] == 'Min')]['H2(kN)'].values   

#%% DE, MCE에 대한 전단력 데이터 Grouping

    shear_force_DE_data_grouped = pd.DataFrame()
    shear_force_MCE_data_grouped = pd.DataFrame()

    ###
    full_DE_load_name_list = 'DE' + pd.Series([11,12,21,22,31,32,41,42,51,52,61,62,71,72]).astype(str)
    full_MCE_load_name_list = 'MCE' + pd.Series([11,12,21,22,31,32,41,42,51,52,61,62,71,72]).astype(str)
    
    # DE를 max, min으로 grouping
    for load_name in full_DE_load_name_list:
        if load_name in DE_load_name_list:
            shear_force_DE_data_grouped['{}_H1_max'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                          (wall_SF_data['Step Type'] == 'Max')]['H1(kN)'].values
                
            shear_force_DE_data_grouped['{}_H1_min'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                          (wall_SF_data['Step Type'] == 'Min')]['H1(kN)'].values
    
            shear_force_DE_data_grouped['{}_H2_max'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                          (wall_SF_data['Step Type'] == 'Max')]['H2(kN)'].values
                
            shear_force_DE_data_grouped['{}_H2_min'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                          (wall_SF_data['Step Type'] == 'Min')]['H2(kN)'].values   

        else:
            shear_force_DE_data_grouped['{}_H1_max'.format(load_name)] = ''
            shear_force_DE_data_grouped['{}_H1_min'.format(load_name)] = ''
            shear_force_DE_data_grouped['{}_H2_max'.format(load_name)] = ''
            shear_force_DE_data_grouped['{}_H2_min'.format(load_name)] = ''

    # MCE를 max, min으로 grouping
    for load_name in full_MCE_load_name_list:
        if load_name in MCE_load_name_list:
            shear_force_MCE_data_grouped['{}_H1_max'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                          (wall_SF_data['Step Type'] == 'Max')]['H1(kN)'].values
                
            shear_force_MCE_data_grouped['{}_H1_min'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                          (wall_SF_data['Step Type'] == 'Min')]['H1(kN)'].values
    
            shear_force_MCE_data_grouped['{}_H2_max'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                          (wall_SF_data['Step Type'] == 'Max')]['H2(kN)'].values
                
            shear_force_MCE_data_grouped['{}_H2_min'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                          (wall_SF_data['Step Type'] == 'Min')]['H2(kN)'].values   

        else:
            shear_force_MCE_data_grouped['{}_H1_max'.format(load_name)] = ''
            shear_force_MCE_data_grouped['{}_H1_min'.format(load_name)] = ''
            shear_force_MCE_data_grouped['{}_H2_max'.format(load_name)] = ''
            shear_force_MCE_data_grouped['{}_H2_min'.format(load_name)] = ''

#%% V(축력) 값 뽑기

    # 축력 불러와서 Grouping
    axial_force_data = wall_SF_data[wall_SF_data['Load Case'].str.contains(gravity_load_name[0])]['V(kN)']

    # result
    axial_force_data.reset_index(inplace=True, drop=True)
    axial_force = axial_force_data.groupby([[i//2 for i in range(0, len(axial_force_data))]], axis=0).min()
    
#%% 결과 정리 후 Input Sheets에 넣기

    # 출력용 Dataframe 만들기
    # Results_S.Wall_Shear 시트
    SF_output = pd.DataFrame()
    SF_output['Name'] = wall_SF_data['Name'].drop_duplicates()
    SF_output.reset_index(inplace=True, drop=True)

    SF_output['Nu'] = axial_force
    SF_output = pd.concat([SF_output, shear_force_G_data_grouped
                           , shear_force_DE_data_grouped, shear_force_MCE_data_grouped], axis=1)
    
    # wall_info 순서에 맞게 sort
    SF_output = pd.merge(wall_info['Name'], SF_output, how='left')
    SF_output = SF_output.dropna(subset='Nu')
    
    # Design_S.Wall 시트
    steel_design_df = wall_info.iloc[:,[5,6,7,9,10]]
    wall_output = pd.concat([wall_info.iloc[:,0:11], steel_design_df], axis=1)
    
    # ETC 시트
    rebar_output = rebar_info.iloc[:,1:]
    
    # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
    SF_output = SF_output.replace(np.nan, '', regex=True)
    wall_output = wall_output.replace(np.nan, '', regex=True)
    rebar_output = rebar_output.replace(np.nan, '', regex=True)
    
# 엑셀로 출력(Using win32com)
    
    # Using win32com...
    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게
    
    wb = excel.Workbooks.Open(wall_design_xlsx_path)
    ws1 = wb.Sheets('Results_S.Wall_Shear')
    ws2 = wb.Sheets('Design_S.Wall')
    ws3 = wb.Sheets('Table_S.Wall_DE')
    ws4 = wb.Sheets('ETC')
    ws5 = wb.Sheets('Results_S.Wall_Rotation')
    
    startrow, startcol = 5, 1
    
    # Results_S.Wall_Shear 시트 입력
    # 값을 입력하기 전에, 우선 해당 셀에 있는 값 지우기
    ws1.Range('A%s:DN%s' %(startrow, 5000)).ClearContents()
    ws1.Range('A%s:DN%s' %(startrow, startrow + SF_output.shape[0] - 1)).Value\
        = list(SF_output.itertuples(index=False, name=None))
    
    # Design_S.Wall 시트 입력
    ws2.Range('A%s:P%s' %(startrow, 5000)).ClearContents()
    ws2.Range('A%s:P%s' %(startrow, startrow + wall_output.shape[0] - 1)).Value\
        = list(wall_output.itertuples(index=False, name=None))
    
    # Design_S.Wall 시트 입력
    ws4.Range('D%s:L%s' %(startrow, 5000)).ClearContents()
    ws4.Range('D%s:L%s' %(startrow, startrow + rebar_output.shape[0] - 1)).Value\
        = list(rebar_output.itertuples(index=False, name=None))
        
    # wb.Save()
    # wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application 
    ###########################################################################

#%% Gage Data & Result에 Node 정보 매칭
    
    gage_data = gage_data.drop_duplicates()
    node_data = node_data.drop_duplicates()

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
    getcontext().rounding = ROUND_HALF_UP # 현재 사용중인 컨텍스트의 반올림 모드를 Round Half Up으로 설정
    def cos_sim(arr, unit_arr):
        result = np.dot(arr, unit_arr) / (np.linalg.norm(arr, axis=1)*np.linalg.norm(unit_arr))
        result_list = []
        for i in result:
            result_round = round(Decimal(i),0) # 반올림
            result_list.append(int(result_round)) # int로 변환
        return result_list
           
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
    gage_data = gage_data.join(element_data.set_index(['J-Node ID', 'Similarity ji-e1', 'Similarity ji-e2'])\
                               ['Property Name'], on=['I-Node ID', 'Similarity ij-e1', 'Similarity ij-e2'])
    gage_data.rename({'Property Name' : 'gage_name'}, axis=1, inplace=True)
    
    # 위에서 join한 두 가지 경우의 이름 열 합치기
    for i in range(len(gage_data)):
        if pd.isnull(gage_data.iloc[i, 18]):
            gage_data.iloc[i, 18] = gage_data.iloc[i, 19]
    
    gage_data = gage_data.iloc[:, 0:19]
    
    
    wall_rot_data = wall_rot_data[wall_rot_data['Load Case']\
                                  .str.contains('|'.join(seismic_load_name_list + gravity_load_name))]
    # 중복되는 데이터 제거
    wall_rot_data = wall_rot_data.drop_duplicates()
    
    ### SWR gage data와 SWR result data 연결하기(Element Name 기준으로)
    wall_rot_data = wall_rot_data.join(gage_data.set_index('Element Name')['gage_name'], on='Element Name')    
        
    ### SWR_total data 만들기
    SWR_max = wall_rot_data[(wall_rot_data['Step Type'] == 'Max') & (wall_rot_data['Performance Level'] == 1)][['Rotation']].values # dataframe을 array로
    SWR_max_gagename = wall_rot_data[(wall_rot_data['Step Type'] == 'Max') & (wall_rot_data['Performance Level'] == 1)][['gage_name']].values # dataframe을 array로
    SWR_max = SWR_max.reshape(gage_num, len(DE_load_name_list)+len(MCE_load_name_list)+len(gravity_load_name)
                              , order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
    SWR_max_gagename = SWR_max_gagename.reshape(gage_num, len(DE_load_name_list)+len(MCE_load_name_list)+len(gravity_load_name)
                                                , order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
    SWR_max = pd.DataFrame(SWR_max) # array를 다시 dataframe으로
    SWR_max_gagename = pd.DataFrame(SWR_max_gagename) # array를 다시 dataframe으로
    
    SWR_min = wall_rot_data[(wall_rot_data['Step Type'] == 'Min') & (wall_rot_data['Performance Level'] == 1)][['Rotation']].values
    SWR_min_gagename = wall_rot_data[(wall_rot_data['Step Type'] == 'Min') & (wall_rot_data['Performance Level'] == 1)][['gage_name']].values
    SWR_min = SWR_min.reshape(gage_num, len(DE_load_name_list)+len(MCE_load_name_list)+len(gravity_load_name), order='F')
    SWR_min_gagename = SWR_min_gagename.reshape(gage_num, len(DE_load_name_list)+len(MCE_load_name_list)+len(gravity_load_name), order='F')
    SWR_min = pd.DataFrame(SWR_min)
    SWR_min_gagename = pd.DataFrame(SWR_min_gagename)
    
    SWR_total = pd.concat([gage_data[['I_V', 'I_H1', 'I_H2', 'J_H1', 'J_H2']], SWR_max_gagename.iloc[:,0]], axis=1)
    for i in range(SWR_max.shape[1]): # DE11_max, DE11_min, DE12_max, DE12_min...순으로 concat
        SWR_total = pd.concat([SWR_total, SWR_max.iloc[:,i], SWR_min.iloc[:,i]], axis=1)
    
    #SWR_total 의 column 명 만들기
    SWR_total_column_name = []
    for load_name in seismic_load_name_list + gravity_load_name:
        SWR_total_column_name.extend([load_name + '_max'])
        SWR_total_column_name.extend([load_name + '_min'])
    
    SWR_total.columns = ['Height', 'i_X', 'i_Y', 'j_X', 'j_Y', 'Name'] + SWR_total_column_name
    
    # 해석 결과가 없는 지진파에 대해 blank column 만들기
    full_DE_load_name_list = 'DE' + pd.Series([11,12,21,22,31,32,41,42,51,52,61,62,71,72]).astype(str)
    full_MCE_load_name_list = 'MCE' + pd.Series([11,12,21,22,31,32,41,42,51,52,61,62,71,72]).astype(str)
    full_load_name_list = pd.concat([full_DE_load_name_list, full_MCE_load_name_list])
    
    SWR_total['Blank'] = ''
    
    # SWR_output = 해석 결과값 only
    SWR_output = SWR_total.loc[:,'Name']
    for load_name in full_load_name_list:
        if load_name in seismic_load_name_list:
            load_name_max = load_name + '_max'
            load_name_min = load_name + '_min'
            SWR_output = pd.concat([SWR_output, SWR_total[load_name_max], SWR_total[load_name_min]], axis=1)
        else:
            SWR_output = pd.concat([SWR_output, SWR_total['Blank'], SWR_total['Blank']], axis=1)
            
    if len(gravity_load_name) == 0:
        SWR_output = pd.concat([SWR_output, SWR_total['Blank'], SWR_total['Blank']], axis=1)
    
    else:
        load_name_max = gravity_load_name[0] + '_max'
        load_name_min = gravity_load_name[0] + '_min'
        SWR_output = pd.concat([SWR_output, SWR_total[load_name_max], SWR_total[load_name_min]], axis=1)
    
    # wall_info 순서에 맞게 sort
    SWR_output = pd.merge(wall_info['Name'], SWR_output, how='left')

#%% 결과 정리 후 Input Sheets에 넣기

    # Table_S.Wall_DE 시트
    # Ground Level(0mm, 1F)에 가장 가까운 층의 row index get
    ground_level_idx = story_info['Height(mm)'].abs().idxmin()
    # story_info의 Index열을 1부터 시작하도록 재지정
    story_info['Index'] = range(story_info.shape[0], 0, -1)
    # Ground Level(0mm, 1F)에 가장 가까운 층을 index 5에 배정
    add_num_new_story = 5 - story_info.iloc[ground_level_idx, 0]
    story_info['Index'] = story_info['Index'] + add_num_new_story
    
    # Wall 이름 split    
    wall_info[['Wall Name', 'Wall Number', 'Story Name']] = wall_info['Name'].str.split('_', expand=True)
    wall_info = pd.merge(wall_info, story_info, how='left')
    # 결과값 없는 부재 제거
    idx_to_slice = SWR_output.iloc[:,1:].dropna().index # dropna로 결과값(DE,MCE) 있는 부재만 남긴 후 idx 추출
    name_to_slice = SWR_output['Name'].iloc[idx_to_slice]
    idx_to_slice2 = wall_info[wall_info['Name'].isin(name_to_slice)].index # 결과값 있는 부재만 slice 후 idx 추출
    wall_info = wall_info.iloc[idx_to_slice2,:]
    wall_info.reset_index(inplace=True, drop=True)
    # 벽체 이름, 번호에 따라 grouping
    wall_name_list = list(wall_info.groupby(['Wall Name', 'Wall Number'], sort=False))
    # 55 row짜리 empty dataframe 만들기
    name_empty = pd.DataFrame(np.nan, index=range(55), columns=range(len(wall_name_list)))
    # dataframe에 이름 채워넣기
    count = 0
    while True:
        name_iter = wall_name_list[count][0][0]
        num_iter = wall_name_list[count][0][1]
        total_iter = wall_info['Name'][(wall_info['Wall Name'] == name_iter) 
                                       & (wall_info['Wall Number'] == num_iter)]
        idx_range = wall_info['Index'][(wall_info['Wall Name'] == name_iter) 
                                       & (wall_info['Wall Number'] == num_iter)]
        name_empty.iloc[idx_range, count] = total_iter
        
        count += 1
        if count == len(wall_name_list):
            break
    # dataframe을 1열로 만들기
    name_output_arr = np.array(name_empty)
    name_output_arr = np.reshape(name_output_arr, (-1, 1), order='F')
    name_output = pd.DataFrame(name_output_arr)
    
    # 정제된 SWR_output 값에 맞는 좌표값을 SWR_output에 merge
    SWR_output = pd.merge(SWR_output, SWR_total[['Name', 'i_X', 'i_Y', 'j_X', 'j_Y']], how='left')
    
    # nan인 칸을 ''로 바꿔주기
    SWR_output = SWR_output.replace(np.nan, '', regex=True)
    SWR_G_output = SWR_output.iloc[:, [57,58]]
    coord_output = SWR_output.iloc[:,59:]
    SWR_output = SWR_output.iloc[:, 0:57]
    name_output = name_output.replace(np.nan, '', regex=True)
    
    #%% ***조작용 코드
    # SWR_avg_total = SWR_avg_total.drop(SWR_avg_total[(SWR_avg_total['DCR_DE_min'] > 0.6) | (SWR_avg_total['DCR_DE_max'] > 0.6)].index) # DE
    # SWR_avg_total = SWR_avg_total.drop(SWR_avg_total[(SWR_avg_total['DCR_MCE_min'] > 0.6) | (SWR_avg_total['DCR_MCE_max'] > 0.6)].index) # DE
    # SWR_avg_total = SWR_avg_total.drop(SWR_avg_total[(SWR_avg_total.iloc[:,4] < -0.0035) | (SWR_avg_total.iloc[:,3] > 0.0035)].index) # MCE
    
    #%% 엑셀로 출력(Using win32com)
    
    # Using win32com...
    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게
    
    wb = excel.Workbooks.Open(wall_design_xlsx_path)
    
    startrow, startcol = 5, 1
    
    # Table_S.Wall_DE 시트 입력
    ws3.Range('B%s:B%s' %(startrow, 5000)).ClearContents()
    ws3.Range('B%s:B%s' %(startrow, startrow + name_output.shape[0] - 1)).Value\
        = [[i] for i in name_output[0]]
    ws3.Range('A4:A4').Value\
        = len(wall_name_list)
    
    # Results_S.Wall_Rotation 시트 입력
    ws5.Range('A%s:BE%s' %(startrow, 5000)).ClearContents()
    ws5.Range('A%s:BE%s' %(startrow, startrow + SWR_output.shape[0] - 1)).Value\
        = list(SWR_output.itertuples(index=False, name=None))
        
    # Results_S.Wall_Rotation 시트 입력 (중력하중)
    ws5.Range('DV%s:DW%s' %(startrow, 5000)).ClearContents()
    ws5.Range('DV%s:DW%s' %(startrow, startrow + SWR_G_output.shape[0] - 1)).Value\
        = list(SWR_G_output.itertuples(index=False, name=None))
        
    # Results_S.Wall_Shear 시트 입력 (좌표)
    ws1.Range('GI%s:GL%s' %(startrow, 5000)).ClearContents()
    ws1.Range('GI%s:GL%s' %(startrow, startrow + SWR_output.shape[0] - 1)).Value\
        = list(coord_output.itertuples(index=False, name=None))
    
    wb.Save()
    # wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application 
        
    #%% 그래프
    if graph == True:
        # Wall 정보 load
        ws_DE = wb.Sheets('Table_S.Wall_DE')
        ws_MCE = wb.Sheets('Table_S.Wall_MCE')
        
        DE_result = ws_DE.Range('I%s:J%s' %(startrow, startrow + name_output.shape[0] - 1)).Value
        DE_result_arr = np.array(DE_result)[:,[0,1]]
        MCE_result = ws_MCE.Range('I%s:J%s' %(startrow, startrow + name_output.shape[0] - 1)).Value
        MCE_result_arr = np.array(MCE_result)[:,[0,1]]
        perform_lv = ws_DE.Range('K%s:M%s' %(startrow, startrow + name_output.shape[0] - 1)).Value
        perform_lv_arr = np.array(perform_lv)[:,[0,1,2]]
        
        # DCR 계산을 위해 결과값, Performance Level 합쳐서 Dataframe 생성
        WR_plot = np.concatenate((DE_result_arr, MCE_result_arr, perform_lv_arr), axis=1)
        WR_plot = pd.DataFrame(WR_plot)
        WR_plot.columns = ['DE_pos', 'DE_neg', 'MCE_pos', 'MCE_neg', 'IO', 'LS', 'CP']
        # DCR 계산
        WR_plot = WR_plot.apply(pd.to_numeric)
        WR_plot['DCR(DE_pos)'] = WR_plot['DE_pos'] / WR_plot['LS']
        WR_plot['DCR(DE_neg)'] = WR_plot['DE_neg'] / WR_plot['LS'] * (-1)
        WR_plot['DCR(MCE_pos)'] = WR_plot['MCE_pos'] / WR_plot['CP']
        WR_plot['DCR(MCE_neg)'] = WR_plot['MCE_neg'] / WR_plot['CP'] * (-1)
        
        WR_plot['Name'] = name_output.copy()
        
        #%% 벽체 해당하는 층 높이 할당
        story = []
        for i in WR_plot['Name']:
            if i == '':
                story.append(np.nan)
            else:
                story.append(i.split('_')[-1])        
        WR_plot['Story Name'] = story
        
        WR_plot = pd.merge(WR_plot, story_info.iloc[:,[1,2]], how='left')

        
        # 결과 dataframe -> pickle
        WR_result = []
        WR_result.append(WR_plot)
        WR_result.append(story_info)
        WR_result.append(DE_load_name_list)
        WR_result.append(MCE_load_name_list)
        with open('pkl/WR.pkl', 'wb') as f:
            pickle.dump(WR_result, f)
            
        count = 1

#%% Wall_SF
# 오류없는 또는 정확한 결과를 위해서는 MCE11, MCE12와 같이 짝이되는 지진파가 함께 있어야 함.

def WSF(self, input_xlsx_path, wall_design_xlsx_path, graph=True, DCR_criteria=1, yticks=2, xlim=3): 
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
#%% Load Data
    # Data Conversion Sheets
    story_info = self.story_info
    wall_info = self.wall_info
    rebar_info = self.rebar_info

    story_info.reset_index(inplace=True, drop=True)
    wall_info.reset_index(inplace=True, drop=True)

    # Analysis Result Sheets
    node_data = self.node_data
    element_data = self.wall_data
    wall_SF_data = self.shear_force_data

    # Seismic Loads List
    load_name_list = self.load_name_list
    gravity_load_name = self.gravity_load_name
    seismic_load_name_list = self.seismic_load_name_list
    DE_load_name_list = self.DE_load_name_list
    MCE_load_name_list = self.MCE_load_name_list

    # 필요없는 전단력 제거(층전단력)
    wall_SF_data = wall_SF_data[wall_SF_data['Name'].str.count('_') == 2] # underbar가 두개 들어간 행만 선택        
    wall_SF_data.reset_index(inplace=True, drop=True)

#%% 중력하중에 대한 전단력 데이터 grouping

    shear_force_G_data_grouped = pd.DataFrame()
    
    # G를 max, min으로 grouping
    for load_name in gravity_load_name:
        shear_force_G_data_grouped['G_H1_max'] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                 (wall_SF_data['Step Type'] == 'Max')]['H1(kN)'].values
            
        shear_force_G_data_grouped['G_H1_min'] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                 (wall_SF_data['Step Type'] == 'Min')]['H1(kN)'].values

        shear_force_G_data_grouped['G_H2_max'] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                 (wall_SF_data['Step Type'] == 'Max')]['H2(kN)'].values
            
        shear_force_G_data_grouped['G_H2_min'] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                 (wall_SF_data['Step Type'] == 'Min')]['H2(kN)'].values   

#%% DE, MCE에 대한 전단력 데이터 Grouping

    shear_force_DE_data_grouped = pd.DataFrame()
    shear_force_MCE_data_grouped = pd.DataFrame()

    ###
    full_DE_load_name_list = 'DE' + pd.Series([11,12,21,22,31,32,41,42,51,52,61,62,71,72]).astype(str)
    full_MCE_load_name_list = 'MCE' + pd.Series([11,12,21,22,31,32,41,42,51,52,61,62,71,72]).astype(str)
    
    # DE를 max, min으로 grouping
    for load_name in full_DE_load_name_list:
        if load_name in DE_load_name_list:
            shear_force_DE_data_grouped['{}_H1_max'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                          (wall_SF_data['Step Type'] == 'Max')]['H1(kN)'].values
                
            shear_force_DE_data_grouped['{}_H1_min'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                          (wall_SF_data['Step Type'] == 'Min')]['H1(kN)'].values
    
            shear_force_DE_data_grouped['{}_H2_max'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                          (wall_SF_data['Step Type'] == 'Max')]['H2(kN)'].values
                
            shear_force_DE_data_grouped['{}_H2_min'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                          (wall_SF_data['Step Type'] == 'Min')]['H2(kN)'].values   

        else:
            shear_force_DE_data_grouped['{}_H1_max'.format(load_name)] = ''
            shear_force_DE_data_grouped['{}_H1_min'.format(load_name)] = ''
            shear_force_DE_data_grouped['{}_H2_max'.format(load_name)] = ''
            shear_force_DE_data_grouped['{}_H2_min'.format(load_name)] = ''

    # MCE를 max, min으로 grouping
    for load_name in full_MCE_load_name_list:
        if load_name in MCE_load_name_list:
            shear_force_MCE_data_grouped['{}_H1_max'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                          (wall_SF_data['Step Type'] == 'Max')]['H1(kN)'].values
                
            shear_force_MCE_data_grouped['{}_H1_min'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                          (wall_SF_data['Step Type'] == 'Min')]['H1(kN)'].values
    
            shear_force_MCE_data_grouped['{}_H2_max'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                          (wall_SF_data['Step Type'] == 'Max')]['H2(kN)'].values
                
            shear_force_MCE_data_grouped['{}_H2_min'.format(load_name)] = wall_SF_data[(wall_SF_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                          (wall_SF_data['Step Type'] == 'Min')]['H2(kN)'].values   

        else:
            shear_force_MCE_data_grouped['{}_H1_max'.format(load_name)] = ''
            shear_force_MCE_data_grouped['{}_H1_min'.format(load_name)] = ''
            shear_force_MCE_data_grouped['{}_H2_max'.format(load_name)] = ''
            shear_force_MCE_data_grouped['{}_H2_min'.format(load_name)] = ''

#%% V(축력) 값 뽑기

    # 축력 불러와서 Grouping
    axial_force_data = wall_SF_data[wall_SF_data['Load Case'].str.contains(gravity_load_name[0])]['V(kN)']

    # result
    axial_force_data.reset_index(inplace=True, drop=True)
    axial_force = axial_force_data.groupby([[i//2 for i in range(0, len(axial_force_data))]], axis=0).min()
    
#%% 결과 정리 후 Input Sheets에 넣기

# 출력용 Dataframe 만들기
    # Results_S.Wall_Shear 시트
    SF_output = pd.DataFrame()
    SF_output['Name'] = wall_SF_data['Name'].drop_duplicates()
    SF_output.reset_index(inplace=True, drop=True)

    SF_output['Nu'] = axial_force
    SF_output = pd.concat([SF_output, shear_force_G_data_grouped
                           , shear_force_DE_data_grouped, shear_force_MCE_data_grouped], axis=1)
    
    # wall_info 순서에 맞게 sort
    SF_output = pd.merge(wall_info['Name'], SF_output, how='left')
    SF_output = SF_output.dropna(subset='Nu')
    
    # Design_S.Wall 시트
    steel_design_df = wall_info.iloc[:,[5,6,7,9,10]]
    wall_output = pd.concat([wall_info.iloc[:,0:11], steel_design_df], axis=1)
    
    # Table_S.Wall_DE 시트
    # Ground Level(0mm, 1F)에 가장 가까운 층의 row index get
    ground_level_idx = story_info['Height(mm)'].abs().idxmin()
    # story_info의 Index열을 1부터 시작하도록 재지정
    story_info['Index'] = range(story_info.shape[0], 0, -1)
    # Ground Level(0mm, 1F)에 가장 가까운 층을 index 5에 배정
    add_num_new_story = 5 - story_info.iloc[ground_level_idx, 0]
    story_info['Index'] = story_info['Index'] + add_num_new_story
    
    # Wall 이름 split    
    wall_info[['Wall Name', 'Wall Number', 'Story Name']] = wall_info['Name'].str.split('_', expand=True)
    wall_info = pd.merge(wall_info, story_info, how='left')
    # 결과값 없는 부재 제거
    SF_output = SF_output.replace(np.nan, '', regex=True)
    idx_to_slice = SF_output.iloc[:,6:].dropna().index # dropna로 결과값(DE,MCE) 있는 부재만 남긴 후 idx 추출
    idx_to_slice2 = wall_info['Name'].iloc[idx_to_slice].index # 결과값 있는 부재만 slice 후 idx 추출
    wall_info = wall_info.iloc[idx_to_slice2,:]
    wall_info.reset_index(inplace=True, drop=True)
    # 벽체 이름, 번호에 따라 grouping
    wall_name_list = list(wall_info.groupby(['Wall Name', 'Wall Number'], sort=False))
    # 55 row짜리 empty dataframe 만들기
    name_empty = pd.DataFrame(np.nan, index=range(55), columns=range(len(wall_name_list)))
    # dataframe에 이름 채워넣기
    count = 0
    while True:
        name_iter = wall_name_list[count][0][0]
        num_iter = wall_name_list[count][0][1]
        total_iter = wall_info['Name'][(wall_info['Wall Name'] == name_iter) 
                                       & (wall_info['Wall Number'] == num_iter)]
        idx_range = wall_info['Index'][(wall_info['Wall Name'] == name_iter) 
                                       & (wall_info['Wall Number'] == num_iter)]
        name_empty.iloc[idx_range, count] = total_iter
        
        count += 1
        if count == len(wall_name_list):
            break
    # dataframe을 1열로 만들기
    name_output_arr = np.array(name_empty)
    name_output_arr = np.reshape(name_output_arr, (-1, 1), order='F')
    name_output = pd.DataFrame(name_output_arr)
    
    # ETC 시트
    rebar_output = rebar_info.iloc[:,1:]
    
    # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
    SF_output = SF_output.replace(np.nan, '', regex=True)
    wall_output = wall_output.replace(np.nan, '', regex=True)
    name_output = name_output.replace(np.nan, '', regex=True)
    rebar_output = rebar_output.replace(np.nan, '', regex=True)
    
# 엑셀로 출력(Using win32com)
    
    # Using win32com...
    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게
    
    wb = excel.Workbooks.Open(wall_design_xlsx_path)
    ws1 = wb.Sheets('Results_S.Wall_Shear')
    ws2 = wb.Sheets('Design_S.Wall')
    ws3 = wb.Sheets('Table_S.Wall_DE')
    ws4 = wb.Sheets('ETC')
    
    startrow, startcol = 5, 1
    
    # Results_S.Wall_Shear 시트 입력
    ws1.Range('A%s:DN%s' %(startrow, 5000)).ClearContents()
    ws1.Range('A%s:DN%s' %(startrow, startrow + SF_output.shape[0] - 1)).Value\
        = list(SF_output.itertuples(index=False, name=None))
    
    # Design_S.Wall 시트 입력
    ws2.Range('A%s:P%s' %(startrow, 5000)).ClearContents()
    ws2.Range('A%s:P%s' %(startrow, startrow + wall_output.shape[0] - 1)).Value\
        = list(wall_output.itertuples(index=False, name=None))
    
    # Table_S.Wall_DE 시트 입력
    # 값을 입력하기 전에, 우선 해당 셀에 있는 값 지우기
    ws3.Range('B%s:B%s' %(startrow, 5000)).ClearContents()
    ws3.Range('B%s:B%s' %(startrow, startrow + name_output.shape[0] - 1)).Value\
        = [[i] for i in name_output[0]] # series -> list 형식만 입력가능
    ws3.Range('A4:A4').Value\
        = len(wall_name_list) # series -> list 형식만 입력가능 
        
    # Design_S.Wall 시트 입력
    ws4.Range('D%s:L%s' %(startrow, 5000)).ClearContents()
    ws4.Range('D%s:L%s' %(startrow, startrow + rebar_output.shape[0] - 1)).Value\
        = list(rebar_output.itertuples(index=False, name=None))
    
    wb.Save()
    # wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application 

#%% 그래프 process

    if graph == True:
        # Wall 정보 load
        ws_DE = wb.Sheets('Table_S.Wall_DE')
        ws_MCE = wb.Sheets('Table_S.Wall_MCE')
        
        DE_result = ws_DE.Range('S%s:S%s' %(startrow, startrow + name_output.shape[0] - 1)).Value
        DE_result_arr = np.array(DE_result)[:,0]
        MCE_result = ws_MCE.Range('S%s:S%s' %(startrow, startrow + name_output.shape[0] - 1)).Value
        MCE_result_arr = np.array(MCE_result)[:,0]
        
        wall_result = name_output.copy()
        wall_result['DE'] = DE_result_arr
        wall_result['MCE'] = MCE_result_arr
        wall_result.columns = ['Name', 'DE', 'MCE']

#%% ***조작용 코드
        # wall_name_to_delete = ['84A-W1_1','84A-W3_1_40F'] 
        # # 지우고싶은 층들을 대괄호 안에 입력(벽 이름만 입력하면 벽 전체 다 없어짐, 벽+층 이름 입력하면 특정 층의 벽만 없어짐)
        
        # for i in wall_name_to_delete:
        #     wall_result = wall_result[wall_result['Name'].str.contains(i) == False]
        
#%% 벽체 해당하는 층 높이 할당
        story = []
        for i in wall_result['Name']:
            if i == '':
                story.append(np.nan)
            else:
                story.append(i.split('_')[-1])        
        wall_result['Story Name'] = story
        
        wall_result = pd.merge(wall_result, story_info.iloc[:,[1,2]], how='left')
        
#%% Change non-numeric objects(e.g. str) into int or float as appropriate.
        wall_result['DE'] = pd.to_numeric(wall_result['DE'])
        wall_result['MCE'] = pd.to_numeric(wall_result['MCE'])
        # Delete rows with missing name or DCR over 1.0e+09
        wall_result = wall_result[wall_result['Name'] != '']
        wall_result = wall_result[wall_result['DE'].abs() < 1.0e+09]
        
#%% 그래프
        count = 1      
        
        # 결과 dataframe -> pickle
        WSF_result = []
        WSF_result.append(wall_result)
        WSF_result.append(story_info)
        WSF_result.append(DE_load_name_list)
        WSF_result.append(MCE_load_name_list)
        with open('pkl/WSF.pkl', 'wb') as f:
            pickle.dump(WSF_result, f)

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

#%% Redesign Horizontal Rebars

def WSF_redesign(wall_design_xlsx_path, rebar_limit=[None,None]): 
    ''' 

    완성된 <Results_Wall> 시트에서 전단보강이 필요한 부재들이 OK될 때까지 자동으로 배근함. \n
    
    세로 생성되는 <Results_Wall_보강> 시트에 보강 결과 출력 (철근 type 변경, 해결 안될 시 spacing은 10mm 간격으로 down)
    
    Parameters
    ----------
    input_path_xlsx : str
                      Data Conversion 엑셀 파일의 경로. (.xlsx)까지 기입해줘야한다. 
                      하나의 파일만 불러온다. \n

    rebar_limit : tuple, optional
                  (철근 type, spacing)의 형태로 입력. 
                  기본값은 ((사용된 수평철근 중) 최소지름,(사용된 수평철근 중) 최소간격)이다. \n

    Yields
    -------
    변경된 수평철근 정보를 <Results_Wall_보강> 시트에 엑셀로 출력한다. \n                               
    
    Raises
    -------
    
    References
    -------
    .. [1] "철근콘크리트 건축구조물의 성능기반 내진설계 지침", 대한건축학회, p.79, 2021
    
    '''
    ### Load Data and check/reassign the Rebar Limit
    # Win32com import
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게
    
    wb = excel.Workbooks.Open(wall_design_xlsx_path)
    ws1 = wb.Sheets('ETC')
    ws2 = wb.Sheets('Design_S.Wall')

    # Count the number of rows in each sheet to get the range of data
    ws1_row_num = ws1.UsedRange.Rows.Count # still ambiguous how the count works
    ws2_row_num = ws2.UsedRange.Rows.Count
    
    startrow, startcol = 5, 1
    
    # Load Data
    h_rebar_info = ws1.Range('A%s:A%s' %(startrow, ws1_row_num)).Value
    h_rebar_info_arr = np.array(h_rebar_info)[:,[0]]
    element_info = ws2.Range('A%s:P%s' %(startrow, ws2_row_num)).Value
    element_info_arr = np.array(element_info)[:,[0,14,15]]

    # Drop NoneType object
    h_rebar_info_arr = h_rebar_info_arr[h_rebar_info_arr != None]
    element_info_arr = element_info_arr[element_info_arr[:,0] != None]
    
    # Convert array to dataframe/Series
    element_info_df = pd.DataFrame(element_info_arr)
    element_info_df.columns = ['Name', 'H.Rebar Type', 'H.Rebar Spacing(mm)']
    
    # Detach 'D' from Rebar Type Column
    # h_rebar_info_list = [i.replace('D', '') for i in h_rebar_info_arr]
    # element_info_df['H.Rebar Type'] = element_info_df['H.Rebar Type'].str.replace('D', '')
    
    # Convert Data type from str to int
    # h_rebar_info_list = list(map(int, h_rebar_info_list))
    element_info_df['H.Rebar Spacing(mm)'] = element_info_df['H.Rebar Spacing(mm)'].astype(int)    

    ### Variables 설정
    # rebar_limit default 값 설정
    if rebar_limit[0] == None:
        rebar_sorted = element_info_df['H.Rebar Type'].sort_values(ascending=False)
        rebar_sorted.reset_index(inplace=True, drop=True)
        rebar_limit[0] = rebar_sorted.iloc[0]        
    if rebar_limit[1] == None:    
        rebar_limit[1] = element_info_df['H.Rebar Spacing(mm)'].min()
        
    # Loop 돌릴 철근 사이즈 리스트 설정
    rebar_limit_size_idx = np.where(h_rebar_info_arr == rebar_limit[0])[0][0] # rebar_limit의 index 구하기
    rebar_size_list = h_rebar_info_arr[:rebar_limit_size_idx+1] # index까지의 철근직경을 list로 만들기

    ### Print on Excel Sheets(Using win32com)        
    # NG인 부재의 Horizontal Rebar 간격 줄이기 (-10mm every iteration)
    while True:
        # 엑셀 읽기
        # H. Rebar 정보 읽기
        h_rebar_space = ws2.Range('P%s:P%s' %(startrow, startrow + element_info_df.shape[0] - 1)).Value
        h_rebar_space_array = np.array(h_rebar_space)[:,0] # list of tuples -> np.array
        # DCR 읽기
        dcr = ws2.Range('X%s:X%s' %(startrow, startrow + element_info_df.shape[0] - 1)).Value
        # DCR 값에 따른 np,array 생성 (NG가 있는 경우 = 1, NG가 없는 경우 = 0)
        dcr_list = []
        for row in dcr:
            if row[0] > 1.05:
                dcr_list.append(1)
            elif row[0] <= 1.05:
                dcr_list.append(0)
        dcr_array = np.array(dcr_list)

        # (NG) & (수평철근간격이 최소철근간격에 도달하지 않은) 부재들의 철근 간격 down
        h_rebar_space_array_updated = np.where(((dcr_array == 1) & (h_rebar_space_array - 10 >= rebar_limit[1]))
                                               , h_rebar_space_array - 10, h_rebar_space_array)

        # 수평철근간격 before & updated가 동일한 경우(철근간격이 update되지 않는 경우) break
        if np.array_equal(h_rebar_space_array, h_rebar_space_array_updated):
            break            

        # Horizontal Rebar 간격의 변경된 값을 Excel에 다시 입력
        ws2.Range('P%s:P%s' %(startrow, startrow + element_info_df.shape[0] - 1)).Value\
        = [[i] for i in h_rebar_space_array_updated]

        # Horizontal Diameter 직경/간격이 변경된(DCR == NG) 경우, 색 변경하기
        # h_rebar_space_diff_idx = np.where(h_rebar_space_array != h_rebar_space_array_updated)
        # for j in h_rebar_space_diff_idx[0]:
        #     ws_retrofit.Range('J%s' %str(startrow+int(j))).Font.ColorIndex = 3 # 3 : 빨간색

    # NG인 부재의 Horizontal Rebar Diameter 늘리기 (<ETC>시트의 철근직경 순서에 따라)
    while True:
        # 엑셀 읽기
        # H. Rebar 정보 읽기
        h_rebar_type = ws2.Range('O%s:O%s' %(startrow, startrow + element_info_df.shape[0] - 1)).Value # list of tuples
        h_rebar_type_df = pd.DataFrame(h_rebar_type) # list of tupels -> dataframe        
        h_rebar_info_idx = pd.Index(h_rebar_info_arr) # <ETC> 시트의 철근직경 순서를 list of index로 만들기
        # 철근직경 순서 index를 매칭시켜, 각 부재의 철근직경에 대한 index list 만들기
        h_rebar_type_idx = h_rebar_info_idx.get_indexer(h_rebar_type_df.iloc[:,0]) 
                                          # get_indexer(list) : list의 값들의 h_rebar_idx에서의 인덱스찾기
        # DCR 읽기
        dcr = ws2.Range('X%s:X%s' %(startrow, startrow + element_info_df.shape[0] - 1)).Value
        # DCR 값에 따른 np,array 생성 (NG가 있는 경우 = 1, NG가 없는 경우 = 0)
        dcr_list = []
        for row in dcr:
            if row[0] > 1.05:
                dcr_list.append(1)
            elif row[0] <= 1.05:
                dcr_list.append(0)
        dcr_array = np.array(dcr_list)
        
        # (NG) & (수평철근직경이 최대철근직경에 도달하지 않은) 부재들의 철근 직경 up
        h_rebar_type_idx_updated = np.where(((dcr_array == 1) 
                                              & (h_rebar_type_idx + 1 <= h_rebar_info_idx.get_loc(rebar_limit[0])))
                                              , h_rebar_type_idx + 1, h_rebar_type_idx)  

        # 수평철근직경 before & updated가 동일한 경우(철근직경이 update되지 않는 경우) break
        if np.array_equal(h_rebar_type_idx, h_rebar_type_idx_updated):
            break 

        # Horizontal Rebar Diameter의 변경된 값을 Excel에 다시 입력
        ws2.Range('O%s:O%s' %(startrow, startrow + element_info_df.shape[0] - 1)).Value\
        = [[h_rebar_info_arr[i]] for i in h_rebar_type_idx_updated]
        
        # Horizontal Diameter 직경/간격이 변경된(DCR == NG) 경우, 색 변경하기
        # h_rebar_type_diff_idx = np.where(h_rebar_type_idx != h_rebar_type_idx_updated)
        # for j in h_rebar_type_diff_idx[0]:
        #     ws_retrofit.Range('I%s' %str(startrow+int(j))).Font.ColorIndex = 3 # 3 : 빨간색

    #
    wb.Save()            
    # wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application

#%% Wall Axial Strain (Preview)

def WAS_plot(self, wall_design_xlsx_path) -> pd.DataFrame:
    '''
    Parameters
    ----------
    wall_design_xlsx_path : str
        File path of "Seismic Design_Shear Wall" EXCEL file

    Returns
    -------
    WAS.pkl : pickle
        Wall Axial Strain results in pd.DataFrame type is saved as pickle in WAS.pkl

    '''

    ### Load Data
    # Data Conversion Sheets
    story_info = self.story_info
    
    # Seismic Loads List
    load_name_list = self.load_name_list
    gravity_load_name = self.gravity_load_name
    seismic_load_name_list = self.seismic_load_name_list
    DE_load_name_list = self.DE_load_name_list
    MCE_load_name_list = self.MCE_load_name_list
    
    ##### Excel 파일 읽는 Function (w/ Xlsx2csv & joblib)
    def read_excel(path:str, sheet_name:str, skip_rows:list=[0,2,3]) -> pd.DataFrame:
        import pandas as pd
        from io import StringIO # if not import, error occurs when using multiprocessing
        from xlsx2csv import Xlsx2csv
        data_buffer = StringIO()
        Xlsx2csv(path, outputencoding="utf-8", ignore_formats='float').convert(data_buffer, sheetname=sheet_name)
        data_buffer.seek(0)
        data_df = pd.read_csv(data_buffer, low_memory=False, skiprows=skip_rows)
        return data_df
    
    ### Read Excel Files (Data Conversion Sheets & Analysis Result Sheets)
    # DE result & name_output
    DE_result = read_excel(wall_design_xlsx_path, sheet_name='Table_S.Wall_DE')
    name_output = pd.DataFrame(DE_result.iloc[:,1])
    name_output.dropna(how='all', inplace=True)
    DE_result = DE_result.iloc[:,[20,21]]
    DE_result.dropna(how='all', inplace=True)
    DE_result_arr = np.array(DE_result)
    # MCE result
    MCE_result = read_excel(wall_design_xlsx_path, sheet_name='Table_S.Wall_MCE')
    MCE_result = MCE_result.iloc[:,[20,21]]
    MCE_result.dropna(how='all', inplace=True)
    MCE_result_arr = np.array(MCE_result)

    ### Create Final Dataframe to Export
    WAS_plot = name_output.copy()
    WAS_plot[['DE(Compressive)', 'DE(Tensile)']] = DE_result_arr
    WAS_plot[['MCE(Compressive)', 'MCE(Tensile)']] = MCE_result_arr
    WAS_plot.columns = ['Name', 'DE(Compressive)', 'DE(Tensile)', 'MCE(Compressive)', 'MCE(Tensile)']
    
    # 벽체 해당하는 층 높이 할당
    story = []
    for i in WAS_plot['Name']:
        if i == '':
            story.append(np.nan)
        else:
            story.append(i.split('_')[-1])        
    WAS_plot['Story Name'] = story
    
    WAS_plot = pd.merge(WAS_plot, story_info.iloc[:,[1,2]], how='left')
    
    # Change non-numeric objects(e.g. str) into int or float as appropriate.
    WAS_plot['DE(Compressive)'] = pd.to_numeric(WAS_plot['DE(Compressive)'])
    WAS_plot['DE(Tensile)'] = pd.to_numeric(WAS_plot['DE(Tensile)'])
    WAS_plot['MCE(Compressive)'] = pd.to_numeric(WAS_plot['MCE(Compressive)'])
    WAS_plot['MCE(Tensile)'] = pd.to_numeric(WAS_plot['MCE(Tensile)'])
    
    # Delete rows with missing name
    WAS_plot = WAS_plot[WAS_plot['Name'] != '']
            
    # 결과 dataframe -> pickle
    WAS_result = []
    WAS_result.append(WAS_plot)
    WAS_result.append(story_info)
    WAS_result.append(DE_load_name_list)
    WAS_result.append(MCE_load_name_list)
    with open('pkl/WAS.pkl', 'wb') as f:
        pickle.dump(WAS_result, f)

#%% Wall Rotation (Preview)

def WR_plot(self, wall_design_xlsx_path) -> pd.DataFrame:
    '''
    Parameters
    ----------
    wall_design_xlsx_path : str
        File path of "Seismic Design_Shear Wall" EXCEL file

    Returns
    -------
    WR.pkl : pickle
        Wall Rotation results in pd.DataFrame type is saved as pickle in WR.pkl

    '''
    
    ### Load Data
    # Data Conversion Sheets
    story_info = self.story_info

    story_info.reset_index(inplace=True, drop=True)

    # Seismic Loads List
    load_name_list = self.load_name_list
    gravity_load_name = self.gravity_load_name
    seismic_load_name_list = self.seismic_load_name_list
    DE_load_name_list = self.DE_load_name_list
    MCE_load_name_list = self.MCE_load_name_list
        
    ##### Excel 파일 읽는 Function (w/ Xlsx2csv & joblib)
    def read_excel(path:str, sheet_name:str, skip_rows:list=[0,2,3]) -> pd.DataFrame:
        import pandas as pd
        from io import StringIO # if not import, error occurs when using multiprocessing
        from xlsx2csv import Xlsx2csv
        data_buffer = StringIO()
        Xlsx2csv(path, outputencoding="utf-8", ignore_formats='float').convert(data_buffer, sheetname=sheet_name)
        data_buffer.seek(0)
        data_df = pd.read_csv(data_buffer, low_memory=False, skiprows=skip_rows)
        return data_df
    
    ### Read Excel Files (Data Conversion Sheets & Analysis Result Sheets)
    # DE result & name_output
    DE_result = read_excel(wall_design_xlsx_path, sheet_name='Table_S.Wall_DE')
    name_output = pd.DataFrame(DE_result.iloc[:,1])
    name_output.dropna(how='all', inplace=True)
    name_output.reset_index(inplace=True, drop=True)
    DE_result = DE_result.iloc[:,[8,9]]
    DE_result.dropna(how='all', inplace=True)
    DE_result_arr = np.array(DE_result)
    # MCE result
    MCE_result = read_excel(wall_design_xlsx_path, sheet_name='Table_S.Wall_MCE')
    MCE_result = MCE_result.iloc[:,[8,9]]
    MCE_result.dropna(how='all', inplace=True)
    MCE_result_arr = np.array(MCE_result)
    # Performance Criteria
    perform_lv = read_excel(wall_design_xlsx_path, sheet_name='Table_S.Wall_DE')
    perform_lv = perform_lv.iloc[:,[10,11,12]]
    perform_lv.dropna(how='all', inplace=True)
    perform_lv_arr = np.array(perform_lv)
    
    # DCR 계산을 위해 결과값, Performance Level 합쳐서 Dataframe 생성
    WR_plot = np.concatenate((DE_result_arr, MCE_result_arr, perform_lv_arr), axis=1)
    WR_plot = pd.DataFrame(WR_plot)
    WR_plot.columns = ['DE_pos', 'DE_neg', 'MCE_pos', 'MCE_neg', 'IO', 'LS', 'CP']
    # DCR 계산
    WR_plot = WR_plot.apply(pd.to_numeric)
    WR_plot['DCR(DE_pos)'] = WR_plot['DE_pos'] / WR_plot['LS']
    WR_plot['DCR(DE_neg)'] = WR_plot['DE_neg'] / WR_plot['LS'] * (-1)
    WR_plot['DCR(MCE_pos)'] = WR_plot['MCE_pos'] / WR_plot['CP']
    WR_plot['DCR(MCE_neg)'] = WR_plot['MCE_neg'] / WR_plot['CP'] * (-1)
    
    WR_plot['Name'] = name_output.copy()
    
    # 벽체 해당하는 층 높이 할당
    story = []
    for i in WR_plot['Name']:
        if i == '':
            story.append(np.nan)
        else:
            story.append(i.split('_')[-1])        
    WR_plot['Story Name'] = story
    
    WR_plot = pd.merge(WR_plot, story_info.iloc[:,[1,2]], how='left')

    
    # 결과 dataframe -> pickle
    WR_result = []
    WR_result.append(WR_plot)
    WR_result.append(story_info)
    WR_result.append(DE_load_name_list)
    WR_result.append(MCE_load_name_list)
    with open('pkl/WR.pkl', 'wb') as f:
        pickle.dump(WR_result, f)
        
#%% Wall Shear Force (Preview)
def WSF_plot(self, wall_design_xlsx_path) -> pd.DataFrame:
    '''
    Parameters
    ----------
    wall_design_xlsx_path : str
        File path of "Seismic Design_Shear Wall" EXCEL file

    Returns
    -------
    WSF.pkl : pickle
        Wall Shear Force results in pd.DataFrame type is saved as pickle in WSF.pkl

    '''
    
    ### Load Data
    # Data Conversion Sheets
    story_info = self.story_info

    story_info.reset_index(inplace=True, drop=True)

    # Seismic Loads List
    load_name_list = self.load_name_list
    gravity_load_name = self.gravity_load_name
    seismic_load_name_list = self.seismic_load_name_list
    DE_load_name_list = self.DE_load_name_list
    MCE_load_name_list = self.MCE_load_name_list
    
    ##### Excel 파일 읽는 Function (w/ Xlsx2csv & joblib)
    def read_excel(path:str, sheet_name:str, skip_rows:list=[0,2,3]) -> pd.DataFrame:
        import pandas as pd
        from io import StringIO # if not import, error occurs when using multiprocessing
        from xlsx2csv import Xlsx2csv
        data_buffer = StringIO()
        Xlsx2csv(path, outputencoding="utf-8", ignore_formats='float').convert(data_buffer, sheetname=sheet_name)
        data_buffer.seek(0)
        data_df = pd.read_csv(data_buffer, low_memory=False, skiprows=skip_rows)
        return data_df

    ### Read Excel Files (Data Conversion Sheets & Analysis Result Sheets)
    # DE result & name_output
    DE_result = read_excel(wall_design_xlsx_path, sheet_name='Table_S.Wall_DE')
    name_output = pd.DataFrame(DE_result.iloc[:,1])
    name_output.dropna(how='all', inplace=True)
    DE_result = DE_result.iloc[:,18]
    DE_result.dropna(how='all', inplace=True)
    DE_result_arr = np.array(DE_result)
    # MCE result
    MCE_result = read_excel(wall_design_xlsx_path, sheet_name='Table_S.Wall_MCE')
    MCE_result = MCE_result.iloc[:,18]
    MCE_result.dropna(how='all', inplace=True)
    MCE_result_arr = np.array(MCE_result)
    
    wall_result = name_output.copy()
    wall_result['DE'] = DE_result_arr
    wall_result['MCE'] = MCE_result_arr
    wall_result.columns = ['Name', 'DE', 'MCE']
    
    # 벽체 해당하는 층 높이 할당
    story = []
    for i in wall_result['Name']:
        if i == '':
            story.append(np.nan)
        else:
            story.append(i.split('_')[-1])        
    wall_result['Story Name'] = story
    
    wall_result = pd.merge(wall_result, story_info.iloc[:,[1,2]], how='left')
    
    # Change non-numeric objects(e.g. str) into int or float as appropriate.
    wall_result['DE'] = pd.to_numeric(wall_result['DE'])
    wall_result['MCE'] = pd.to_numeric(wall_result['MCE'])
    # Delete rows with missing name or DCR over 1.0e+09
    wall_result = wall_result[wall_result['Name'] != '']
    wall_result = wall_result[wall_result['DE'].abs() < 1.0e+09]
    
    # 결과 dataframe -> pickle
    WSF_result = []
    WSF_result.append(wall_result)
    WSF_result.append(story_info)
    WSF_result.append(DE_load_name_list)
    WSF_result.append(MCE_load_name_list)
    with open('pkl/WSF.pkl', 'wb') as f:
        pickle.dump(WSF_result, f)