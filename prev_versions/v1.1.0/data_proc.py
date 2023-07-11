import os
import pandas as pd
import numpy as np
import win32com.client
import pythoncom
import re
import warnings
from collections import deque

#%% Node, Element, Mass, Load Import

def import_midas(input_xlsx_path, DL_name='DL', LL_name='LL'\
                 , import_node=True, import_DL=True, import_LL=True\
                 , import_mass=True, **kwargs):
    
    #%% 변수 정리(default값=True)
    import_beam = kwargs['import_beam'] if 'import_beam' in kwargs.keys() else True
    import_column = kwargs['import_column'] if 'import_column' in kwargs.keys() else False
    import_wall = kwargs['import_wall'] if 'import_wall' in kwargs.keys() else True
    import_plate = kwargs['import_plate'] if 'import_plate' in kwargs.keys() else False
    import_WR_gage = kwargs['import_WR_gage'] if 'import_WR_gage' in kwargs.keys() else True
    import_WAS_gage = kwargs['import_WAS_gage'] if 'import_WAS_gage' in kwargs.keys() else True
    import_I_beam = kwargs['import_I_beam'] if 'import_I_beam' in kwargs.keys() else True

    '''    
    Midas GEN 모델을 Perform-3D로 import할 수 있는 파일 형식(.csv)으로 변환.
    
    Parameters
    ----------
    input_path : str
                 Data Conversion 엑셀 파일의 경로.
                 
    input_xlsx : str
                 Data Conversion 엑셀 파일의 이름. 확장자명(.xlsx)까지 
                 기입해줘야한다. 하나의 파일만 불러온다.
    
    DL_name : str
              Perform-3D에서 생성한 고정하중(Dead Load)의 이름을 입력함.
              
    LL_name : str
              Perform-3D에서 생성한 활하중(Live Load)의 이름을 입력함.

    import_node : bool, optional, default=True
                  True = node의 csv파일을 생성함.
                  False = node의 csv파일을 생성 안 함.

    import_DL : bool, optional, default=True
                True = 고정하중의 csv파일을 생성함.
                False = 고정하중의 csv파일을 생성 안 함.

    import_LL : bool, optional, default=True
                True = 활하중의 csv파일을 생성함.
                False = 활하중의 csv파일을 생성 안 함.
            
    import_mass : bool, optional, default=True
                  True = mass의 csv파일을 생성함.
                  False = mass의 csv파일을 생성 안 함.

    Returns
    -------        
    csv files or None
    Data Conversion 엑셀파일과 같은 경로에 csv 파일들을 생성함.
    
    Other Parameters
    ----------------
    import_beam : bool, optional, default=True
                  True = beam의 csv파일을 생성함.
                  False = beam의 csv파일을 생성 안 함.
                  
    import_column : bool, optional, default=True
                    True = column의 csv파일을 생성함.
                    False = column의 csv파일을 생성 안 함.

    import_wall : bool, optional, default=True
                  True = wall의 csv파일을 생성함.
                  False = wall의 csv파일을 생성 안 함.
                  
    import_plate : bool, optional, default=True
                   True = plate의 csv파일을 생성함.
                   False = plate의 csv파일을 생성 안 함.

    import_WR_gage : bool, optional, default=True
                      True = Wall Rotation Gage의 csv파일을 생성함.
                      False = Wall Rotation Gage의 csv파일을 생성 안 함.
                   
    import_WAS_gage : bool, optional, default=True
                     True = Wall Axial Strain Gage의 csv파일을 생성함.
                     False = Wall Axial Strain Gage의 csv파일을 생성 안 함.

    import_I_beam : bool, optional, default=True
                     True = Imbedded Beam의 csv파일을 생성함.
                     False = Imbedded Beam의 csv파일을 생성 안 함.                
    Raises
    -------
    
    '''
    
    #%% 변수, 이름 지정
    
    DL_name = [DL_name] # DL에 포함시킬 하중이름 포함("DL_XX"와 같은 형태의 하중들만 있을 경우, "DL"만 넣어주면 됨)
    LL_name = [LL_name]
    
    input_xlsx_sheet = 'Nodes'
    nodal_load_raw_xlsx_sheet = 'Nodal Loads'
    mass_raw_xlsx_sheet = 'Story Mass'
    element_raw_xlsx_sheet = 'Elements'
    story_info_xlsx_sheet = 'Story Data'
    
    # Output 경로 설정
    output_csv_dir = os.path.dirname(input_xlsx_path) # 또는 '경로'
    
    node_DL_merged_csv = 'DL.csv'
    node_LL_merged_csv = 'LL.csv'
    mass_csv = 'Mass.csv'
    node_csv = 'Node.csv'
    mass_node_csv = 'Node(Mass).csv'
    beam_csv = 'Beam.csv'
    column_csv = 'Column.csv'
    wall_csv = 'Wall.csv'
    plate_csv = 'Plate.csv'
    WR_gage_csv = 'Shear Wall Rotation Gage.csv'
    WAS_gage_csv = 'Axial Strain Gage.csv'
    I_beam_csv = 'Imbedded Beam.csv'
    
    #%% Nodal Load 뽑기
    
    # Node 정보 load
    node = pd.read_excel(input_xlsx_path, sheet_name = input_xlsx_sheet
                         , skiprows = 3, usecols=[0,1,2,3], index_col = 0)  # Node 열을 인덱스로 지정
    node.columns = ['X(mm)', 'Y(mm)', 'Z(mm)']
    
    if (import_DL == True) or (import_LL == True):
    
        # Nodal Load 정보 load
        nodal_load = pd.read_excel(input_xlsx_path, sheet_name = nodal_load_raw_xlsx_sheet
                                   , skiprows = 3, usecols=[0,1,2,3,4,5,6,7], index_col = 0)
        nodal_load.columns = ['Loadcase', 'FX(kN)', 'FY(kN)', 'FZ(kN)', 'MX(kN-mm)', 'MY(kN-mm)', 'MZ(kN-mm)']
        
        # Nodal Load를 DL과 LL로 분리
        DL = []
        LL = []
        
        for i in DL_name:
            DL_temp = nodal_load.loc[lambda x: nodal_load['Loadcase'].str.contains(i), :]  # lambda로 만든 함수로 Loadcase가 i인 행만 slicing
            DL.append(DL_temp)
            
        for i in LL_name:
            LL_temp = nodal_load.loc[lambda x: nodal_load['Loadcase'].str.contains(i), :]  # lambda로 만든 함수로 Loadcase가 i인 행만 slicing
            LL.append(LL_temp)
            
        DL = pd.concat(DL)
        LL = pd.concat(LL)
        
        DL2 = DL.drop('Loadcase', axis=1)  # axis=1(열), axis=0(행)
        LL2 = LL.drop('Loadcase', axis=1)  # 필요없어진 Loadcase 열은 drop으로 떨굼
        
        # Node와 Nodal Load를 element number 기준으로 병합
        node_DL_merged = pd.merge(node, DL2, left_index=True, right_index=True)  # node 좌표와 하중을 결합하여 dataframe 만들기, merge : 공통된 index를 기준으로 합침
        node_LL_merged = pd.merge(node, LL2, left_index=True, right_index=True)  # left_index, right_index는 뭔지 기억은 안나는데 오류고치기위해서 더함
        
        # DL, LL 결과값을 csv로 변환
        if import_DL == True:
            node_DL_merged.to_csv(output_csv_dir+'\\'+node_DL_merged_csv, mode='w', index=False)  # to_csv 사용. index=False로 index 열은 떨굼
        
        if import_LL == True:
            node_LL_merged.to_csv(output_csv_dir+'\\'+node_LL_merged_csv, mode='w', index=False)
    
    #%% Mass, Node 뽑기
    if import_mass == True:
        
        # Mass 정보 load
        mass = pd.read_excel(input_xlsx_path, sheet_name = mass_raw_xlsx_sheet
                             , skiprows = 3, usecols=[0,1,2,3,4,5,6,7,8,9,10])
        mass.columns = ['Story', 'Z(mm)', 'Trans Mass X-dir(kN/g)', 'Trans Mass Y-dir(kN/g)', 'Rotat Mass(kN/g-mm^2)',\
                        'X(mm)_Mass', 'Y(mm)_Mass', 'X(mm)_Stiffness', 'Y(mm)_Stiffness', 'X(mm)', 'Y(mm)']
        
        # Mass가 0인 층 제거
        mass = mass[(mass['Trans Mass X-dir(kN/g)'] != 0) & (mass['Rotat Mass(kN/g-mm^2)'] != 0)]
        mass.reset_index(inplace=True, drop=True) 
        
        # 필요없는 열 제거
        mass2 = mass.drop('Story', axis=1)
        
        # 열 재배치
        mass2 = mass2[['X(mm)', 'Y(mm)', 'Z(mm)', 'Trans Mass X-dir(kN/g)', 'Trans Mass Y-dir(kN/g)', 'Rotat Mass(kN/g-mm^2)']]
        
        # 형태 맞추기 위해 열 추가
        mass2.insert(5, 'Trans Mass Z-dir(kN/g)', 0)  # insert로 5번째 열의 위치에 column 삽입
        mass2.insert(6, 'Rotat Mass X-dir(kN/g mm^2)', 0)
        mass2.insert(7, 'Rotat Mass Y-dir(kN/g mm^2)', 0)
        
        # Mass 결과값을 csv로 변환
        mass2.to_csv(output_csv_dir+'\\'+mass_csv, mode='w', index=False)
    
        # Node 결과값을 csv로 변환
        if import_node == True:
            # Mass의 nodes(좌표) 추가        
            node_mass_considered = pd.concat([node, mass2.iloc[:,[0,1,2]]])
            node_mass_considered.to_csv(output_csv_dir+'\\'+node_csv, mode='w', index=False) # Import할 Mass의 좌표를 포함한 모든 좌표를 csv로 출력함
       
        else:
            mass2.iloc[:,[0,1,2]].to_csv(output_csv_dir+'\\'+mass_node_csv, mode='w', index=False) # Import할 Mass의 좌표만 csv로 출력함
            
    else:
        if import_node == True:
            # Node 결과값을 csv로 변환
            node.to_csv(output_csv_dir+'\\'+node_csv, mode='w', index=False)
        
    #%% Beam Element 뽑기
    
    # Index로 지정되어있던 Node 번호를 다시 reset
    node.index.name = 'Node'
    node.reset_index(inplace=True)
    
    # Element 정보 load
    element = pd.read_excel(input_xlsx_path, sheet_name = element_raw_xlsx_sheet
                            , skiprows = [0,2,3], usecols=[0,1,2,3,4,5,6,7,8,9,10,11])
    
    # Beam Element만 추출(slicing)
    if (import_beam == True) or (import_column == True):
        
        frame = element.loc[lambda x: element['Type'] == 'BEAM', :]
        
        # 필요한 열만 추출(drop하기에는 drop할 열이 너무 많아서...)
        frame_node_1 = frame.loc[:, 'Node1']
        frame_node_2 = frame.loc[:, 'Node2']
        
        frame_node_1.name = 'Node'  # Merge(같은 열을 기준으로 두 dataframe 결합)를 사용하기 위해 index를 Node로 바꾸기
        frame_node_2.name = 'Node'
        
        # Merge로 Node 번호에 맞는 좌표를 결합
        frame_node_1_coord = pd.merge(frame_node_1, node, how='left', on='Node')  # how='left' : 두 데이터프레임 중 왼쪽 데이터프레임은 그냥 두고 오른쪽 데이터프레임값을 대응시킴
        frame_node_2_coord = pd.merge(frame_node_2, node, how='left', on='Node')
        
        # Node1, Node2의 좌표를 모두 결합시켜 출력
        frame_node_1_coord = frame_node_1_coord.drop('Node', axis=1)
        frame_node_2_coord = frame_node_2_coord.drop('Node', axis=1)
        
        frame_node_1_coord.columns = ['X_1(mm)', 'Y_1(mm)', 'Z_1(mm)']  # 결합 때 이름이 중복되면 안되서 이름 바꿔줌
        frame_node_2_coord.columns = ['X_2(mm)', 'Y_2(mm)', 'Z_2(mm)']
        
        frame_node_coord = pd.concat([frame_node_1_coord, frame_node_2_coord], axis=1)
        
        # Column, Beam 나누기
        column_node_coord = frame_node_coord[abs(frame_node_coord['Z_1(mm)'] - frame_node_coord['Z_2(mm)']) > 10]
        beam_node_coord = frame_node_coord[abs(frame_node_coord['Z_1(mm)'] - frame_node_coord['Z_2(mm)']) <= 10]
    
        # 부재의 orientation 맞춘 후 csv로 출력
        if import_column == True:
            # Z 좌표가 더 작은 노드를 i-node로!
            column_node_coord_pos = column_node_coord[column_node_coord['Z_2(mm)'] >= column_node_coord['Z_1(mm)']]
            column_node_coord_neg = column_node_coord[column_node_coord['Z_2(mm)'] < column_node_coord['Z_1(mm)']]
            
            column_node_coord_neg = column_node_coord_neg.iloc[:,[3,4,5,0,1,2]]
            column_node_coord_neg.columns = ['X_1(mm)', 'Y_1(mm)', 'Z_1(mm)', 'X_2(mm)', 'Y_2(mm)', 'Z_2(mm)']
            
            # pos, neg 합치기
            column_node_coord = pd.concat([column_node_coord_pos, column_node_coord_neg]\
                                          , ignore_index=True)
            
            # 출력
            column_node_coord.to_csv(output_csv_dir+'\\'+column_csv, mode='w', index=False)
        
        if import_beam == True:
            # X 좌표가 더 작은 노드를 i-node로!
            beam_node_coord_pos = beam_node_coord[beam_node_coord['X_2(mm)'] > beam_node_coord['X_1(mm)']]
            beam_node_coord_neg = beam_node_coord[beam_node_coord['X_2(mm)'] < beam_node_coord['X_1(mm)']]
            beam_node_coord_zero = beam_node_coord[beam_node_coord['X_2(mm)'] == beam_node_coord['X_1(mm)']]
            
            beam_node_coord_neg = beam_node_coord_neg.iloc[:,[3,4,5,0,1,2]]
            beam_node_coord_neg.columns = beam_node_coord_pos.columns.values
            
            # Y 좌표가 더 작은 노드를 i-node로!
            beam_node_coord_zero_pos = beam_node_coord_zero[beam_node_coord_zero['Y_2(mm)'] >= beam_node_coord_zero['Y_1(mm)']]
            beam_node_coord_zero_neg = beam_node_coord_zero[beam_node_coord_zero['Y_2(mm)'] < beam_node_coord_zero['Y_1(mm)']]
            
            beam_node_coord_zero_neg = beam_node_coord_zero_neg.iloc[:,[3,4,5,0,1,2]]
            beam_node_coord_zero_neg.columns = beam_node_coord_zero_pos.columns.values
            
            # pos, neg 합치기
            beam_node_coord = pd.concat([beam_node_coord_pos, beam_node_coord_neg\
                                         , beam_node_coord_zero_pos, beam_node_coord_zero_neg]\
                                        , ignore_index=True)
            
            # 출력
            beam_node_coord.to_csv(output_csv_dir+'\\'+beam_csv, mode='w', index=False)
    
    #%% Wall Element 뽑기
    if import_wall == True:
        
        # Wall Element만 추출(slicing)
        wall = element.loc[lambda x: element['Type'] == 'WALL', :]
        
        # 필요한 열만 추출
        wall_node_1 = wall.loc[:, 'Node1']
        wall_node_2 = wall.loc[:, 'Node2']
        wall_node_3 = wall.loc[:, 'Node3']
        wall_node_4 = wall.loc[:, 'Node4']
        
        wall_node_1.name = 'Node'
        wall_node_2.name = 'Node'
        wall_node_3.name = 'Node'
        wall_node_4.name = 'Node'
        
        # Merge로 Node 번호에 맞는 좌표를 결합
        wall_node_1_coord = pd.merge(wall_node_1, node, how='left')
        wall_node_2_coord = pd.merge(wall_node_2, node, how='left')
        wall_node_3_coord = pd.merge(wall_node_3, node, how='left')
        wall_node_4_coord = pd.merge(wall_node_4, node, how='left')
        
        # Node1, Node2, Node3, Node4의 좌표를 모두 결합시켜 출력
        wall_node_1_coord = wall_node_1_coord.drop('Node', axis=1)
        wall_node_2_coord = wall_node_2_coord.drop('Node', axis=1)
        wall_node_3_coord = wall_node_3_coord.drop('Node', axis=1)
        wall_node_4_coord = wall_node_4_coord.drop('Node', axis=1)
        
        wall_node_1_coord.columns = ['X_1(mm)', 'Y_1(mm)', 'Z_1(mm)']
        wall_node_2_coord.columns = ['X_2(mm)', 'Y_2(mm)', 'Z_2(mm)']
        wall_node_3_coord.columns = ['X_3(mm)', 'Y_3(mm)', 'Z_3(mm)']
        wall_node_4_coord.columns = ['X_4(mm)', 'Y_4(mm)', 'Z_4(mm)']
        
        wall_node_coord = pd.concat([wall_node_1_coord, wall_node_2_coord, wall_node_3_coord, wall_node_4_coord], axis=1)
                
        ### 부재의 orientation 맞춘 후 csv로 출력
        # X 좌표가 더 작은 노드를 i-node로!
        # 허용 오차
        tolerance = 5 # mm
        wall_node_coord_pos = wall_node_coord[wall_node_coord['X_2(mm)'] - wall_node_coord['X_1(mm)'] > tolerance]
        wall_node_coord_neg = wall_node_coord[wall_node_coord['X_2(mm)'] - wall_node_coord['X_1(mm)'] < -tolerance]
        wall_node_coord_zero = wall_node_coord[(wall_node_coord['X_2(mm)'] - wall_node_coord['X_1(mm)'] >= -tolerance) 
                                               & (wall_node_coord['X_2(mm)'] - wall_node_coord['X_1(mm)'] <= tolerance)]
        
        wall_node_coord_neg = wall_node_coord_neg.iloc[:,[3,4,5,0,1,2,9,10,11,6,7,8]]
        wall_node_coord_neg.columns = wall_node_coord_pos.columns.values
        
        # Y 좌표가 더 작은 노드를 i-node로!
        wall_node_coord_zero_pos = wall_node_coord_zero[wall_node_coord_zero['Y_2(mm)'] >= wall_node_coord_zero['Y_1(mm)']]
        wall_node_coord_zero_neg = wall_node_coord_zero[wall_node_coord_zero['Y_2(mm)'] < wall_node_coord_zero['Y_1(mm)']]
        
        wall_node_coord_zero_neg = wall_node_coord_zero_neg.iloc[:,[3,4,5,0,1,2,9,10,11,6,7,8]]
        wall_node_coord_zero_neg.columns = wall_node_coord_zero_pos.columns.values
        
        # pos, neg 합치기
        wall_node_coord = pd.concat([wall_node_coord_pos, wall_node_coord_neg\
                                     , wall_node_coord_zero_pos, wall_node_coord_zero_neg]\
                                    , ignore_index=True)
        
        # Wall Element 결과값을 csv로 변환
        wall_node_coord.to_csv(output_csv_dir+'\\'+wall_csv, mode='w', index=False) 
    
    #%% Axial Strain Gage 뽑기
    
    if import_WAS_gage == True:
        
        # Wall Element만 추출(slicing)
        wall = element.loc[lambda x: element['Type'] == 'WALL', :]
        
        wall_gage = wall.loc[:,['Wall ID', 'Node1', 'Node2', 'Node3', 'Node4']]
        
        # Merge로 Node 번호에 맞는 좌표를 결합
        wall_gage = pd.merge(wall_gage, node, how='left', left_on='Node1', right_on='Node', suffixes=(None, '1'))
        wall_gage = pd.merge(wall_gage, node, how='left', left_on='Node2', right_on='Node', suffixes=(None, '2'))
        wall_gage = pd.merge(wall_gage, node, how='left', left_on='Node3', right_on='Node', suffixes=(None, '3'))
        wall_gage = pd.merge(wall_gage, node, how='left', left_on='Node4', right_on='Node', suffixes=(None, '4'))
        
        ### 부재의 orientation 맞추기        
        
        # X 좌표가 더 작은 노드를 i-node로!
        # 허용 오차
        tolerance = 5 # mm
        wall_gage_pos = wall_gage[wall_gage['X(mm)2'] - wall_gage['X(mm)'] > tolerance]
        wall_gage_neg = wall_gage[wall_gage['X(mm)2'] - wall_gage['X(mm)'] < -tolerance]
        wall_gage_zero = wall_gage[(wall_gage['X(mm)2'] - wall_gage['X(mm)'] >= -tolerance) 
                                   & (wall_gage['X(mm)2'] - wall_gage['X(mm)'] <= tolerance)]
        
        # node1(5,6,7,8), node2(9,10,11,12), node3(13,14,15,16), node4(17,18,19,20)
        wall_gage_neg = wall_gage_neg.iloc[:,[0,2,1,4,3,9,10,11,12,5,6,7,8,17,18,19,20,13,14,15,16]]
        wall_gage_neg.columns = wall_gage_pos.columns.values
        
        # Y 좌표가 더 작은 노드를 i-node로!
        wall_gage_zero_pos = wall_gage_zero[wall_gage_zero['Y(mm)2'] >= wall_gage_zero['Y(mm)']]
        wall_gage_zero_neg = wall_gage_zero[wall_gage_zero['Y(mm)2'] < wall_gage_zero['Y(mm)']]
        
        wall_gage_zero_neg = wall_gage_zero_neg.iloc[:,[0,2,1,4,3,9,10,11,12,5,6,7,8,17,18,19,20,13,14,15,16]]
        wall_gage_zero_neg.columns = wall_gage_zero_pos.columns.values
        
        # pos, neg 합치기
        wall_gage = pd.concat([wall_gage_pos, wall_gage_neg, wall_gage_zero_pos, wall_gage_zero_neg]\
                              , ignore_index=True)
        
        # 필요한 열 뽑고 재정렬
        wall_gage = wall_gage.iloc[:,[0,1,2,3,4,6,7,8,10,11,12,14,15,16,18,19,20]]
        wall_gage.columns = ['Wall ID', 'Node1', 'Node2', 'Node3', 'Node4', 'X(mm)1'\
                             , 'Y(mm)1', 'Z(mm)1', 'X(mm)2', 'Y(mm)2', 'Z(mm)2', 'X(mm)3'\
                             , 'Y(mm)3', 'Z(mm)3', 'X(mm)4', 'Y(mm)4', 'Z(mm)4']
        
        # 각 Wall element의 z좌표 추출
        wall_gage['Z(mm)'] =  wall_gage[['Z(mm)1', 'Z(mm)2', 'Z(mm)3', 'Z(mm)4']].min(axis=1)
              
        # 벽체의 4개 node list 만들기
        wall_gage['Node List'] = wall_gage.loc[:,['Node1', 'Node2', 'Node3', 'Node4']]\
                                 .values.tolist()
        
        # 같은 Wall ID, Z(mm)에 따라 Sorting              
        wall_gage_sorted = wall_gage.loc[:,['Wall ID', 'Z(mm)', 'Node List']]\
                           .set_index(['Wall ID', 'Z(mm)'])
        
        wall_gage_sorted = wall_gage_sorted.sort_values(['Wall ID', 'Z(mm)'])
               
        # For loop 돌리면서 Wall ID, Z(mm)에 따른 Node Data 리스트/ 겹치는 Node 리스트 만들기
        # For loop 돌리면서 Wall ID, Z(mm)에 따라 Node 리스트 업데이트(겹치는거 없애면서)
        gage_node_data_list = []
        for idx, gage_node_data in wall_gage_sorted.groupby(['Wall ID', 'Z(mm)'])['Node List']:
            # series -> list
            gage_node_data = list(gage_node_data)
            # deque 생성
            gage_node_dq = deque()
            
            # 노드를 위치 순서대로 deque에 insert
            for i in range(0,len(gage_node_data)):
                gage_node_dq.insert(int(i*1+0), gage_node_data[i][0])
                gage_node_dq.insert(int(i*2+1), gage_node_data[i][1])
                gage_node_dq.insert(int(i*3+2), gage_node_data[i][2])
                gage_node_dq.insert(int(i*4+3), gage_node_data[i][3])
            gage_node_dq = list(gage_node_dq) 
            
            # count == 1인 노드만 추출(중복되는 노드들 제거하는 법 몰라서 우회함)
            gage_node_dq_flat = []
            for i in gage_node_dq:
                if gage_node_dq.count(i) == 1:
                    gage_node_dq_flat.append(i)
            
            # 합쳐져 있는 노드들 분류
            for i in range(0,len(gage_node_dq_flat)//4):
                temp = [gage_node_dq_flat[i+len(gage_node_dq_flat)//4*0]
                        , gage_node_dq_flat[i+len(gage_node_dq_flat)//4*1]
                        , gage_node_dq_flat[i+len(gage_node_dq_flat)//4*2]
                        , gage_node_dq_flat[i+len(gage_node_dq_flat)//4*3]]        
                gage_node_data_list.append(temp)
               
        # Node 번호에 맞는 좌표 매칭 후 출력
        gage_node_coord = pd.DataFrame(gage_node_data_list)
        gage_node_coord.columns = ['Node1', 'Node2', 'Node3', 'Node4']
        
        # WR_gage_node를 as_gage_node로 나누고 재배열
        WAS_gage_node_coord_1 = gage_node_coord[['Node1', 'Node4']]
        WAS_gage_node_coord_2 = gage_node_coord[['Node2', 'Node3']]
        
        WAS_gage_node_coord_1.columns = ['Node1', 'Node2']
        WAS_gage_node_coord_2.columns = ['Node1', 'Node2']
        WAS_gage_node_coord = pd.concat([WAS_gage_node_coord_1, WAS_gage_node_coord_2])
        
        WAS_gage_node_coord.drop_duplicates(inplace=True)
        
        # Merge로 Node 번호에 맞는 좌표를 결합
        WAS_gage_node_coord = pd.merge(WAS_gage_node_coord, node, how='left', left_on='Node1', right_on='Node', suffixes=(None, '1'))
        WAS_gage_node_coord = pd.merge(WAS_gage_node_coord, node, how='left', left_on='Node2', right_on='Node', suffixes=(None, '2'))
        
        WAS_gage_node_coord = WAS_gage_node_coord.iloc[:,[3,4,5,7,8,9]]
        
        # Gage Element 결과값을 csv로 변환
        WAS_gage_node_coord.to_csv(output_csv_dir+'\\'+WAS_gage_csv, mode='w', index=False)

    #%% Shear Wall Gage 뽑기
    
    if import_WR_gage == True:
        
        # Story Info data 불러오기
        story_info = pd.read_excel(input_xlsx_path, sheet_name=story_info_xlsx_sheet\
                                   , skiprows=[0,2,3], usecols=[1,2,3,4], keep_default_na=False)
        
        # Wall Element만 추출(slicing)
        wall = element.loc[lambda x: element['Type'] == 'WALL', :]
        
        wall_gage = wall.loc[:,['Wall ID', 'Node1', 'Node2', 'Node3', 'Node4']]
        
        # Merge로 Node 번호에 맞는 좌표를 결합
        wall_gage = pd.merge(wall_gage, node, how='left', left_on='Node1', right_on='Node', suffixes=(None, '1'))
        wall_gage = pd.merge(wall_gage, node, how='left', left_on='Node2', right_on='Node', suffixes=(None, '2'))
        wall_gage = pd.merge(wall_gage, node, how='left', left_on='Node3', right_on='Node', suffixes=(None, '3'))
        wall_gage = pd.merge(wall_gage, node, how='left', left_on='Node4', right_on='Node', suffixes=(None, '4'))
        
        ### 부재의 orientation 맞추기
        # 허용 오차
        tolerance = 5 # mm
        wall_gage_pos = wall_gage[wall_gage['X(mm)2'] - wall_gage['X(mm)'] > tolerance]
        wall_gage_neg = wall_gage[wall_gage['X(mm)2'] - wall_gage['X(mm)'] < -tolerance]
        wall_gage_zero = wall_gage[(wall_gage['X(mm)2'] - wall_gage['X(mm)'] >= -tolerance) 
                                   & (wall_gage['X(mm)2'] - wall_gage['X(mm)'] <= tolerance)]
        
        # node1(5,6,7,8), node2(9,10,11,12), node3(13,14,15,16), node4(17,18,19,20)
        wall_gage_neg = wall_gage_neg.iloc[:,[0,2,1,4,3,9,10,11,12,5,6,7,8,17,18,19,20,13,14,15,16]]
        wall_gage_neg.columns = wall_gage_pos.columns.values
        
        # Y 좌표가 더 작은 노드를 i-node로!
        wall_gage_zero_pos = wall_gage_zero[wall_gage_zero['Y(mm)2'] >= wall_gage_zero['Y(mm)']]
        wall_gage_zero_neg = wall_gage_zero[wall_gage_zero['Y(mm)2'] < wall_gage_zero['Y(mm)']]
        
        wall_gage_zero_neg = wall_gage_zero_neg.iloc[:,[0,2,1,4,3,9,10,11,12,5,6,7,8,17,18,19,20,13,14,15,16]]
        wall_gage_zero_neg.columns = wall_gage_zero_pos.columns.values
        
        # pos, neg 합치기
        wall_gage = pd.concat([wall_gage_pos, wall_gage_neg, wall_gage_zero_pos, wall_gage_zero_neg]\
                              , ignore_index=True)
        
        # 필요한 열 뽑고 재정렬
        wall_gage = wall_gage.iloc[:,[0,1,2,3,4,6,7,8,10,11,12,14,15,16,18,19,20]]
        wall_gage.columns = ['Wall ID', 'Node1', 'Node2', 'Node3', 'Node4', 'X(mm)1'\
                             , 'Y(mm)1', 'Z(mm)1', 'X(mm)2', 'Y(mm)2', 'Z(mm)2', 'X(mm)3'\
                             , 'Y(mm)3', 'Z(mm)3', 'X(mm)4', 'Y(mm)4', 'Z(mm)4']
        
        # 각 Wall element의 z좌표 추출
        wall_gage['Z(mm)'] =  wall_gage[['Z(mm)1', 'Z(mm)2', 'Z(mm)3']].min(axis=1)
              
        # 벽체의 4개 node list 만들기
        wall_gage['Node List'] = wall_gage.loc[:,['Node1', 'Node2', 'Node3', 'Node4']]\
                                 .values.tolist()
        
        # 같은 Wall ID, Z(mm)에 따라 Sorting              
        wall_gage_sorted = wall_gage.loc[:,['Wall ID', 'Z(mm)', 'Node List']]\
                           .set_index(['Wall ID', 'Z(mm)'])
        
        wall_gage_sorted = wall_gage_sorted.sort_values(['Wall ID', 'Z(mm)'])
               
        # For loop 돌리면서 Wall ID, Z(mm)에 따른 Node Data 리스트/ 겹치는 Node 리스트 만들기
        # For loop 돌리면서 Wall ID, Z(mm)에 따라 Node 리스트 업데이트(겹치는거 없애면서)
        gage_node_data_list = []
        for idx, gage_node_data in wall_gage_sorted.groupby(['Wall ID', 'Z(mm)'])['Node List']:
            gage_node_data = list(gage_node_data)
            # deque 생성
            gage_node_dq = deque()
            
            # 노드를 위치 순서대로 deque에 insert
            for i in range(0,len(gage_node_data)):
                gage_node_dq.insert(int(i*1+0), gage_node_data[i][0])
                gage_node_dq.insert(int(i*2+1), gage_node_data[i][1])
                gage_node_dq.insert(int(i*3+2), gage_node_data[i][2])
                gage_node_dq.insert(int(i*4+3), gage_node_data[i][3])
            gage_node_dq = list(gage_node_dq) 
            
            # count == 1인 노드만 추출(중복되는 노드들 제거하는 법 몰라서 우회함)
            gage_node_dq_flat = []
            for i in gage_node_dq:
                if gage_node_dq.count(i) == 1:
                    gage_node_dq_flat.append(i)
            
            # 합쳐져 있는 노드들 분류
            for i in range(0,len(gage_node_dq_flat)//4):
                temp = [gage_node_dq_flat[i+len(gage_node_dq_flat)//4*0]
                        , gage_node_dq_flat[i+len(gage_node_dq_flat)//4*1]
                        , gage_node_dq_flat[i+len(gage_node_dq_flat)//4*2]
                        , gage_node_dq_flat[i+len(gage_node_dq_flat)//4*3]]        
                gage_node_data_list.append(temp)
               
        # Node 번호에 맞는 좌표 매칭 후 출력
        gage_node_coord = pd.DataFrame(gage_node_data_list)
        gage_node_coord.columns = ['Node1', 'Node2', 'Node3', 'Node4']
        
        # 벽체 노드 순서 바꾸기(Midas 1234 -> Perform-3d 1243)
        # gage_node_coord = gage_node_coord.iloc[:,[0,1,3,2]]
        
        # Merge로 Node 번호에 맞는 좌표를 결합
        gage_node_coord = pd.merge(gage_node_coord, node, how='left', left_on='Node1', right_on='Node', suffixes=(None, '1'))
        gage_node_coord = pd.merge(gage_node_coord, node, how='left', left_on='Node2', right_on='Node', suffixes=(None, '2'))
        gage_node_coord = pd.merge(gage_node_coord, node, how='left', left_on='Node3', right_on='Node', suffixes=(None, '3'))
        gage_node_coord = pd.merge(gage_node_coord, node, how='left', left_on='Node4', right_on='Node', suffixes=(None, '4'))
    
        gage_node_coord = gage_node_coord.iloc[:, [5,6,7,9,10,11,17,18,19,13,14,15]]

        
        ### WR gage가 분할층에서 나눠지지 않게 만들기 
        # 분할층 노드가 포함되지 않은 부재 slice
        gage_node_coord_no_divide = gage_node_coord[(gage_node_coord['Z(mm)'].isin(story_info['Level']))\
                                                    & (gage_node_coord['Z(mm)3'].isin(story_info['Level']))]
        
        # 분할층 노드가 상부에만(k,l-node) 포함되는 부재 slice
        gage_node_coord_divide = gage_node_coord[(gage_node_coord['Z(mm)'].isin(story_info['Level']))\
                                                 & (~gage_node_coord['Z(mm)3'].isin(story_info['Level']))]
        
        # gage_node_coord_divide 노드들의 상부 노드(k,l-node)의 z좌표를 다음 측으로 격상
        next_level_list = []
        for i in gage_node_coord_divide['Z(mm)3']:
            level_bigger = story_info['Level'][story_info['Level']-i >= 0]
            next_level = level_bigger.sort_values(ignore_index=True)[0]

            next_level_list.append(next_level)
        
        pd.options.mode.chained_assignment = None # SettingWithCopyWarning 안뜨게 하기

        gage_node_coord_divide.loc[:, 'Z(mm)3'] = next_level_list
        gage_node_coord_divide.loc[:, 'Z(mm)4'] = next_level_list
        
        gage_node_coord = pd.concat([gage_node_coord_no_divide, gage_node_coord_divide]\
                                    , ignore_index=True)
        
        # Gage Element 결과값을 csv로 변환
        gage_node_coord.to_csv(output_csv_dir+'\\'+WR_gage_csv, mode='w', index=False)
    
    #%% Imbedded Beam 뽑기
    
    if import_I_beam == True:
        
        # Beam, Wall Element 추출(slicing)
        frame = element[element['Type'] == 'BEAM']
        wall = element[element['Type'] == 'WALL']

        # 필요한 열만 추출(drop하기에는 drop할 열이 너무 많아서...)
        frame_node_1 = frame.loc[:, 'Node1']
        frame_node_2 = frame.loc[:, 'Node2']
        wall_node_3 = wall.loc[:, 'Node3']
        wall_node_4 = wall.loc[:, 'Node4']
        
        # Merge로 Node 번호에 맞는 좌표를 결합
        frame_node_1_coord = pd.merge(frame_node_1, node, how='left', left_on='Node1', right_on='Node')  # how='left' : 두 데이터프레임 중 왼쪽 데이터프레임은 그냥 두고 오른쪽 데이터프레임값을 대응시킴
        frame_node_2_coord = pd.merge(frame_node_2, node, how='left', left_on='Node2', right_on='Node')
        wall_node_3_coord = pd.merge(wall_node_3, node, how='left', left_on='Node3', right_on='Node')
        wall_node_4_coord = pd.merge(wall_node_4, node, how='left', left_on='Node4', right_on='Node')
        
        # Node1, Node2의 좌표를 모두 결합시켜 출력
        frame_node_1_coord = frame_node_1_coord.drop('Node', axis=1)
        frame_node_2_coord = frame_node_2_coord.drop('Node', axis=1)
        wall_node_3_coord = wall_node_3_coord.drop('Node', axis=1)
        wall_node_4_coord = wall_node_4_coord.drop('Node', axis=1)
        
        frame_node_1_coord.columns = ['Node1', 'X_1(mm)', 'Y_1(mm)', 'Z_1(mm)']  # 결합 때 이름이 중복되면 안되서 이름 바꿔줌
        frame_node_2_coord.columns = ['Node2', 'X_2(mm)', 'Y_2(mm)', 'Z_2(mm)']
        wall_node_3_coord.columns = ['Node3', 'X_3(mm)', 'Y_3(mm)', 'Z_3(mm)']
        wall_node_4_coord.columns = ['Node4', 'X_4(mm)', 'Y_4(mm)', 'Z_4(mm)']
        
        frame_node_coord = pd.concat([frame_node_1_coord, frame_node_2_coord], axis=1)
        wall_node_coord = pd.concat([wall_node_3_coord, wall_node_4_coord], axis=1)
        
        # Beam 추출 (Column 제외)
        beam_node_coord = frame_node_coord[abs(frame_node_coord['Z_1(mm)'] - frame_node_coord['Z_2(mm)']) <= 10]   
        
        # node1-node2, node3-node4의 방향 vector 생성
        beam_node_coord['X_1-X_2'] = beam_node_coord['X_1(mm)'] - beam_node_coord['X_2(mm)']
        beam_node_coord['Y_1-Y_2'] = beam_node_coord['Y_1(mm)'] - beam_node_coord['Y_2(mm)']
        wall_node_coord['X_3-X_4'] = wall_node_coord['X_3(mm)'] - wall_node_coord['X_4(mm)']
        wall_node_coord['Y_3-Y_4'] = wall_node_coord['Y_3(mm)'] - wall_node_coord['Y_4(mm)']

        # N1-N2, N3-N4 벡터의 Cosine Similarity 구하기
        # (running time 단축을 위해 아래의 두 function은 np.array 데이터 형식으로 계산함)
        beam_node_coord_np = beam_node_coord.to_numpy()
        wall_node_coord_np = wall_node_coord.to_numpy()
        
        # 두 벡터(array)의 Cosine Similarity 구하는 함수
        def cos_sim(vector1, vector2):
            result = np.dot(vector1, vector2) / (np.linalg.norm(vector1)*np.linalg.norm(vector2))
            return result
        
        # 두 개의 행렬(matrix)를 입력받아 i_beam 정보 찾기
        def find_i_beam(matrix1, matrix2): # matrix 형태 : beam_node_coord_np, wall_node_coord_np
            i_beam_list = []
            for matrix1_row in matrix1:
                for matrix2_row in matrix2:
                    vector1 = np.array([matrix1_row[8], matrix1_row[9]])
                    vector2 = np.array([matrix2_row[8], matrix2_row[9]])
                    
                    # N1=N3 or N1=N4 or N2=N3 or N2=N4인 경우
                    if (matrix1_row[0] == matrix2_row[0]) | (matrix1_row[0] == matrix2_row[4])\
                        | (matrix1_row[4] == matrix2_row[0]) | (matrix1_row[4] == matrix2_row[4]):
                        # 방향 벡터가 같은 경우
                        if abs(cos_sim(vector1, vector2)) >= 0.98:
                            i_beam_list.append(matrix2_row)
            # list of arrays -> array
            i_beam_matrix = np.vstack(i_beam_list)
            # Drop duplicates
            i_beam_matrix = np.unique(i_beam_matrix, axis=0)
            # print(beam_node_coord_np.shape[0]), print(wall_node_coord_np.shape[0]), print(i_beam_matrix.shape[0])
            
            return i_beam_matrix
        
        # 무한루프 돌리면서 Beam과 바로 만나는 Imbedded Beam부터 순서대로 찾기
        # import time
        # time_start = time.time()
        
        i_beam_matrix = beam_node_coord_np.copy()
        while True:
            # 
            i_beam_matrix_updated = find_i_beam(i_beam_matrix, wall_node_coord_np)
            print(i_beam_matrix_updated.shape[0])
            
            # if np.array_equal(matrix1, matrix3):
            if i_beam_matrix.shape[0] == i_beam_matrix_updated.shape[0]:
                break
            
            i_beam_matrix = i_beam_matrix_updated.copy()
            
        # time_end = time.time()
        # time_run = (time_end-time_start)/60
        # print('\n', 'total time = %0.7f min' %(time_run))
        
        # 기존에 있던 보와 동일한 위치에 생성된 Imbedded Beam 제거
        def view1D(a, b): # a, b are arrays
            a = np.ascontiguousarray(a)
            b = np.ascontiguousarray(b)
            void_dt = np.dtype((np.void, a.dtype.itemsize * a.shape[1]))
            return a.view(void_dt).ravel(),  b.view(void_dt).ravel()

        def setdiff_nd(a,b):
            # a,b are the nD input arrays
            A,B = view1D(a,b)    
            return a[~np.isin(A,B)]
        
        beam_node_coord_np_rev = beam_node_coord_np[:,[4,5,6,7,0,1,2,3]]
        i_beam_matrix_unique = setdiff_nd(i_beam_matrix[:,0:8], beam_node_coord_np[:,0:8])
        i_beam_matrix_unique = setdiff_nd(i_beam_matrix_unique, beam_node_coord_np_rev)
    
        
        # np.array -> pd.dataframe
        I_beam_node_coord = pd.DataFrame(i_beam_matrix_unique)
        I_beam_node_coord.columns = ['Node1', 'X_1(mm)', 'Y_1(mm)', 'Z_1(mm)'
                                     , 'Node2', 'X_2(mm)', 'Y_2(mm)', 'Z_2(mm)']
        
        I_beam_node_coord = I_beam_node_coord.loc[:,['X_1(mm)', 'Y_1(mm)', 'Z_1(mm)', 'X_2(mm)', 'Y_2(mm)', 'Z_2(mm)']]
            
        # 출력
        I_beam_node_coord.to_csv(output_csv_dir+'\\'+I_beam_csv, mode='w', index=False)         
                
    #%% Plate Element 뽑기
    if (import_plate == True) or ('PLATE' in element['Type']):
        
    # Plate Element만 추출(slicing)
        plate = element.loc[lambda x: element['Type'] == 'PLATE', :]
        
        # 필요한 열만 추출
        plate_node_1 = plate.loc[:, 'Node1']
        plate_node_2 = plate.loc[:, 'Node2']
        plate_node_3 = plate.loc[:, 'Node3']
        plate_node_4 = plate.loc[:, 'Node4']
        
        plate_node_1.name = 'Node'
        plate_node_2.name = 'Node'
        plate_node_3.name = 'Node'
        plate_node_4.name = 'Node'
        
        # Merge로 Node 번호에 맞는 좌표를 결합
        plate_node_1_coord = pd.merge(plate_node_1, node, how='left')
        plate_node_2_coord = pd.merge(plate_node_2, node, how='left')
        plate_node_3_coord = pd.merge(plate_node_3, node, how='left')
        plate_node_4_coord = pd.merge(plate_node_4, node, how='left')
        
        # Node1, Node2, Node3, Node4의 좌표를 모두 결합시켜 출력
        plate_node_1_coord = plate_node_1_coord.drop('Node', axis=1)
        plate_node_2_coord = plate_node_2_coord.drop('Node', axis=1)
        plate_node_3_coord = plate_node_3_coord.drop('Node', axis=1)
        plate_node_4_coord = plate_node_4_coord.drop('Node', axis=1)
        
        plate_node_1_coord.columns = ['X_1(mm)', 'Y_1(mm)', 'Z_1(mm)']
        plate_node_2_coord.columns = ['X_2(mm)', 'Y_2(mm)', 'Z_2(mm)']
        plate_node_3_coord.columns = ['X_3(mm)', 'Y_3(mm)', 'Z_3(mm)']
        plate_node_4_coord.columns = ['X_4(mm)', 'Y_4(mm)', 'Z_4(mm)']
        
        # plate_node_coord_list = [plate_node_1_coord, plate_node_2_coord, plate_node_3_coord, plate_node_4_coord]
        plate_node_coord = pd.concat([plate_node_1_coord, plate_node_2_coord, plate_node_3_coord, plate_node_4_coord], axis=1)
        
        # plate Element 결과값을 csv로 변환
        plate_node_coord.to_csv(output_csv_dir+'\\'+plate_csv, mode='w', index=False)
                
#%% Frame, Element, Section , Drift, Constraint Naming

def naming(input_xlsx_path, drift_position=[2,5,7,11]):
    '''
    
    모델링에 사용되는 모든 이름들을 동일한 규칙에 의해 출력함.
    (Frame, Constraints, Section, Drift Names)
    
    Parameters
    ----------
    input_path : str
                 Data Conversion 엑셀 파일의 경로.
                 
    input_xlsx : str
                 Data Conversion 엑셀 파일의 이름. 확장자명(.xlsx)까지 기입해줘야한다. 하나의 파일만 불러온다.
    
    drift_position : list of int, optional, default=[2,5,7,11]
                     drift 게이지를 설치할 위치. 대괄호 안에는 반드시 정수를 입력해야하며, 각각의 정수는 방향(시계)을 의미한다. 

    Returns
    -------        
    name_output : pandas.core.frame.DataFrame or None
    따로 출력되지 않으며, Data Conversion 엑셀파일의 Output_Naming 시트에 자동 입력됨.
                   
    Raises
    -------
    
    '''

    #%% wall, frame 이름 만들기 위한 정보 load
    
    naming_info_xlsx_sheet = 'Naming' # wall naming 관련된 정보만 들어있는 시트
    story_info_xlsx_sheet = 'Story Data' # 층 정보 sheet
    
    naming_info = pd.read_excel(input_xlsx_path\
                                , sheet_name = naming_info_xlsx_sheet\
                                , skiprows = 3, usecols=[0,1,2,3,4,5,6,7,8,9,10,11])
    
    # Wall에 대해 정리
    wall_info = naming_info.iloc[:,[8,9,10,11]]
    wall_info.columns = ['Name', 'Story(from)', 'Story(to)', 'Amount']
    wall_info = wall_info[wall_info['Name'].notna()]

    # Beam에 대해서도 똑같이...
    beam_info = naming_info.iloc[:,[0,1,2,3]]
    beam_info.columns = ['Name', 'Story(from)', 'Story(to)', 'Amount']
    beam_info = beam_info[beam_info['Name'].notna()]
    
    # Column에 대해서도 똑같이...
    column_info = naming_info.iloc[:,[4,5,6,7]]
    column_info.columns = ['Name', 'Story(from)', 'Story(to)', 'Amount']
    column_info = column_info[column_info['Name'].notna()]
    
    #%% story 정보 load
    story_info = pd.read_excel(input_xlsx_path\
                               , sheet_name = story_info_xlsx_sheet\
                               , skiprows = [0,2,3], usecols=[0,1,2,3,4])

    story_info_reversed = story_info[::-1]
    story_info_reversed.reset_index(inplace=True, drop=True)
    # 배열이 내가 원하는 방향과 반대로 되어있어서, 리스트 거꾸로만들었음

    #%% Section 이름 뽑기
    if wall_info.shape[0] != 0:
    
        # for문으로 wall naming에 사용할 섹션 이름(wall_name_output) 뽑기
        wall_name_output = [] # 결과로 나올 wall_name_output 리스트 미리 정의
    
        for wall_name_parameter, amount_parameter, story_from_parameter, story_to_parameter\
            in zip(wall_info['Name'], wall_info['Amount'], wall_info['Story(from)'], wall_info['Story(to)']):  # for 문에 조건 여러개 달고싶을 때는 zip으로 묶어서~ 
            
            story_from_index = story_info_reversed[story_info_reversed['Story Name'] == story_from_parameter].index[0]  # story_from이 문자열이라 story_from을 사용해서 slicing이 안되기 때문에(내 지식선에서) .index로 story_from의 index만 뽑음
            story_to_index = story_info_reversed[story_info_reversed['Story Name'] == story_to_parameter].index[0]  # 마찬가지로 story_to의 index만 뽑음
            story_window = story_info_reversed['Story Name'][story_from_index : story_to_index + 1]  # 내가 원하는 층 구간(story_from부터 story_to까지)만 뽑아서 리스트로 만들기
            for i in range(1, amount_parameter + 1):  # (벽체 개수(amount))에 맞게 numbering하기 위해 1,2,3,4...amount[i]개의 배열을 만듦. 첫 시작을 1로 안하면 index 시작은 0이 default값이기 때문에 1씩 더해줌
                for current_story_name in story_window:
                    if isinstance(current_story_name, str) == False:  # 층이름이 int인 경우, 이름조합을 위해 str로 바꿈
                        current_story_name = str(current_story_name)
                    else:
                        pass
                    
                    wall_name_output.append(wall_name_parameter + '_' + str(i) + '_' + current_story_name)  # 반복될때마다 생성되는 section 이름을 .append를 이용하여 리스트의 끝에 하나씩 쌓아줌. i값은 숫자라 .astype(str)로 string으로 바꿔줌
    
        # 층전단력 확인을 위한 층 섹션 이름 뽑기
        # Base section 추가하기
        story_section_name_output = ['Base']
    
        # 각 층 전단력 확인을 위한 각 층 section 추가하기
        for i in story_info_reversed['Story Name'][1:story_info_reversed.shape[0]-1]:
            story_section_name_output.append(i + '_Shear')

    #%% Frame 이름 뽑기
        
    # Wall Frame 이름 뽑기
    frame_wall_name_output = []

    for row in wall_info.values: # for문을 빠르게 연산하기 위해 dataframe -> array    
        for i in range(1, int(row[3]) + 1):  
            frame_wall_name_output.append(row[0] + '_' + str(i))
            
    # Beam Frame 이름 뽑기
    frame_beam_name_output = []

    for row in beam_info.values:    
        for i in range(1, int(row[3]) + 1):
            frame_beam_name_output.append(row[0] + '_' + str(i))
            
    # Column Frame 이름 뽑기
    if column_info.shape[0] != 0:
        frame_column_name_output = []
    
        for row in column_info.values:    
            for i in range(1, int(row[3]) + 1):
                frame_column_name_output.append(row[0] + '_' + str(i))
        frame_column_name_output = pd.Series(frame_column_name_output)
        
    else:
        frame_column_name_output = pd.Series([], dtype='float64')
            
    #%% Constraints 이름 뽑기

    constraints_name = []

    for row in story_info_reversed.values:
        if row[4] >= 2:
            for i in range(1, int(row[4]) + 1):
                constraints_name.append(row[1] + '-' + str(i))
        else: constraints_name.append(row[1])
        
    constraints_name_output = constraints_name[1:]

    #%% Drift 이름 뽑기

    # Drift의 방향 지정
    direction_list = ['X', 'Y']

    drift_name_output = []

    for position in drift_position:
        for direction in direction_list:
            for current_story_name in story_info['Story Name'][1:story_info.shape[0]]:
                if isinstance(current_story_name, str) == False:  # 층이름이 int인 경우, 이름조합을 위해 str로 바꿈
                    current_story_name = str(current_story_name)
                drift_name_output.append(current_story_name + '_' + str(int(position)) + '_' + direction)
                    
    #%% 출력

    name_output = pd.DataFrame(({'Frame(Beam) Name': pd.Series(frame_beam_name_output),\
                                  'Frame(Column) Name': pd.Series(frame_column_name_output),\
                                  'Frame(Wall) Name': pd.Series(frame_wall_name_output),\
                                  'Constraints Name': pd.Series(constraints_name_output),\
                                  'Section(Wall) Name': pd.Series(wall_name_output),\
                                  'Section(Shear) Name': pd.Series(story_section_name_output),\
                                  'Drift Name': pd.Series(drift_name_output)}))

    # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
    name_output = name_output.replace(np.nan, '', regex=True)
    
    # Using win32com...
    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게

    wb = excel.Workbooks.Open(input_xlsx_path)
    ws = wb.Sheets('Output_Naming')
    
    startrow, startcol = 5, 1

    # 이름 열 입력
    ws.Range(ws.Cells(startrow, startcol),\
              ws.Cells(startrow + name_output.shape[0]-1,\
                      name_output.shape[1])).Value\
    = list(name_output.itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능   
    
    wb.Save()
    # wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application  

#%% Convert C.Beam, G.Beam, Wall

def convert_property(input_xlsx_path, get_beam=True, get_column=True, get_wall=True):
    '''
    
    User가 입력한 부재 정보들을 Perform-3D에 입력할 수 있는 형식으로 변환하여 Data Conversion 엑셀파일의 Output_Properties 시트에 작성.
    
    Parameters
    ----------
    input_path : str
                 Data Conversion 엑셀 파일의 경로.
                 
    input_xlsx : str
                 Data Conversion 엑셀 파일의 이름. 확장자명(.xlsx)까지 기입해줘야한다. 하나의 파일만 불러온다.

    get_beam : bool, optional, default=True
               True = C.Beam의 정보를 Perform-3D 입력용 정보로 변환함.
               False = C.Beam의 정보를 변환하지 않음.
               
    get_column : bool, optional, default=True
                 True = G.Column의 정보를 Perform-3D 입력용 정보로 변환함.
                 False = G.Column의 정보를 변환하지 않음.
               
    get_wall : bool, optional, default=True
               True = Wall의 정보를 Perform-3D 입력용 정보로 변환함.
               False = Wall의 정보를 변환하지 않음.

    Returns
    --------       
    beam_output : pandas.core.frame.DataFrame or None
                  C.Beam Properties의 정보를 Perform-3D 입력용으로 변환한 정보.
                  Output_C.Beam Properties 시트에 입력됨.
                     
    wall_output : pandas.core.frame.DataFrame or None
                  Wall Properties의 정보를 Perform-3D 입력용으로 변환한 정보.
                  Output_Wall Properties 시트에 입력됨.   
                  
    Raises
    -------
    
    '''    
    #%% 파일 load
    
    pd.options.mode.chained_assignment = None # SettingWithCopyWarning 안뜨게 하기
    # UserWarning: openpyxl 안뜨게 하기
    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
    
    input_data_raw = pd.ExcelFile(input_xlsx_path)
    input_data_sheets = pd.read_excel(input_data_raw\
                                      , ['C.Beam Properties', 'Wall Properties'\
                                         , 'G.Column Properties', 'Story Data'
                                         , 'ETC', 'Naming'], skiprows=3)
    input_data_raw.close()
    
    # Wall 정보 load
    wall = input_data_sheets['Wall Properties'].iloc[:,np.r_[0:11, 21,22]]
    wall.columns = ['Name', 'Story(from)', 'Story(to)', 'Thickness', 'Vertical Rebar(DXX)',\
                    'V. Rebar Space', 'Horizontal Rebar(DXX)', 'H. Rebar Space', 'Type', 'Length', 'Element length', 'Fibers(Concrete)', 'Fibers(Rebar)']

    wall = wall.dropna(axis=0, how='all')
    wall.reset_index(inplace=True, drop=True)
    
    saved_wall_story_from = wall['Story(from)']
    saved_wall_story_to = wall['Story(to)']
    
    wall = wall.fillna(method='ffill')
    
    wall['Story(from)'] = saved_wall_story_from
    wall['Story(to)'] = saved_wall_story_to

    # Column 정보 load
    column = input_data_sheets['G.Column Properties'].iloc[:,0:18]
    column.columns = ['Name', 'Story(from)', 'Story(to)', 'b(mm)', 'h(mm)'
                      , 'Cover Thickness(mm)', '내진상세 여부', 'Type(Main)'
                      , 'Main Rebar(DXX)', 'Type(Hoop)', 'Hoop Rebar(DXX)'
                      , 'EA(Layer1)', 'Row(Layer1)', 'EA(Layer2)', 'Row(Layer2)'
                      , 'EA(Hoop_X)', 'EA(Hoop_Y)', 'Spacing(Hoop)']

    column = column.dropna(axis=0, how='all')
    column.reset_index(inplace=True, drop=True)
    
    saved_column_story_from = column['Story(from)']
    saved_column_story_to = column['Story(to)']
    saved_column_rebar = column.iloc[:,[11,12,13,14,15,16,17]]
    
    column = column.fillna(method='ffill')
    
    column['Story(from)'] = saved_column_story_from
    column['Story(to)'] = saved_column_story_to
    column.iloc[:,[11,12,13,14,15,16,17]] = saved_column_rebar

    # Beam 정보 load
    beam = input_data_sheets['C.Beam Properties'].iloc[:,0:21]
    beam.columns = ['Name', 'Story(from)', 'Story(to)', 'Length(mm)', 'b(mm)',\
                    'h(mm)', 'Cover Thickness(mm)', 'Type', '배근', '내진상세 여부',\
                    'Main Rebar(DXX)', 'Stirrup Rebar(DXX)', 'X-Bracing Rebar', 'Top(1)', 'Top(2)',\
                    'Top(3)', 'EA(Stirrup)', 'Spacing(Stirrup)', 'EA(Diagonal)', 'Degree(Diagonal)', 'D(mm)']

    beam = beam.dropna(axis=0, how='all')
    beam.reset_index(inplace=True, drop=True)
    
    saved_beam_story_from = beam['Story(from)']
    saved_beam_story_to = beam['Story(to)']
    saved_beam_rebar = beam.iloc[:,[12,13,14,15,16,17,18,19]]
    
    beam = beam.fillna(method='ffill')
    
    beam['Story(from)'] = saved_beam_story_from
    beam['Story(to)'] = saved_beam_story_to
    beam.iloc[:,[12,13,14,15,16,17,18,19]] = saved_beam_rebar

    # 구분 조건 load
    naming_criteria = input_data_sheets['ETC']

    # Story 정보 load
    story_info = input_data_sheets['Story Data'].iloc[:,0:3]
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    story_name = story_info.loc[:, 'Story Name']
    story_name = story_name[::-1]  # 층 이름 재배열
    story_name.reset_index(drop=True, inplace=True)

    # 벽체,기둥,보 개수 load
    num_of_elem = input_data_sheets['Naming']
    
    num_of_beam = num_of_elem.iloc[:,[0,3]]
    num_of_wall = num_of_elem.iloc[:,[8,11]]
    num_of_column = num_of_elem.iloc[:,[4,7]]
    
    num_of_beam = num_of_beam.dropna(axis=0)
    num_of_wall = num_of_wall.dropna(axis=0)
    num_of_column = num_of_column.dropna(axis=0)

    num_of_beam.columns = ['Name', 'EA']
    num_of_wall.columns = ['Name', 'EA']
    num_of_column.columns = ['Name', 'EA']

    #%% 부재 이름 설정할 때 필요한 함수들

    # 층 나누는 함수 (12F~15F)
    def str_div(temp_list):
        first = []
        second = []
        
        for i in temp_list:
            if '~' in i:
                first.append(i.split('~')[0])
                second.append(i.split('~')[1])
            elif '-' in i:
                second.append(i.split('-')[0])
                first.append(i.split('-')[1])
            else:
                first.append(i)
                second.append(i)
        
        first = pd.Series(first).str.strip()
        second = pd.Series(second).str.strip()
        
        return first, second

    # 층, 철근 나누는 함수 (12F~15F, D10@300)
    def rebar_div(temp_list1, temp_list2):
        first = []
        second = []
        third = []
        
        for i, j in zip(temp_list1, temp_list2):
            if isinstance(i, str) : # string인 경우
                if '@' in i:
                    first.append(i.split('@')[0].strip())
                    second.append(i.split('@')[1])
                    third.append(np.nan)
                elif '-' in i:
                    third.append(i.split('-')[0])
                    first.append(i.split('-')[1].strip())
                    second.append(np.nan)
                else: 
                    first.append(i.strip())
                    second.append(j)
                    third.append(np.nan)
            else: # string 아닌 경우
                first.append(i)
                second.append(j)
                third.append(np.nan)

        return first, second, third

    # 철근 지름 앞의 D 떼주는 함수 (D10...)
    def str_extract(sth_str):
        result = int(re.findall(r'[0-9]+', sth_str)[0])
        
        return result

    #%% 데이터베이스
    steel_geometry_database = naming_criteria.iloc[:,[0,1,2]].dropna()
    steel_geometry_database.columns = ['Name', 'Diameter(mm)', 'Area(mm^2)']

    new_steel_geometry_name = []

    for i in steel_geometry_database['Name']:
        if isinstance(i, int):
            new_steel_geometry_name.append(i)
        else:
            new_steel_geometry_name.append(str_extract(i))

    steel_geometry_database['Name'] = new_steel_geometry_name

    #%% 1. Wall
    #%% 불러온 wall 정보 정리하기
    if get_wall == True:
        
        # 글자가 합쳐져 있을 경우 글자 나누기 - 층 (12F~15F, D10@300)
        new_story = wall[['Story(from)', 'Story(to)']]
        new_story = new_story.fillna(method='ffill', axis=1)
              
        wall['Story(from)'] = new_story['Story(from)']
        wall['Story(to)'] = new_story['Story(to)']
    
        # V. Rebar 나누기
        v_rebar_div = rebar_div(wall['Vertical Rebar(DXX)'], wall['V. Rebar Space'])
        wall['Vertical Rebar(DXX)'] = v_rebar_div[0]
        wall['V. Rebar Space'] = v_rebar_div[1]
        wall['V. Rebar EA'] = v_rebar_div[2]
    
        # H. Rebar 나누기
        h_rebar_div = rebar_div(wall['Horizontal Rebar(DXX)'], wall['H. Rebar Space'])
        wall['Horizontal Rebar(DXX)'] = h_rebar_div[0]
        wall['H. Rebar Space'] = h_rebar_div[1]
    
        # 철근의 앞에붙은 D 떼어주기
        new_v_rebar = []
        new_h_rebar = []
    
        for i in wall['Vertical Rebar(DXX)']:
            if isinstance(i, int):
                new_v_rebar.append(i)
            else:
                new_v_rebar.append(str_extract(i))
                
        for j in wall['Horizontal Rebar(DXX)']:
            if isinstance(j, int):
                new_h_rebar.append(j)
            else:
                new_h_rebar.append(str_extract(j))
                
        wall['Vertical Rebar(DXX)'] = new_v_rebar
        wall['Horizontal Rebar(DXX)'] = new_h_rebar
    
        # Rebar Space 데이터값 모두 float로 바꿔주기
        v_rebar_spacing_float = []
        h_rebar_spacing_float = []
        v_rebar_ea_float = []
    
        for i, j, k in zip(wall['V. Rebar Space'], wall['H. Rebar Space'], wall['V. Rebar EA']):
            
            if not isinstance(i, float):
                v_rebar_spacing_float.append(float(i))
                
            else: v_rebar_spacing_float.append(i)
                
            if not isinstance(j, float):
                h_rebar_spacing_float.append(float(j))
                
            else: h_rebar_spacing_float.append(j)
            
            if not isinstance(k, float):
                v_rebar_ea_float.append(float(k))
                
            else: v_rebar_ea_float.append(k)
            
        wall['V. Rebar Space'] = v_rebar_spacing_float
        wall['H. Rebar Space'] = h_rebar_spacing_float
        wall['V. Rebar EA'] = v_rebar_ea_float
    
        #%% 이름 구분 조건 load & 정리
    
        # 층 구분 조건에  story_name의 index 매칭시켜서 새로 열 만들기
        naming_criteria_1_index = []
        naming_criteria_2_index = []
    
        for i, j in zip(naming_criteria.iloc[:,5].dropna(), naming_criteria.iloc[:,6].dropna()):
            naming_criteria_1_index.append(pd.Index(story_name).get_loc(i))
            naming_criteria_2_index.append(pd.Index(story_name).get_loc(j))
    
        ### 구분 조건이 층 순서에 상관없이 작동되게 재정렬
        # 구분 조건에 해당하는 콘크리트 강도 재정렬
        naming_criteria_property = pd.concat([pd.Series(naming_criteria_1_index, name='Story(from) Index'), naming_criteria.iloc[:,7].dropna()], axis=1)
    
        naming_criteria_property['Story(from) Index'] = pd.Categorical(naming_criteria_property['Story(from) Index'], naming_criteria_1_index.sort())
        naming_criteria_property.sort_values('Story(from) Index', inplace=True)
        naming_criteria_property.reset_index(inplace=True)
    
        # 구분 조건 재정렬
        naming_criteria_1_index.sort()
        naming_criteria_2_index.sort()
    
        #%% 시작층, 끝층 정리
    
        naming_from_index = []
        naming_to_index = []
    
        for naming_from, naming_to in zip(wall['Story(from)'], wall['Story(to)']):
            if isinstance(naming_from, str) == False:
                naming_from = str(naming_from)
            if isinstance(naming_to, str) == False:
                naming_from = str(naming_from)
                
            naming_from_index.append(pd.Index(story_name).get_loc(naming_from))
            naming_to_index.append(pd.Index(story_name).get_loc(naming_to))
    
        #%%  층 이름을 etc의 이름 구분 조건에 맞게 나누어서 리스트로 정리
    
        naming_from_index_list = []
        naming_to_index_list = []
        naming_criteria_property_index_list = []
    
        for current_naming_from_index, current_naming_to_index in zip(naming_from_index, naming_to_index):  # 부재의 시작과 끝 층 loop
            naming_from_index_sublist = [current_naming_from_index]
            naming_to_index_sublist = [current_naming_to_index]
            naming_criteria_property_index_sublist = []
                
            for i, j, k in zip(naming_criteria_1_index, naming_criteria_2_index, naming_criteria_property.index):
                if (i >= current_naming_from_index) and (i <= current_naming_to_index):
                    naming_from_index_sublist.append(i)
                    naming_criteria_property_index_sublist.append(k)
                                
                    if (j >= current_naming_from_index) and (j <= current_naming_to_index):
                        naming_to_index_sublist.append(j)
                    else:
                        naming_to_index_sublist.append(i-1)
                        
                    if i != current_naming_from_index:
                        naming_criteria_property_index_sublist.append(k-1)
                                            
                elif (i < current_naming_from_index) and (j >= current_naming_to_index):
                    naming_criteria_property_index_sublist.append(k)
                    
                elif (i < current_naming_from_index) and (j <= current_naming_to_index):
                    naming_to_index_sublist.append(j)
                    
                else:
                    if max(naming_criteria_1_index) < current_naming_from_index:
                        naming_criteria_property_index_sublist.append(max(naming_criteria_property.index))
                        
                    elif min(naming_criteria_1_index) > current_naming_to_index:
                            naming_criteria_property_index_sublist.append(min(naming_criteria_property.index))
                    
                naming_from_index_sublist = list(set(naming_from_index_sublist))
                naming_to_index_sublist = list(set(naming_to_index_sublist))
                naming_criteria_property_index_sublist = list(set(naming_criteria_property_index_sublist))
                        
                # sublist 안의 element들을 내림차순으로 정렬            
                naming_from_index_sublist.sort(reverse = True)
                naming_to_index_sublist.sort(reverse = True)
                naming_criteria_property_index_sublist.sort(reverse = True)
            
            # sublist를 합쳐 list로 완성
            naming_from_index_list.append(naming_from_index_sublist)
            naming_to_index_list.append(naming_to_index_sublist)
            naming_criteria_property_index_list.append(naming_criteria_property_index_sublist)        
    
        # 부재명 만들기, 기타 input sheet의 정보들 부재명에 따라 정리
        wall_info = wall.copy()  # input sheet에서 나온 properties
        wall_info.reset_index(drop=True, inplace=True)  # ?빼도되나?
    
        name_output = []  # new names
        property_output = []  # 이름 구분 조건에 따라 할당되는 properties를 새로운 부재명에 맞게 다시 정리한 output
        wall_info_output = []  # input sheet에서 나온 properties를 새로운 부재명에 맞게 다시 정리한 output
    
        count = 1000
        count_list = [] # 벽체이름을 오름차순으로 바꾸기 위한 index 만들기
    
        for i, j in zip(num_of_wall['Name'], num_of_wall['EA']):
            
            for k in range(1,int(j)+1):
    
                for current_wall_name, current_naming_from_index_list, current_naming_to_index_list, current_naming_criteria_property_index_list, current_wall_info_index\
                            in zip(wall['Name'], naming_from_index_list, naming_to_index_list, naming_criteria_property_index_list, wall_info.index):
    
                    if i == current_wall_name:
    
                        for p, q, r in zip(current_naming_from_index_list, current_naming_to_index_list, current_naming_criteria_property_index_list):
                            if p != q:
                                for s in range(p, q+1):
    
                                    count_list.append(count + s)
                                    
                                    name_output.append(current_wall_name + '_' + str(k) + '_' + str(story_name[s]))
                                    
                                    property_output.append(naming_criteria_property.iloc[:,-1][r])  # 각 이름에 맞게 property 할당 (index의 index 사용하였음)
                                    wall_info_output.append(wall_info.iloc[current_wall_info_index])
                                    
                            else:
                                count_list.append(count + q)
                                
                                name_output.append(current_wall_name + '_' + str(k) + '_' + str(story_name[q]))  # 시작과 끝층이 같으면 둘 중 한 층만 표기
                                
                                property_output.append(naming_criteria_property.iloc[:,-1][r])  # 각 이름에 맞게 property 할당 (index의 index 사용하였음)
                                wall_info_output.append(wall_info.iloc[current_wall_info_index])  
                                
                count += 1000
                
        wall_info_output = pd.DataFrame(wall_info_output)
        wall_info_output.reset_index(drop=True, inplace=True)  # 왜인지는 모르겠는데 index가 이상해져서..
    
        wall_info_output['Concrete Strength(CXX)'] = property_output  # 이름 구분 조건에 따른 property를 중간결과물에 재할당
    
        # 중간결과
        if (len(name_output) == 0) or (len(property_output) == 0):  # 구분 조건없이 을 경우는 wall_info를 바로 출력
            wall_ongoing = wall_info
        else:
            wall_ongoing = pd.concat([pd.Series(name_output, name='Name'), wall_info_output, pd.Series(count_list, name='Count')], axis = 1)  # 중간결과물 : 부재명 변경, 콘크리트 강도 추가, 부재명과 콘크리트 강도에 따른 properties
    
        wall_ongoing = wall_ongoing.sort_values(by=['Count']) # 층 오름차순으로 sort!(주석처리 for 내림차순)
        wall_ongoing.reset_index(inplace=True, drop=True)
    
        # 최종 sheet에 미리 넣을 수 있는 것들도 넣어놓기
        wall_output = wall_ongoing.iloc[:,[0,10,4,15,9,5,6,14,7,8,12,13]]  
    
        # 철근지름에 다시 D붙이기
        wall_output.loc[:,'Vertical Rebar(DXX)'] = 'D' + wall_output['Vertical Rebar(DXX)'].astype(str)
        wall_output.loc[:,'Horizontal Rebar(DXX)'] = 'D' + wall_output['Horizontal Rebar(DXX)'].astype(str)
        
        # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
        wall_output = wall_output.replace(np.nan, '', regex=True)

    #%% 2. Column
    #%% 불러온 Column 정보 정리
    if get_column == True:
        
        # 글자가 합쳐져 있을 경우 글자 나누기 - 층 (12F~15F, D10@300)
        new_story = column[['Story(from)', 'Story(to)']]
        new_story = new_story.fillna(method='ffill', axis=1)
              
        column['Story(from)'] = new_story['Story(from)']
        column['Story(to)'] = new_story['Story(to)']
    
        # 철근의 앞에붙은 D 떼어주기
        new_m_rebar = []
        new_h_rebar = []
    
        for i in column['Main Rebar(DXX)']:
            if isinstance(i, int):
                new_m_rebar.append(i)
            else:
                new_m_rebar.append(str_extract(i))
                
        for j in column['Hoop Rebar(DXX)']:
            if isinstance(j, int):
                new_h_rebar.append(j)
            else:
                new_h_rebar.append(str_extract(j))
                
        column['Main Rebar(DXX)'] = new_m_rebar
        column['Hoop Rebar(DXX)'] = new_h_rebar
    
        #%% 이름 구분 조건 load & 정리
    
        # 층 구분 조건에  story_name의 index 매칭시켜서 새로 열 만들기
        naming_criteria_1_index = []
        naming_criteria_2_index = []
    
        for i, j in zip(naming_criteria.iloc[:,5].dropna(), naming_criteria.iloc[:,6].dropna()):
            naming_criteria_1_index.append(pd.Index(story_name).get_loc(i))
            naming_criteria_2_index.append(pd.Index(story_name).get_loc(j))
    
        ### 구분 조건이 층 순서에 상관없이 작동되게 재정렬
        # 구분 조건에 해당하는 콘크리트 강도 재정렬
        naming_criteria_property = pd.concat([pd.Series(naming_criteria_1_index, name='Story(from) Index'), naming_criteria.iloc[:,7].dropna()], axis=1)
    
        naming_criteria_property['Story(from) Index'] = pd.Categorical(naming_criteria_property['Story(from) Index'], naming_criteria_1_index.sort())
        naming_criteria_property.sort_values('Story(from) Index', inplace=True)
        naming_criteria_property.reset_index(inplace=True)
    
        # 구분 조건 재정렬
        naming_criteria_1_index.sort()
        naming_criteria_2_index.sort()
    
        #%% 시작층, 끝층 정리
    
        naming_from_index = []
        naming_to_index = []
    
        for naming_from, naming_to in zip(column['Story(from)'], column['Story(to)']):
            if isinstance(naming_from, str) == False:
                naming_from = str(naming_from)
            if isinstance(naming_to, str) == False:
                naming_from = str(naming_from)
                
            naming_from_index.append(pd.Index(story_name).get_loc(naming_from))
            naming_to_index.append(pd.Index(story_name).get_loc(naming_to))
    
        #%%  층 이름을 etc의 이름 구분 조건에 맞게 나누어서 리스트로 정리
    
        naming_from_index_list = []
        naming_to_index_list = []
        naming_criteria_property_index_list = []
    
        for current_naming_from_index, current_naming_to_index in zip(naming_from_index, naming_to_index):  # 부재의 시작과 끝 층 loop
            naming_from_index_sublist = [current_naming_from_index]
            naming_to_index_sublist = [current_naming_to_index]
            naming_criteria_property_index_sublist = []
                
            for i, j, k in zip(naming_criteria_1_index, naming_criteria_2_index, naming_criteria_property.index):
                if (i >= current_naming_from_index) and (i <= current_naming_to_index):
                    naming_from_index_sublist.append(i)
                    naming_criteria_property_index_sublist.append(k)
                                
                    if (j >= current_naming_from_index) and (j <= current_naming_to_index):
                        naming_to_index_sublist.append(j)
                    else:
                        naming_to_index_sublist.append(i-1)
                        
                    if i != current_naming_from_index:
                        naming_criteria_property_index_sublist.append(k-1)
                                            
                elif (i < current_naming_from_index) and (j >= current_naming_to_index):
                    naming_criteria_property_index_sublist.append(k)
                    
                elif (i < current_naming_from_index) and (j <= current_naming_to_index):
                    naming_to_index_sublist.append(j)
                    
                else:
                    if max(naming_criteria_1_index) < current_naming_from_index:
                        naming_criteria_property_index_sublist.append(max(naming_criteria_property.index))
                        
                    elif min(naming_criteria_1_index) > current_naming_to_index:
                            naming_criteria_property_index_sublist.append(min(naming_criteria_property.index))
                    
                naming_from_index_sublist = list(set(naming_from_index_sublist))
                naming_to_index_sublist = list(set(naming_to_index_sublist))
                naming_criteria_property_index_sublist = list(set(naming_criteria_property_index_sublist))
                        
                # sublist 안의 element들을 내림차순으로 정렬            
                naming_from_index_sublist.sort(reverse = True)
                naming_to_index_sublist.sort(reverse = True)
                naming_criteria_property_index_sublist.sort(reverse = True)
            
            # sublist를 합쳐 list로 완성
            naming_from_index_list.append(naming_from_index_sublist)
            naming_to_index_list.append(naming_to_index_sublist)
            naming_criteria_property_index_list.append(naming_criteria_property_index_sublist)        
    
        # 부재명 만들기, 기타 input sheet의 정보들 부재명에 따라 정리
        column_info = column.copy()  # input sheet에서 나온 properties
        column_info.reset_index(drop=True, inplace=True)  # ?빼도되나?
    
        name_output = []  # new names
        property_output = []  # 이름 구분 조건에 따라 할당되는 properties를 새로운 부재명에 맞게 다시 정리한 output
        column_info_output = []  # input sheet에서 나온 properties를 새로운 부재명에 맞게 다시 정리한 output
    
        count = 1000
        count_list = [] # 벽체이름을 오름차순으로 바꾸기 위한 index 만들기
    
        for i, j in zip(num_of_column['Name'], num_of_column['EA']):
            
            for k in range(1,int(j)+1):
    
                for current_column_name, current_naming_from_index_list, current_naming_to_index_list, current_naming_criteria_property_index_list, current_column_info_index\
                            in zip(column['Name'], naming_from_index_list, naming_to_index_list, naming_criteria_property_index_list, column_info.index):
    
                    if i == current_column_name:
                    
                        for p, q, r in zip(current_naming_from_index_list, current_naming_to_index_list, current_naming_criteria_property_index_list):
                            if p != q:
                                for s in range(p, q+1):
    
                                    count_list.append(count + s)
                                    
                                    name_output.append(current_column_name + '_' + str(k) + '_' + str(story_name[s]))
                                    
                                    property_output.append(naming_criteria_property.iloc[:,-1][r])  # 각 이름에 맞게 property 할당 (index의 index 사용하였음)
                                    column_info_output.append(column_info.iloc[current_column_info_index])
                                    
                            else:
                                count_list.append(count + q)
                                
                                name_output.append(current_column_name + '_' + str(k) + '_' + str(story_name[q]))  # 시작과 끝층이 같으면 둘 중 한 층만 표기
                                
                                property_output.append(naming_criteria_property.iloc[:,-1][r])  # 각 이름에 맞게 property 할당 (index의 index 사용하였음)
                                column_info_output.append(column_info.iloc[current_column_info_index])  
                                
                count += 1000
                
        column_info_output = pd.DataFrame(column_info_output)
        column_info_output.reset_index(drop=True, inplace=True)  # 왜인지는 모르겠는데 index가 이상해져서..
    
        column_info_output['Concrete Strength(CXX)'] = property_output  # 이름 구분 조건에 따른 property를 중간결과물에 재할당
    
        # 중간결과
        if (len(name_output) == 0) or (len(property_output) == 0):  # 구분 조건없이 을 경우는 column_info를 바로 출력
            column_ongoing = column_info
        else:
            column_ongoing = pd.concat([pd.Series(name_output, name='Name'), column_info_output, pd.Series(count_list, name='Count')], axis = 1)  # 중간결과물 : 부재명 변경, 콘크리트 강도 추가, 부재명과 콘크리트 강도에 따른 properties
    
        column_ongoing = column_ongoing.sort_values(by=['Count'])
        column_ongoing.reset_index(inplace=True, drop=True)
    
        # 최종 sheet에 미리 넣을 수 있는 것들도 넣어놓기
        column_output = column_ongoing.iloc[:,[0,4,5,19,7,8,9,10,11,12,13,14,15,16,17,18]]  
    
        # 철근지름에 다시 D붙이기
        column_output.loc[:,'Main Rebar(DXX)'] = 'D' + column_output['Main Rebar(DXX)'].astype(str)
        column_output.loc[:,'Hoop Rebar(DXX)'] = 'D' + column_output['Hoop Rebar(DXX)'].astype(str)
        
        # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
        column_output = column_output.replace(np.nan, '', regex=True)
    
    #%% 3. Beam
    #%% 불러온 Beam 정보 정리
    if get_beam == True:
        
        # 글자가 합쳐져 있을 경우 글자 나누기 - 층 (12F~15F, D10@300)
        new_story = beam[['Story(from)', 'Story(to)']]
        new_story = new_story.fillna(method='ffill', axis=1)
              
        beam['Story(from)'] = new_story['Story(from)']
        beam['Story(to)'] = new_story['Story(to)']
    
        # 철근의 앞에붙은 D 떼어주기
        new_m_rebar = []
        new_s_rebar = []
    
        for i in beam['Main Rebar(DXX)']:
            if isinstance(i, int):
                new_m_rebar.append(i)
            else:
                new_m_rebar.append(str_extract(i))
                
        for j in beam['Stirrup Rebar(DXX)']:
            if isinstance(j, int):
                new_s_rebar.append(j)
            else:
                new_s_rebar.append(str_extract(j))
                
        beam['Main Rebar(DXX)'] = new_m_rebar
        beam['Stirrup Rebar(DXX)'] = new_s_rebar
    
        #%% 이름 구분 조건 load & 정리
    
        # 층 구분 조건에  story_name의 index 매칭시켜서 새로 열 만들기
        naming_criteria_1_index = []
        naming_criteria_2_index = []
    
        for i, j in zip(naming_criteria.iloc[:,8].dropna(), naming_criteria.iloc[:,9].dropna()):
            naming_criteria_1_index.append(pd.Index(story_name).get_loc(i))
            naming_criteria_2_index.append(pd.Index(story_name).get_loc(j))
    
        ### 구분 조건이 층 순서에 상관없이 작동되게 재정렬
        # 구분 조건에 해당하는 콘크리트 강도 재정렬
        naming_criteria_property = pd.concat([pd.Series(naming_criteria_1_index, name='Story(from) Index'), naming_criteria.iloc[:,10].dropna()], axis=1)
    
        naming_criteria_property['Story(from) Index'] = pd.Categorical(naming_criteria_property['Story(from) Index'], naming_criteria_1_index.sort())
        naming_criteria_property.sort_values('Story(from) Index', inplace=True)
        naming_criteria_property.reset_index(inplace=True)
    
        # 구분 조건 재정렬
        naming_criteria_1_index.sort()
        naming_criteria_2_index.sort()
    
        #%% 시작층, 끝층 정리
    
        naming_from_index = []
        naming_to_index = []
    
        for naming_from, naming_to in zip(beam['Story(from)'], beam['Story(to)']):
            if isinstance(naming_from, str) == False:
                naming_from = str(naming_from)
            if isinstance(naming_to, str) == False:
                naming_from = str(naming_from)
                
            naming_from_index.append(pd.Index(story_name).get_loc(naming_from))
            naming_to_index.append(pd.Index(story_name).get_loc(naming_to))
    
        #%%  층 이름을 etc의 이름 구분 조건에 맞게 나누어서 리스트로 정리
    
        naming_from_index_list = []
        naming_to_index_list = []
        naming_criteria_property_index_list = []
    
        for current_naming_from_index, current_naming_to_index in zip(naming_from_index, naming_to_index):  # 부재의 시작과 끝 층 loop
            naming_from_index_sublist = [current_naming_from_index]
            naming_to_index_sublist = [current_naming_to_index]
            naming_criteria_property_index_sublist = []
                
            for i, j, k in zip(naming_criteria_1_index, naming_criteria_2_index, naming_criteria_property.index):
                if (i >= current_naming_from_index) and (i <= current_naming_to_index):
                    naming_from_index_sublist.append(i)
                    naming_criteria_property_index_sublist.append(k)
                                
                    if (j >= current_naming_from_index) and (j <= current_naming_to_index):
                        naming_to_index_sublist.append(j)
                    else:
                        naming_to_index_sublist.append(i-1)
                        
                    if i != current_naming_from_index:
                        naming_criteria_property_index_sublist.append(k-1)
                                            
                elif (i < current_naming_from_index) and (j >= current_naming_to_index):
                    naming_criteria_property_index_sublist.append(k)
                    
                elif (i < current_naming_from_index) and (j <= current_naming_to_index):
                    naming_to_index_sublist.append(j)
                    
                else:
                    if max(naming_criteria_1_index) < current_naming_from_index:
                        naming_criteria_property_index_sublist.append(max(naming_criteria_property.index))
                        
                    elif min(naming_criteria_1_index) > current_naming_to_index:
                            naming_criteria_property_index_sublist.append(min(naming_criteria_property.index))
                    
                naming_from_index_sublist = list(set(naming_from_index_sublist))
                naming_to_index_sublist = list(set(naming_to_index_sublist))
                naming_criteria_property_index_sublist = list(set(naming_criteria_property_index_sublist))
                        
                # sublist 안의 element들을 내림차순으로 정렬            
                naming_from_index_sublist.sort(reverse = True)
                naming_to_index_sublist.sort(reverse = True)
                naming_criteria_property_index_sublist.sort(reverse = True)
            
            # sublist를 합쳐 list로 완성
            naming_from_index_list.append(naming_from_index_sublist)
            naming_to_index_list.append(naming_to_index_sublist)
            naming_criteria_property_index_list.append(naming_criteria_property_index_sublist)        
    
        # 부재명 만들기, 기타 input sheet의 정보들 부재명에 따라 정리
        beam_info = beam.copy()  # input sheet에서 나온 properties
        beam_info.reset_index(drop=True, inplace=True)  # ?빼도되나?
    
        name_output = []  # new names
        property_output = []  # 이름 구분 조건에 따라 할당되는 properties를 새로운 부재명에 맞게 다시 정리한 output
        beam_info_output = []  # input sheet에서 나온 properties를 새로운 부재명에 맞게 다시 정리한 output
    
        count = 1000
        count_list = [] # 벽체이름을 오름차순으로 바꾸기 위한 index 만들기
    
        for i, j in zip(num_of_beam['Name'], num_of_beam['EA']):
            
            for k in range(1,int(j)+1):
    
                for current_beam_name, current_naming_from_index_list, current_naming_to_index_list, current_naming_criteria_property_index_list, current_beam_info_index\
                            in zip(beam['Name'], naming_from_index_list, naming_to_index_list, naming_criteria_property_index_list, beam_info.index):
    
                    if i == current_beam_name:
                        
                        
                        
                        for p, q, r in zip(current_naming_from_index_list, current_naming_to_index_list, current_naming_criteria_property_index_list):
                            if p != q:
                                for s in range(p, q+1):
    
                                    count_list.append(count + s)
                                    
                                    name_output.append(current_beam_name + '_' + str(k) + '_' + str(story_name[s]))
                                    
                                    property_output.append(naming_criteria_property.iloc[:,-1][r])  # 각 이름에 맞게 property 할당 (index의 index 사용하였음)
                                    beam_info_output.append(beam_info.iloc[current_beam_info_index])
                                    
                            else:
                                count_list.append(count + q)
                                
                                name_output.append(current_beam_name + '_' + str(k) + '_' + str(story_name[q]))  # 시작과 끝층이 같으면 둘 중 한 층만 표기
                                
                                property_output.append(naming_criteria_property.iloc[:,-1][r])  # 각 이름에 맞게 property 할당 (index의 index 사용하였음)
                                beam_info_output.append(beam_info.iloc[current_beam_info_index])  
                                
                count += 1000
                
        beam_info_output = pd.DataFrame(beam_info_output)
        beam_info_output.reset_index(drop=True, inplace=True)  # 왜인지는 모르겠는데 index가 이상해져서..
    
        beam_info_output['Concrete Strength(CXX)'] = property_output  # 이름 구분 조건에 따른 property를 중간결과물에 재할당
    
        # 중간결과
        if (len(name_output) == 0) or (len(property_output) == 0):  # 구분 조건없이 을 경우는 beam_info를 바로 출력
            beam_ongoing = beam_info
        else:
            beam_ongoing = pd.concat([pd.Series(name_output, name='Name'), beam_info_output, pd.Series(count_list, name='Count')], axis = 1)  # 중간결과물 : 부재명 변경, 콘크리트 강도 추가, 부재명과 콘크리트 강도에 따른 properties
    
        beam_ongoing = beam_ongoing.sort_values(by=['Count'])
        beam_ongoing.reset_index(inplace=True, drop=True)
    
        # 최종 sheet에 미리 넣을 수 있는 것들도 넣어놓기
        beam_output = beam_ongoing.iloc[:,[0,4,5,6,21,22,8,9,10,11,12,13,14,15,16,17,18,19,20]]  
    
        # 철근지름에 다시 D붙이기
        beam_output.loc[:,'Main Rebar(DXX)'] = 'D' + beam_output['Main Rebar(DXX)'].astype(str)
        beam_output.loc[:,'Stirrup Rebar(DXX)'] = 'D' + beam_output['Stirrup Rebar(DXX)'].astype(str)
        
        # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
        beam_output = beam_output.replace(np.nan, '', regex=True)

    #%% Printout
    # Using win32com...

    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게

    wb = excel.Workbooks.Open(input_xlsx_path)
    ws_beam = wb.Sheets('Output_C.Beam Properties')
    ws_column = wb.Sheets('Output_G.Column Properties')
    ws_wall = wb.Sheets('Output_Wall Properties')

    startrow, startcol = 5, 1

    if get_beam == True:        
        ws_beam.Range(ws_beam.Cells(startrow, startcol),\
                      ws_beam.Cells(startrow + beam_output.shape[0]-1,\
                                    startcol + beam_output.shape[1]-1)).Value\
        = list(beam_output.itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능
        
    if get_column == True:        
        ws_column.Range(ws_column.Cells(startrow, startcol),\
                        ws_column.Cells(startrow + column_output.shape[0]-1,\
                                        startcol + column_output.shape[1]-1)).Value\
        = list(column_output.itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능

    if get_wall == True:    
        ws_wall.Range(ws_wall.Cells(startrow, startcol),\
                      ws_wall.Cells(startrow + wall_output.shape[0]-1,\
                                    startcol + wall_output.shape[1]-1)).Value\
        = list(wall_output.itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능

    wb.Save()
    # wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application 
    
#%% Convert Column Nu

def convert_property_col_Nu(input_xlsx_path, result_path, result_xlsx='Analysis Result'
                            , g_col_group_name = 'G.Column'):
    '''
    
    User가 입력한 부재 정보들을 Perform-3D에 입력할 수 있는 형식으로 변환하여 Data Conversion 엑셀파일의 Output_Properties 시트에 작성.
    
    Parameters
    ----------
    input_path : str
                 Data Conversion 엑셀 파일의 경로.
                 
    input_xlsx : str
                 Data Conversion 엑셀 파일의 이름. 확장자명(.xlsx)까지 기입해줘야한다. 하나의 파일만 불러온다.

    get_beam : bool, optional, default=True
               True = C.Beam의 정보를 Perform-3D 입력용 정보로 변환함.
               False = C.Beam의 정보를 변환하지 않음.
               
    get_column : bool, optional, default=True
                 True = G.Column의 정보를 Perform-3D 입력용 정보로 변환함.
                 False = G.Column의 정보를 변환하지 않음.
               
    get_wall : bool, optional, default=True
               True = Wall의 정보를 Perform-3D 입력용 정보로 변환함.
               False = Wall의 정보를 변환하지 않음.

    Returns
    --------       
    beam_output : pandas.core.frame.DataFrame or None
                  C.Beam Properties의 정보를 Perform-3D 입력용으로 변환한 정보.
                  Output_C.Beam Properties 시트에 입력됨.
                     
    wall_output : pandas.core.frame.DataFrame or None
                  Wall Properties의 정보를 Perform-3D 입력용으로 변환한 정보.
                  Output_Wall Properties 시트에 입력됨.   
                  
    Raises
    -------
    
    '''    
    #%% 파일 load
    
    pd.options.mode.chained_assignment = None # SettingWithCopyWarning 안뜨게 하기
    # UserWarning: openpyxl 안뜨게 하기
    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
    
    input_data_raw = pd.ExcelFile(input_xlsx_path)
    input_data_sheets = pd.read_excel(input_data_raw\
                                      , ['G.Column Properties', 'Story Data'
                                         , 'ETC', 'Naming'], skiprows=3)
    input_data_raw.close()

    # Column 정보 load
    column = input_data_sheets['G.Column Properties'].iloc[:,0:18]
    column.columns = ['Name', 'Story(from)', 'Story(to)', 'b(mm)', 'h(mm)'
                      , 'Cover Thickness(mm)', '내진상세 여부', 'Type(Main)'
                      , 'Main Rebar(DXX)', 'Type(Hoop)', 'Hoop Rebar(DXX)'
                      , 'EA(Layer1)', 'Row(Layer1)', 'EA(Layer2)', 'Row(Layer2)'
                      , 'EA(Hoop_X)', 'EA(Hoop_Y)', 'Spacing(Hoop)']

    column = column.dropna(axis=0, how='all')
    column.reset_index(inplace=True, drop=True)
    column = column.fillna(method='ffill')

    # 구분 조건 load
    naming_criteria = input_data_sheets['ETC']

    # Story 정보 load
    story_info = input_data_sheets['Story Data'].iloc[:,0:3]
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    story_name = story_info.loc[:, 'Story Name']
    story_name = story_name[::-1]  # 층 이름 재배열
    story_name.reset_index(drop=True, inplace=True)

    # 벽체,기둥,보 개수 load
    num_of_elem = input_data_sheets['Naming']
    
    num_of_beam = num_of_elem.iloc[:,[0,3]]
    num_of_wall = num_of_elem.iloc[:,[8,11]]
    num_of_column = num_of_elem.iloc[:,[4,7]]
    
    num_of_beam = num_of_beam.dropna(axis=0)
    num_of_wall = num_of_wall.dropna(axis=0)
    num_of_column = num_of_column.dropna(axis=0)

    num_of_beam.columns = ['Name', 'EA']
    num_of_wall.columns = ['Name', 'EA']
    num_of_column.columns = ['Name', 'EA']

    #%% 부재 이름 설정할 때 필요한 함수들

    # 층 나누는 함수 (12F~15F)
    def str_div(temp_list):
        first = []
        second = []
        
        for i in temp_list:
            if '~' in i:
                first.append(i.split('~')[0])
                second.append(i.split('~')[1])
            elif '-' in i:
                second.append(i.split('-')[0])
                first.append(i.split('-')[1])
            else:
                first.append(i)
                second.append(i)
        
        first = pd.Series(first).str.strip()
        second = pd.Series(second).str.strip()
        
        return first, second

    # 층, 철근 나누는 함수 (12F~15F, D10@300)
    def rebar_div(temp_list1, temp_list2):
        first = []
        second = []
        third = []
        
        for i, j in zip(temp_list1, temp_list2):
            if isinstance(i, str) : # string인 경우
                if '@' in i:
                    first.append(i.split('@')[0].strip())
                    second.append(i.split('@')[1])
                    third.append(np.nan)
                elif '-' in i:
                    third.append(i.split('-')[0])
                    first.append(i.split('-')[1].strip())
                    second.append(np.nan)
                else: 
                    first.append(i.strip())
                    second.append(j)
                    third.append(np.nan)
            else: # string 아닌 경우
                first.append(i)
                second.append(j)
                third.append(np.nan)

        return first, second, third

    # 철근 지름 앞의 D 떼주는 함수 (D10...)
    def str_extract(sth_str):
        result = int(re.findall(r'[0-9]+', sth_str)[0])
        
        return result

    #%% 데이터베이스
    steel_geometry_database = naming_criteria.iloc[:,[0,1,2]].dropna()
    steel_geometry_database.columns = ['Name', 'Diameter(mm)', 'Area(mm^2)']

    new_steel_geometry_name = []

    for i in steel_geometry_database['Name']:
        if isinstance(i, int):
            new_steel_geometry_name.append(i)
        else:
            new_steel_geometry_name.append(str_extract(i))

    steel_geometry_database['Name'] = new_steel_geometry_name

    #%% 2. Column
    #%% 불러온 Column 정보 정리
        
    # 글자가 합쳐져 있을 경우 글자 나누기 (12F~15F, D10@300)
    # 층 나누기

    if column['Story(to)'].isnull().any() == True:
        column['Story(to)'] = str_div(column['Story(from)'])[1]
        column['Story(from)'] = str_div(column['Story(from)'])[0]
    else: pass

    # 철근의 앞에붙은 D 떼어주기
    new_m_rebar = []
    new_h_rebar = []

    for i in column['Main Rebar(DXX)']:
        if isinstance(i, int):
            new_m_rebar.append(i)
        else:
            new_m_rebar.append(str_extract(i))
            
    for j in column['Hoop Rebar(DXX)']:
        if isinstance(j, int):
            new_h_rebar.append(j)
        else:
            new_h_rebar.append(str_extract(j))
            
    column['Main Rebar(DXX)'] = new_m_rebar
    column['Hoop Rebar(DXX)'] = new_h_rebar

    #%% 이름 구분 조건 load & 정리

    # 층 구분 조건에  story_name의 index 매칭시켜서 새로 열 만들기
    naming_criteria_1_index = []
    naming_criteria_2_index = []

    for i, j in zip(naming_criteria.iloc[:,5].dropna(), naming_criteria.iloc[:,6].dropna()):
        naming_criteria_1_index.append(pd.Index(story_name).get_loc(i))
        naming_criteria_2_index.append(pd.Index(story_name).get_loc(j))

    ### 구분 조건이 층 순서에 상관없이 작동되게 재정렬
    # 구분 조건에 해당하는 콘크리트 강도 재정렬
    naming_criteria_property = pd.concat([pd.Series(naming_criteria_1_index, name='Story(from) Index'), naming_criteria.iloc[:,7].dropna()], axis=1)

    naming_criteria_property['Story(from) Index'] = pd.Categorical(naming_criteria_property['Story(from) Index'], naming_criteria_1_index.sort())
    naming_criteria_property.sort_values('Story(from) Index', inplace=True)
    naming_criteria_property.reset_index(inplace=True)

    # 구분 조건 재정렬
    naming_criteria_1_index.sort()
    naming_criteria_2_index.sort()

    #%% 시작층, 끝층 정리

    naming_from_index = []
    naming_to_index = []

    for naming_from, naming_to in zip(column['Story(from)'], column['Story(to)']):
        if isinstance(naming_from, str) == False:
            naming_from = str(naming_from)
        if isinstance(naming_to, str) == False:
            naming_from = str(naming_from)
            
        naming_from_index.append(pd.Index(story_name).get_loc(naming_from))
        naming_to_index.append(pd.Index(story_name).get_loc(naming_to))

    #%%  층 이름을 etc의 이름 구분 조건에 맞게 나누어서 리스트로 정리

    naming_from_index_list = []
    naming_to_index_list = []
    naming_criteria_property_index_list = []

    for current_naming_from_index, current_naming_to_index in zip(naming_from_index, naming_to_index):  # 부재의 시작과 끝 층 loop
        naming_from_index_sublist = [current_naming_from_index]
        naming_to_index_sublist = [current_naming_to_index]
        naming_criteria_property_index_sublist = []
            
        for i, j, k in zip(naming_criteria_1_index, naming_criteria_2_index, naming_criteria_property.index):
            if (i >= current_naming_from_index) and (i <= current_naming_to_index):
                naming_from_index_sublist.append(i)
                naming_criteria_property_index_sublist.append(k)
                            
                if (j >= current_naming_from_index) and (j <= current_naming_to_index):
                    naming_to_index_sublist.append(j)
                else:
                    naming_to_index_sublist.append(i-1)
                    
                if i != current_naming_from_index:
                    naming_criteria_property_index_sublist.append(k-1)
                                        
            elif (i < current_naming_from_index) and (j >= current_naming_to_index):
                naming_criteria_property_index_sublist.append(k)
                
            elif (i < current_naming_from_index) and (j <= current_naming_to_index):
                naming_to_index_sublist.append(j)
                
            else:
                if max(naming_criteria_1_index) < current_naming_from_index:
                    naming_criteria_property_index_sublist.append(max(naming_criteria_property.index))
                    
                elif min(naming_criteria_1_index) > current_naming_to_index:
                        naming_criteria_property_index_sublist.append(min(naming_criteria_property.index))
                
            naming_from_index_sublist = list(set(naming_from_index_sublist))
            naming_to_index_sublist = list(set(naming_to_index_sublist))
            naming_criteria_property_index_sublist = list(set(naming_criteria_property_index_sublist))
                    
            # sublist 안의 element들을 내림차순으로 정렬            
            naming_from_index_sublist.sort(reverse = True)
            naming_to_index_sublist.sort(reverse = True)
            naming_criteria_property_index_sublist.sort(reverse = True)
        
        # sublist를 합쳐 list로 완성
        naming_from_index_list.append(naming_from_index_sublist)
        naming_to_index_list.append(naming_to_index_sublist)
        naming_criteria_property_index_list.append(naming_criteria_property_index_sublist)        

    # 부재명 만들기, 기타 input sheet의 정보들 부재명에 따라 정리
    column_info = column.copy()  # input sheet에서 나온 properties
    column_info.reset_index(drop=True, inplace=True)  # ?빼도되나?

    name_output = []  # new names
    property_output = []  # 이름 구분 조건에 따라 할당되는 properties를 새로운 부재명에 맞게 다시 정리한 output
    column_info_output = []  # input sheet에서 나온 properties를 새로운 부재명에 맞게 다시 정리한 output

    count = 1000
    count_list = [] # 벽체이름을 오름차순으로 바꾸기 위한 index 만들기

    for i, j in zip(num_of_column['Name'], num_of_column['EA']):
        
        for k in range(1,int(j)+1):

            for current_column_name, current_naming_from_index_list, current_naming_to_index_list, current_naming_criteria_property_index_list, current_column_info_index\
                        in zip(column['Name'], naming_from_index_list, naming_to_index_list, naming_criteria_property_index_list, column_info.index):

                if i == current_column_name:
                
                    for p, q, r in zip(current_naming_from_index_list, current_naming_to_index_list, current_naming_criteria_property_index_list):
                        if p != q:
                            for s in range(p, q+1):

                                count_list.append(count + s)
                                
                                name_output.append(current_column_name + '_' + str(k) + '_' + str(story_name[s]))
                                
                                property_output.append(naming_criteria_property.iloc[:,-1][r])  # 각 이름에 맞게 property 할당 (index의 index 사용하였음)
                                column_info_output.append(column_info.iloc[current_column_info_index])
                                
                        else:
                            count_list.append(count + q)
                            
                            name_output.append(current_column_name + '_' + str(k) + '_' + str(story_name[q]))  # 시작과 끝층이 같으면 둘 중 한 층만 표기
                            
                            property_output.append(naming_criteria_property.iloc[:,-1][r])  # 각 이름에 맞게 property 할당 (index의 index 사용하였음)
                            column_info_output.append(column_info.iloc[current_column_info_index])  
                            
            count += 1000
            
    column_info_output = pd.DataFrame(column_info_output)
    column_info_output.reset_index(drop=True, inplace=True)  # 왜인지는 모르겠는데 index가 이상해져서..

    column_info_output['Concrete Strength(CXX)'] = property_output  # 이름 구분 조건에 따른 property를 중간결과물에 재할당

    # 중간결과
    if (len(name_output) == 0) or (len(property_output) == 0):  # 구분 조건없이 을 경우는 column_info를 바로 출력
        column_ongoing = column_info
    else:
        column_ongoing = pd.concat([pd.Series(name_output, name='Name'), column_info_output, pd.Series(count_list, name='Count')], axis = 1)  # 중간결과물 : 부재명 변경, 콘크리트 강도 추가, 부재명과 콘크리트 강도에 따른 properties

    column_ongoing = column_ongoing.sort_values(by=['Count'])
    column_ongoing.reset_index(inplace=True, drop=True)

#%% Nu값 불러오기
    to_load_list = []
    file_names = os.listdir(result_path)
    for file_name in file_names:
        if (result_xlsx in file_name) and ('~$' not in file_name):
            to_load_list.append(file_name)

    P_data = pd.DataFrame()

    for i in to_load_list:
        result_data_raw = pd.ExcelFile(result_path + '\\' + i)
        result_data_sheets = pd.read_excel(result_data_raw, ['Element Data - Frame Types', 'Frame Results - End Forces'], skiprows=[0,2])
        
        column_name_to_slice = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'P J-End']
        P_data_temp = result_data_sheets['Frame Results - End Forces'].loc[:,column_name_to_slice]
        P_data = pd.concat([P_data, P_data_temp])
        
    column_name = result_data_sheets['Element Data - Frame Types'].loc[:,['Element Name', 'Property Name']]
    
#%% 지진파 이름 list 만들기
    load_name_list = []
    for i in P_data['Load Case'].drop_duplicates():
        new_i = i.split('+')[1]
        new_i = new_i.strip()
        load_name_list.append(new_i)

    gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
    seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]

    seismic_load_name_list.sort()
    
    DE_load_name_list = [x for x in load_name_list if 'DE' in x]
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]

#%% Nu값 뽑기
    # 축력 불러와서 Grouping
    P_data = P_data[P_data['Group Name'] == g_col_group_name]    
    P_data = P_data[P_data['Step Type'].str.contains('Min')]
    P_data = P_data[P_data['Load Case'].str.contains(gravity_load_name[0])]
    P_data = pd.merge(P_data, column_name, how='left')
    P_data.reset_index(inplace=True, drop=True)

    # 부호 반대로
    P_data['P J-End'] = -P_data['P J-End']

    # result
    P = P_data[['Property Name', 'P J-End']]
    P.columns = ['Name', 'Nu(kN)']
    
#%% Column Output 출력

    # 최종 sheet에 미리 넣을 수 있는 것들도 넣어놓기
    column_output = column_ongoing.iloc[:,[0,4,5,19,7,8,9,10,11,12,13,14,15,16,17,18]]  

    # 철근지름에 다시 D붙이기
    column_output.loc[:,'Main Rebar(DXX)'] = 'D' + column_output['Main Rebar(DXX)'].astype(str)
    column_output.loc[:,'Hoop Rebar(DXX)'] = 'D' + column_output['Hoop Rebar(DXX)'].astype(str)
    
    # Nu merge
    column_output = pd.merge(column_output, P, how='left')
    
    # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
    column_output = column_output.replace(np.nan, '', regex=True)

    #%% Printout
    # Using win32com...

    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게

    wb = excel.Workbooks.Open(input_xlsx_path)
    ws_column = wb.Sheets('Output_G.Column Properties')

    startrow, startcol = 5, 1

    ws_column.Range(ws_column.Cells(startrow, startcol),\
                    ws_column.Cells(startrow + column_output.shape[0]-1,\
                                    startcol + column_output.shape[1]-1)).Value\
    = list(column_output.itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능

    wb.Save()
    # wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application 
    
#%% Property Assign Macro (Wall)

# def property_assign_macro()