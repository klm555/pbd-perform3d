import pandas as pd
import numpy as np
import win32com.client

#%% Node, Element, Mass, Load Import

def import_midas(input_path, input_xlsx, DL_name='DL', LL_name='LL'\
                 , import_node=True, import_DL=True, import_LL=True\
                 , import_mass=True, **kwargs):
    
    defaultkwargs = {'import_beam':True, 'import_column':True,\
                     'import_wall':True, 'import_plate':True\
                     'import_SWR_gage':True, 'import_AS_gage':True}
        
    kwargs = {**defaultkwargs, **kwargs}
    '''
    
    Midas GEN 모델을 Perform-3D로 import할 수 있는 파일 형식(.csv)으로 변환.
    
    Parameters
    ----------
    input_path : str
                 Data Conversion 엑셀 파일의 경로.
                 
    input_xlsx : str
                 Data Conversion 엑셀 파일의 이름. 확장자명(.xlsx)까지 기입해줘야한다. 하나의 파일만 불러온다.
    
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

    import_SWR_gage : bool, optional, default=True
                      True = Wall Rotation Gage의 csv파일을 생성함.
                      False = Wall Rotation Gage의 csv파일을 생성 안 함.
                   
    import_AS_gage : bool, optional, default=True
                     True = Wall Axial Strain Gage의 csv파일을 생성함.
                     False = Wall Axial Strain Gage의 csv파일을 생성 안 함.                   
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
    output_csv_dir = input_path # 또는 '경로'
    
    node_DL_merged_csv = 'DL.csv'
    node_LL_merged_csv = 'LL.csv'
    mass_csv = 'Mass.csv'
    node_csv = 'Node.csv'
    beam_csv = 'Beam.csv'
    column_csv = 'Column.csv'
    wall_csv = 'Wall.csv'
    plate_csv = 'Plate.csv'
    wall_gage_csv = 'Shear Wall Rotation Gage.csv'
    as_gage_csv = 'Axial Strain Gage.csv'
    
    #%% Nodal Load 뽑기
    
    # Node 정보 load
    node = pd.read_excel(input_path + '\\' + input_xlsx, sheet_name = input_xlsx_sheet, skiprows = 3, index_col = 0)  # Node 열을 인덱스로 지정
    node.columns = ['X(mm)', 'Y(mm)', 'Z(mm)']
    
    if (import_DL == True) or (import_LL == True):
    
        # Nodal Load 정보 load
        nodal_load = pd.read_excel(input_path+'\\'+input_xlsx, sheet_name = nodal_load_raw_xlsx_sheet, skiprows = 3, index_col = 0)
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
        mass = pd.read_excel(input_path+'\\'+input_xlsx, sheet_name = mass_raw_xlsx_sheet, skiprows = 3)
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
        
        # Mass의 nodes(좌표) 추가        
        node_mass_considered = pd.concat([node, mass2.iloc[:,[0,1,2]]])
    
        # Node 결과값을 csv로 변환
        if import_node == True:
            node_mass_considered.to_csv(output_csv_dir+'\\'+node_csv, mode='w', index=False) # Import할 Mass의 좌표를 포함한 모든 좌표를 csv로 출력함
       
        else:
            mass2.iloc[:,[0,1,2]].to_csv(output_csv_dir+'\\'+node_csv, mode='w', index=False) # Import할 Mass의 좌표만 csv로 출력함
            
    else:
        # Node 결과값을 csv로 변환
        node.to_csv(output_csv_dir+'\\'+node_csv, mode='w', index=False)
        
    #%% Beam Element 뽑기
    
    # Index로 지정되어있던 Node 번호를 다시 reset
    node.index.name = 'Node'
    node.reset_index(inplace=True)
    
    # Element 정보 load
    element = pd.read_excel(input_path+'\\'+input_xlsx, sheet_name = element_raw_xlsx_sheet, skiprows = [0,2,3])
    
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
        wall_node_coord_pos = wall_node_coord[wall_node_coord['X_2(mm)'] > wall_node_coord['X_1(mm)']]
        wall_node_coord_neg = wall_node_coord[wall_node_coord['X_2(mm)'] < wall_node_coord['X_1(mm)']]
        wall_node_coord_zero = wall_node_coord[wall_node_coord['X_2(mm)'] == wall_node_coord['X_1(mm)']]
        
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
    
    if import_AS_gage == True:
        
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
        wall_gage_pos = wall_gage[wall_gage['X(mm)2'] > wall_gage['X(mm)']]
        wall_gage_neg = wall_gage[wall_gage['X(mm)2'] < wall_gage['X(mm)']]
        wall_gage_zero = wall_gage[wall_gage['X(mm)2'] == wall_gage['X(mm)']]
        
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
        duplicates_list = []
        gage_node_data_list = []
        for idx, gage_node_data in wall_gage_sorted.groupby(['Wall ID', 'Z(mm)'])['Node List']:
            
            # 같은 Wall ID를 가지지만 붙어있지 않은 벽체 구별해내기
            gage_node_list_flat = [i for gage_node_list in gage_node_data for i in gage_node_list]
            duplicates = list(set([i for i in gage_node_list_flat if gage_node_list_flat.count(i) > 1])) # 위의 리스트에서 겹치는 부재들 remove
            
            duplicates_list.append(duplicates)
            gage_node_data_list.append(gage_node_data)
        
        # 같은 wall mark를 갖고, 겹치는 Node가 있는 벽체, 없는 벽체를 구분
        gage_node_list = []
        for gage_node_data, duplicates in zip(gage_node_data_list, duplicates_list):
            
            if len(gage_node_data) > 1: # 같은 Index(Wall ID, Z(mm))에 2개 이상의 벽체가 Assign 되어있을때    
                gage_node_sublist = []
                for gage_node_subdata in gage_node_data:
                                
                    if any(i in gage_node_subdata for i in duplicates): # duplicates_list의 중복되는 node가 하나라도 포함되어있는 경우
                        gage_node_sublist.append(gage_node_subdata)
                    
                    else:
                        gage_node_list.append([gage_node_subdata])
                            
                    gage_node_list.append(gage_node_sublist)
                
            else:
                gage_node_list.append(gage_node_data.tolist())
        
        # Node List 생성 (Node 번호순으로 재배열)        
        gage_node_list_zip = []
        for gage_node_sublist in gage_node_list:
            if len(gage_node_sublist) > 1:
                
                # 같은 Index(Wall ID, Z(mm))인 부재들의 Nodes를 Index에 맞춰 재배열한 list 만들기
                gage_node_sublist_zip = [list(i) for i in zip(*gage_node_sublist)]
                gage_node_list_zip.append(gage_node_sublist_zip)
    
            elif len(gage_node_sublist) == 1:
                gage_node_list_zip.append(gage_node_sublist)
                
        # 위에서 재배열한 list를 flatten
        gage_node_list_flat = []
        for gage_node_sublist_zip in gage_node_list_zip:
            if len(gage_node_sublist_zip) > 1:
                gage_node_sublist_flat = [i for gage_node_sublist_sublist_zip in gage_node_sublist_zip for i in gage_node_sublist_sublist_zip]
                gage_node_list_flat.append(gage_node_sublist_flat)
                
            elif len(gage_node_sublist_zip) == 1:
                gage_node_list_flat.append(gage_node_sublist_zip[0])
                
        # 중복되는 list 제거
        gage_node_list_flat_set = set(map(tuple, gage_node_list_flat)) # list -> tuple (to make it hashable)
        gage_node_list_flat_reduced = map(list, gage_node_list_flat_set) # tuple -> list  
    
        # sublist에서 중복되는 element 제거
        gage_node_list_flat_set_reduced = []
        for i in gage_node_list_flat_set:
            temp = [x for x in i if i.count(x) == 1]
            gage_node_list_flat_set_reduced.append(temp)
    
        # Node 번호에 맞는 좌표 매칭 후 출력
        gage_node_coord = pd.DataFrame(gage_node_list_flat_set_reduced)
        gage_node_coord.columns = ['Node1', 'Node2', 'Node3', 'Node4']
        
        # SWR_gage_node를 as_gage_node로 나누고 재배열
        as_gage_node_coord_1 = gage_node_coord[['Node1', 'Node4']]
        as_gage_node_coord_2 = gage_node_coord[['Node2', 'Node3']]
        
        as_gage_node_coord_1.columns = ['Node1', 'Node2']
        as_gage_node_coord_2.columns = ['Node1', 'Node2']
        as_gage_node_coord = pd.concat([as_gage_node_coord_1, as_gage_node_coord_2])
        
        as_gage_node_coord.drop_duplicates(inplace=True)
        
        # Merge로 Node 번호에 맞는 좌표를 결합
        as_gage_node_coord = pd.merge(as_gage_node_coord, node, how='left', left_on='Node1', right_on='Node', suffixes=(None, '1'))
        as_gage_node_coord = pd.merge(as_gage_node_coord, node, how='left', left_on='Node2', right_on='Node', suffixes=(None, '2'))
        
        as_gage_node_coord = as_gage_node_coord.iloc[:,[3,4,5,7,8,9]]
        
        # Gage Element 결과값을 csv로 변환
        as_gage_node_coord.to_csv(output_csv_dir+'\\'+as_gage_csv, mode='w', index=False)

    #%% Shear Wall Gage 뽑기
    
    if import_SWR_gage == True:
        
        # Story Info data 불러오기
        story_info = pd.read_excel(input_path + '\\' + input_xlsx, sheet_name=story_info_xlsx_sheet\
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
        # X 좌표가 더 작은 노드를 i-node로!
        wall_gage_pos = wall_gage[wall_gage['X(mm)2'] > wall_gage['X(mm)']]
        wall_gage_neg = wall_gage[wall_gage['X(mm)2'] < wall_gage['X(mm)']]
        wall_gage_zero = wall_gage[wall_gage['X(mm)2'] == wall_gage['X(mm)']]
        
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
        
        ### SWR gage가 분할층에서 나눠지지 않게 만들기 
        # 분할층 노드가 포함되지 않은 부재 slice
        wall_gage_no_divide = wall_gage[(wall_gage['Z(mm)'].isin(story_info['Level']))\
                                        & (wall_gage['Z(mm)3'].isin(story_info['Level']))]
        
        # 분할층 노드가 상부에만(k,l-node) 포함되는 부재 slice
        wall_gage_divide = wall_gage[~wall_gage['Z(mm)3'].isin(story_info['Level'])]
        
        # wall_gage_divide 노드들의 상부 노드(k,l-node)의 z좌표를 다음 측으로 격상
        next_level_list = []
        for i in wall_gage_divide['Z(mm)3']:
            level_bigger = story_info['Level'][story_info['Level']-i >= 0]
            next_level = level_bigger.sort_values(ignore_index=True)[0]

            next_level_list.append(next_level)
        
        pd.options.mode.chained_assignment = None # SettingWithCopyWarning 안뜨게 하기

        wall_gage_divide.loc[:,'Z(mm)3'] = next_level_list
        wall_gage_divide['Z(mm)4'] = next_level_list
                
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
        duplicates_list = []
        gage_node_data_list = []
        for idx, gage_node_data in wall_gage_sorted.groupby(['Wall ID', 'Z(mm)'])['Node List']:
            
            # 같은 Wall ID를 가지지만 붙어있지 않은 벽체 구별해내기
            gage_node_list_flat = [i for gage_node_list in gage_node_data for i in gage_node_list]
            duplicates = list(set([i for i in gage_node_list_flat if gage_node_list_flat.count(i) > 1])) # 위의 리스트에서 겹치는 부재들 remove
            
            duplicates_list.append(duplicates)
            gage_node_data_list.append(gage_node_data)
        
        # 같은 wall mark를 갖고, 겹치는 Node가 있는 벽체, 없는 벽체를 구분
        gage_node_list = []
        for gage_node_data, duplicates in zip(gage_node_data_list, duplicates_list):
            
            if len(gage_node_data) > 1: # 같은 Index(Wall ID, Z(mm))에 2개 이상의 벽체가 Assign 되어있을때    
                gage_node_sublist = []
                for gage_node_subdata in gage_node_data:
                                
                    if any(i in gage_node_subdata for i in duplicates): # duplicates_list의 중복되는 node가 하나라도 포함되어있는 경우
                        gage_node_sublist.append(gage_node_subdata)
                    
                    else:
                        gage_node_list.append([gage_node_subdata])
                            
                    gage_node_list.append(gage_node_sublist)
                
            else:
                gage_node_list.append(gage_node_data.tolist())
        
        # Node List 생성 (Node 번호순으로 재배열)        
        gage_node_list_zip = []
        for gage_node_sublist in gage_node_list:
            if len(gage_node_sublist) > 1:
                
                # 같은 Index(Wall ID, Z(mm))인 부재들의 Nodes를 Index에 맞춰 재배열한 list 만들기
                gage_node_sublist_zip = [list(i) for i in zip(*gage_node_sublist)]
                gage_node_list_zip.append(gage_node_sublist_zip)
    
            elif len(gage_node_sublist) == 1:
                gage_node_list_zip.append(gage_node_sublist)
                
        # 위에서 재배열한 list를 flatten
        gage_node_list_flat = []
        for gage_node_sublist_zip in gage_node_list_zip:
            if len(gage_node_sublist_zip) > 1:
                gage_node_sublist_flat = [i for gage_node_sublist_sublist_zip in gage_node_sublist_zip for i in gage_node_sublist_sublist_zip]
                gage_node_list_flat.append(gage_node_sublist_flat)
                
            elif len(gage_node_sublist_zip) == 1:
                gage_node_list_flat.append(gage_node_sublist_zip[0])
                
        # 중복되는 list 제거
        gage_node_list_flat_set = set(map(tuple, gage_node_list_flat)) # list -> tuple (to make it hashable)
        gage_node_list_flat_reduced = map(list, gage_node_list_flat_set) # tuple -> list  
    
        # sublist에서 중복되는 element 제거
        gage_node_list_flat_set_reduced = []
        for i in gage_node_list_flat_set:
            temp = [x for x in i if i.count(x) == 1]
            gage_node_list_flat_set_reduced.append(temp)
    
        # Node 번호에 맞는 좌표 매칭 후 출력
        gage_node_coord = pd.DataFrame(gage_node_list_flat_set_reduced)
        gage_node_coord.columns = ['Node1', 'Node2', 'Node3', 'Node4']
        
        # Merge로 Node 번호에 맞는 좌표를 결합
        gage_node_coord = pd.merge(gage_node_coord, node, how='left', left_on='Node1', right_on='Node', suffixes=(None, '1'))
        gage_node_coord = pd.merge(gage_node_coord, node, how='left', left_on='Node2', right_on='Node', suffixes=(None, '2'))
        gage_node_coord = pd.merge(gage_node_coord, node, how='left', left_on='Node3', right_on='Node', suffixes=(None, '3'))
        gage_node_coord = pd.merge(gage_node_coord, node, how='left', left_on='Node4', right_on='Node', suffixes=(None, '4'))
    
        gage_node_coord = gage_node_coord.iloc[:, [5,6,7,9,10,11,17,18,19,13,14,15]]
        
        # Gage Element 결과값을 csv로 변환
        gage_node_coord.to_csv(output_csv_dir+'\\'+wall_gage_csv, mode='w', index=False)
    
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

def naming(input_path, input_xlsx, drift_position=[2,5,7,11]):
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

    #%% section, frame 이름 만들기 위한 정보 load
    
    section_info_xlsx_sheet = 'Wall Naming' # section naming 관련된 정보만 들어있는 시트
    beam_info_xlsx_sheet = 'Beam Naming'
    column_info_xlsx_sheet = 'Column Naming'
    story_info_xlsx_sheet = 'Story Data' # 층 정보 sheet
    drift_info_xlsx_sheet = 'ETC' # Drift 정보 sheet

    section_info = pd.read_excel(input_path + '\\' + input_xlsx, sheet_name = section_info_xlsx_sheet, skiprows = 3)
    section_info.columns = ['Name', 'Story(from)', 'Story(to)', 'Amount']

    # Beam에 대해서도 똑같이...
    beam_info = pd.read_excel(input_path + '\\' + input_xlsx, sheet_name = beam_info_xlsx_sheet, skiprows = 3)
    beam_info.columns = ['Name', 'Story(from)', 'Story(to)', 'Amount']
    
    # Column에 대해서도 똑같이...
    column_info = pd.read_excel(input_path + '\\' + input_xlsx, sheet_name = column_info_xlsx_sheet, skiprows = 3)
    column_info.columns = ['Name', 'Story(from)', 'Story(to)', 'Amount']

    #%% story 정보 load
    
    story_info = pd.read_excel(input_path + '\\' + input_xlsx, sheet_name = story_info_xlsx_sheet, skiprows = [0,2,3])

    story_info_reversed = story_info[::-1] # 배열이 내가 원하는 방향과 반대로 되어있어서, 리스트 거꾸로만들었음
    story_info_reversed.reset_index(inplace=True, drop=True)

    #%% Section 이름 뽑기

    # for문으로 section naming에 사용할 섹션 이름(section_name_output) 뽑기
    section_name_output = [] # 결과로 나올 section_name_output 리스트 미리 정의

    for wall_name_parameter, amount_parameter, story_from_parameter, story_to_parameter\
        in zip(section_info['Name'], section_info['Amount'], section_info['Story(from)'], section_info['Story(to)']):  # for 문에 조건 여러개 달고싶을 때는 zip으로 묶어서~ 
        
        story_from_index = story_info_reversed[story_info_reversed['Story Name'] == story_from_parameter].index[0]  # story_from이 문자열이라 story_from을 사용해서 slicing이 안되기 때문에(내 지식선에서) .index로 story_from의 index만 뽑음
        story_to_index = story_info_reversed[story_info_reversed['Story Name'] == story_to_parameter].index[0]  # 마찬가지로 story_to의 index만 뽑음
        story_window = story_info_reversed['Story Name'][story_from_index : story_to_index + 1]  # 내가 원하는 층 구간(story_from부터 story_to까지)만 뽑아서 리스트로 만들기
        for i in range(1, amount_parameter + 1):  # (벽체 개수(amount))에 맞게 numbering하기 위해 1,2,3,4...amount[i]개의 배열을 만듦. 첫 시작을 1로 안하면 index 시작은 0이 default값이기 때문에 1씩 더해줌
            for current_story_name in story_window:
                if isinstance(current_story_name, str) == False:  # 층이름이 int인 경우, 이름조합을 위해 str로 바꿈
                    current_story_name = str(current_story_name)
                else:
                    pass
                
                section_name_output.append(wall_name_parameter + '_' + str(i) + '_' + current_story_name)  # 반복될때마다 생성되는 section 이름을 .append를 이용하여 리스트의 끝에 하나씩 쌓아줌. i값은 숫자라 .astype(str)로 string으로 바꿔줌

    # 층전단력 확인을 위한 층 섹션 이름 뽑기
    # Base section 추가하기
    story_section_name_output = ['Base']

    # 각 층 전단력 확인을 위한 각 층 section 추가하기
    for i in story_info_reversed['Story Name'][1:story_info_reversed.shape[0]]:
        story_section_name_output.append(i + '_Shear')

    #%% Frame 이름 뽑기

    # Wall Frame 이름 뽑기
    frame_wall_name_output = []

    for row in section_info.values: # for문을 빠르게 연산하기 위해 dataframe -> array    
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
        
    constraints_name = constraints_name[1:]

    #%% Drift 이름 뽑기

    # Drift의 방향 지정
    direction_list = ['X', 'Y']

    drift_name_output = []

    for position in drift_position:
        for direction in direction_list:
            for current_story_name in story_info['Story Name']:
                if isinstance(current_story_name, str) == False:  # 층이름이 int인 경우, 이름조합을 위해 str로 바꿈
                    current_story_name = str(current_story_name)
                drift_name_output.append(current_story_name + '_' + str(int(position)) + '_' + direction)
                    
    #%% 출력

    name_output = pd.DataFrame(({'Frame(Beam) Name': pd.Series(frame_beam_name_output),\
                                 'Frame(Column) Name': pd.Series(frame_column_name_output),\
                                 'Frame(Wall) Name': pd.Series(frame_wall_name_output),\
                                 'Constraints Name': pd.Series(constraints_name),\
                                 'Section(Wall) Name': pd.Series(section_name_output),\
                                 'Section(Shear) Name': pd.Series(story_section_name_output),\
                                 'Drift Name': pd.Series(drift_name_output)}))

    # Output 경로 설정
    # name_output_xlsx = 'Naming Output Sheets.xlsx'
    # 개별 엑셀파일로 출력
    # name_output.to_excel(input_path+ '\\'+ name_output_xlsx, sheet_name = 'Name List', index = False)

    # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
    name_output = name_output.replace(np.nan, '', regex=True)
    
    # Using win32com...
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application') # 엑셀 실행
    excel.Visible = False # 엑셀창 안보이게

    wb = excel.Workbooks.Open(input_path + '\\' + input_xlsx)
    ws = wb.Sheets('Output_Naming')
    
    startrow, startcol = 5, 1

    # 이름 열 입력
    ws.Range(ws.Cells(startrow, startcol),\
             ws.Cells(startrow + name_output.shape[0]-1,\
                      name_output.shape[1])).Value\
    = list(name_output.itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능   
    
    wb.Close(SaveChanges=1) # Closing the workbook
    excel.Quit() # Closing the application    