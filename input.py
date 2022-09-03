import pandas as pd

#%% Node, Element, Mass, Load Import

def import_midas(input_path, input_xlsx, DL_name='DL', LL_name='LL'\
                 , import_node=True, import_DL=True, import_LL=True\
                 , import_mass=True, **kwargs):
    
    defaultkwargs = {'import_beam':True, 'import_wall':True, 'import_plate':True\
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
    
# Input 경로 설정
input_path = r'D:\이형우\내진성능평가\광명 4R\103'
input_xlsx = 'Input Sheets(103_9)_v.1.8.xlsx'

DL_name = ['DL'] # DL에 포함시킬 하중이름 포함("DL_XX"와 같은 형태의 하중들만 있을 경우, "DL"만 넣어주면 됨)
LL_name = ['LL']

input_xlsx_sheet = 'Nodes'
nodal_load_raw_xlsx_sheet = 'Nodal Loads'
mass_raw_xlsx_sheet = 'Story Mass'
element_raw_xlsx_sheet = 'Elements'

# Output 경로 설정
output_csv_dir = input_path # 또는 '경로'

node_DL_merged_csv = 'DL.csv'
node_LL_merged_csv = 'LL.csv'
mass_csv = 'Mass.csv'
node_csv = 'Node.csv'
beam_csv = 'Beam.csv'
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
    node_mass_considered = node.append(mass2.iloc[:,[0,1,2]])

    # Node 결과값을 csv로 변환
    if import_node == True:
        node_mass_considered.to_csv(output_csv_dir+'\\'+node_csv, mode='w', index=False)
        
    else:
        mass2.iloc[:,[0,1,2]].to_csv(output_csv_dir+'\\'+node_csv, mode='w', index=False)
        
else:
    # Node 결과값을 csv로 변환
    node.to_csv(output_csv_dir+'\\'+node_csv, mode='w', index=False)

#%% Beam Element 뽑기

# Index로 지정되어있던 Node 번호를 다시 reset
node.index.name = 'Node'
node.reset_index(inplace=True)

# Element 정보 load
element = pd.read_excel(input_path+'\\'+input_xlsx, sheet_name = element_raw_xlsx_sheet, skiprows = 3)
element.columns = ['Element', 'Type', 'Wall Type', 'Sub Type', 'Wall ID', 'Material', 'Property', 'B-Angle', 'Node1', 'Node2', 'Node3', 'Node4']

# Beam Element만 추출(slicing)
if import_beam == True:
    
    beam = element.loc[lambda x: element['Type'] == 'BEAM', :]
    
    # 필요한 열만 추출(drop하기에는 drop할 열이 너무 많아서...)
    beam_node_1 = beam.loc[:, 'Node1']
    beam_node_2 = beam.loc[:, 'Node2']
    
    beam_node_1.name = 'Node'  # Merge(같은 열을 기준으로 두 dataframe 결합)를 사용하기 위해 index를 Node로 바꾸기
    beam_node_2.name = 'Node'
    
    # Merge로 Node 번호에 맞는 좌표를 결합
    beam_node_1_coord = pd.merge(beam_node_1, node, how='left')  # how='left' : 두 데이터프레임 중 왼쪽 데이터프레임은 그냥 두고 오른쪽 데이터프레임값을 대응시킴
    beam_node_2_coord = pd.merge(beam_node_2, node, how='left')
    
    # Node1, Node2의 좌표를 모두 결합시켜 출력
    beam_node_1_coord = beam_node_1_coord.drop('Node', axis=1)
    beam_node_2_coord = beam_node_2_coord.drop('Node', axis=1)
    
    beam_node_1_coord.columns = ['X_1(mm)', 'Y_1(mm)', 'Z_1(mm)']  # 결합 때 이름이 중복되면 안되서 이름 바꿔줌
    beam_node_2_coord.columns = ['X_2(mm)', 'Y_2(mm)', 'Z_2(mm)']
    
    beam_node_coord = pd.concat([beam_node_1_coord, beam_node_2_coord], axis=1)
    
    # Beam Element 결과값을 csv로 변환
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
    
    # wall_node_coord_list = [wall_node_1_coord, wall_node_2_coord, wall_node_3_coord, wall_node_4_coord]
    wall_node_coord = pd.concat([wall_node_1_coord, wall_node_2_coord, wall_node_3_coord, wall_node_4_coord], axis=1)
    
    # Wall Element 결과값을 csv로 변환
    wall_node_coord.to_csv(output_csv_dir+'\\'+wall_csv, mode='w', index=False) 

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

#%% Shear Wall Gage 뽑기
if import_SWR_gage == True:
    
    # Wall Element만 추출(slicing)
    wall = element.loc[lambda x: element['Type'] == 'WALL', :]
    
    wall_gage = wall.loc[:,['Wall ID', 'Node1', 'Node2', 'Node3', 'Node4']]
    
    # Merge로 Node 번호에 맞는 좌표를 결합
    wall_gage = pd.merge(wall_gage, node, how='left', left_on='Node1', right_on='Node', suffixes=(None, '1'))
    wall_gage = pd.merge(wall_gage, node, how='left', left_on='Node2', right_on='Node', suffixes=(None, '2'))
    wall_gage = pd.merge(wall_gage, node, how='left', left_on='Node3', right_on='Node', suffixes=(None, '3'))
    wall_gage = pd.merge(wall_gage, node, how='left', left_on='Node4', right_on='Node', suffixes=(None, '4'))
    
    # 필요한 열 뽑고 재정렬
    wall_gage = wall_gage.iloc[:,[0,1,2,3,4,6,7,8,10,11,12,14,15,16,18,19,20]]
    wall_gage.columns = ['Wall ID', 'Node1', 'Node2', 'Node3', 'Node4', 'X(mm)1'\
                         , 'Y(mm)1', 'Z(mm)1', 'X(mm)2', 'Y(mm)2', 'Z(mm)2', 'X(mm)3'\
                         , 'Y(mm)3', 'Z(mm)3', 'X(mm)4', 'Y(mm)4', 'Z(mm)4']
    
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
        duplicates = list(set([i for i in gage_node_list_flat if gage_node_list_flat.count(i) > 1]))
        
        duplicates_list.append(duplicates)
        gage_node_data_list.append(gage_node_data)
    
    # 겹치는 Node가 있는 벽체, 없는 벽체를 구분
    gage_node_list = []
    for gage_node_data, duplicates in zip(gage_node_data_list, duplicates_list):
        
        if len(gage_node_data) > 1: # 같은 Index(Wall ID, Z(mm))에 2개 이상의 벽체가 Assign 되어있을때    
            gage_node_sublist = []
            for gage_node_subdata in gage_node_data:
                            
                if any(i in gage_node_subdata for i in duplicates):
                    gage_node_sublist.append(gage_node_subdata)
                
                else:
                    gage_node_list.append(gage_node_subdata)
                        
                # gage_node_sublist_sublist_set = set(map(tuple, gage_node_sublist_sublist))
                gage_node_list.append(gage_node_sublist)
            
        else:
            gage_node_list.append(gage_node_data.tolist())
    
    # 중복되는 노드 제거한 후 Node List 생성        
    gage_node_list_zip = []
    for gage_node_sublist in gage_node_list:
        if len(gage_node_sublist) > 1:

     
            # 같은 Index(Wall ID, Z(mm))인 부재들의 Nodes를 Index에 맞춰 재배열한 list 만들기
            gage_node_sublist_zip = [list(i) for i in zip(*gage_node_sublist)]
            gage_node_list_zip.append(gage_node_sublist_zip)
            
        if len(gage_node_sublist) == 1:
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




            
    
    # Node1, Node2, Node3, Node4의 좌표를 모두 결합시켜 출력
    wall_gage_node_1_coord = wall_gage_node_1_coord.drop('Node', axis=1)
    wall_gage_node_2_coord = wall_gage_node_2_coord.drop('Node', axis=1)
    wall_gage_node_3_coord = wall_gage_node_3_coord.drop('Node', axis=1)
    wall_gage_node_4_coord = wall_gage_node_4_coord.drop('Node', axis=1)
    
    wall_gage_node_1_coord.columns = ['X_1(mm)', 'Y_1(mm)', 'Z_1(mm)']
    wall_gage_node_2_coord.columns = ['X_2(mm)', 'Y_2(mm)', 'Z_2(mm)']
    wall_gage_node_3_coord.columns = ['X_3(mm)', 'Y_3(mm)', 'Z_3(mm)']
    wall_gage_node_4_coord.columns = ['X_4(mm)', 'Y_4(mm)', 'Z_4(mm)']
    
    # wall_node_coord_list = [wall_node_1_coord, wall_node_2_coord, wall_node_3_coord, wall_node_4_coord]
    wall_gage_node_coord = pd.concat([wall_gage_node_1_coord, wall_gage_node_2_coord\
                                      , wall_gage_node_3_coord, wall_gage_node_4_coord], axis=1)
    
    # Wall Element 결과값을 csv로 변환
    wall_gage_node_coord.to_csv(output_csv_dir+'\\'+wall_gage_csv, mode='w', index=False) 
    
    
    # 필요한 열만 추출
    wall_ID = wall.loc[:, 'Wall ID']
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
    
    # wall_node_coord_list = [wall_node_1_coord, wall_node_2_coord, wall_node_3_coord, wall_node_4_coord]
    wall_node_coord = pd.concat([wall_node_1_coord, wall_node_2_coord, wall_node_3_coord, wall_node_4_coord], axis=1)
    
    # Wall Element 결과값을 csv로 변환
    wall_gage_node_coord.to_csv(output_csv_dir+'\\'+wall_csv, mode='w', index=False) 