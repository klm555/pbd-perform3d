import pandas as pd
import numpy as np
import pickle
import os
from collections import deque  # Double-ended Queue : 자료의 앞, 뒤 양 방향에서 자료를 추가하거나 제거가능
import matplotlib.pyplot as plt
import matplotlib as mpl
import win32com.client
import pythoncom

#%% Beam Rotation (DCR)
def BR(self, input_xlsx_path, beam_design_xlsx_path, graph=True, DCR_criteria=1, yticks=3, xlim=3):

#%% Load Data
    # Data Conversion Sheets
    story_info = self.story_info
    beam_info = self.beam_info

    # Analysis Result Sheets
    node_data = self.node_data
    element_data = self.frame_data
    beam_rot_data = self.beam_rot_data

    # Seismic Loads List
    load_name_list = self.load_name_list
    gravity_load_name = self.gravity_load_name
    seismic_load_name_list = self.seismic_load_name_list
    DE_load_name_list = self.DE_load_name_list
    MCE_load_name_list = self.MCE_load_name_list
    
    # Data Conversion Sheets
    # story_info = result.story_info
    # beam_info = result.beam_info

    # # Analysis Result Sheets
    # node_data = result.node_data
    # element_data = result.frame_data
    # beam_rot_data = result.beam_rot_data

    # # Seismic Loads List
    # load_name_list = result.load_name_list
    # gravity_load_name = result.gravity_load_name
    # seismic_load_name_list = result.seismic_load_name_list
    # DE_load_name_list = result.DE_load_name_list
    # MCE_load_name_list = result.MCE_load_name_list
    
#%% Process Data    
    # node, element data에서 필요한 정보만 추출
    node_data = node_data.iloc[:,[0,3]]    
    element_data = element_data.iloc[:,[0,1,2]]

    # temporary ((L), (R) 등 지우기)
    element_data.loc[:, 'Property Name'] = element_data.loc[:, 'Property Name'].str.split('(').str[0]

    # 필요한 부재만 선별
    prop_name = beam_info.iloc[:,0]
    prop_name.name = 'Property Name'
    element_data = element_data[element_data['Property Name'].isin(prop_name)]

    element_data = element_data.drop_duplicates()
    node_data = node_data.drop_duplicates()   

    # Analysis Result에 Element, Node 정보 매칭    
    beam_rot_data = pd.merge(beam_rot_data, element_data, how='left')
    beam_rot_data = pd.merge(beam_rot_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
    
    # 필요없는 부재 빼기, 필요한 부재만 추출
    beam_rot_data = beam_rot_data[(beam_rot_data['Point ID'] == 1) | (beam_rot_data['Point ID'] == 5)]
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
    beam_rot_data = beam_rot_data.iloc[:, [0,1,7,10,2,3,4,11]]
    
    # 지진하중, i,j 노드, Max,Min에 따라 Rotation 데이터 Grouping
    BR_grouped_list = list(beam_rot_data.groupby(['Load Case', 'Point ID', 'Step Type']))
    
    # 해석 결과 상관없이 Full 지진하중 이름 list 만들기
    full_DE_load_name_list = 'DE' + pd.Series([11,12,21,22,31,32,41,42,51,52,61,62,71,72]).astype(str)
    full_MCE_load_name_list = 'MCE' + pd.Series([11,12,21,22,31,32,41,42,51,52,61,62,71,72]).astype(str)
    full_load_name_list = pd.concat([full_DE_load_name_list, full_MCE_load_name_list])
    
    # 이름만 들어간 Dataframe 만들기
    BR_output = pd.DataFrame(prop_name)
    # 지진하중, i,j 노드, Max,Min loop 돌리기
    for load_name in full_load_name_list:
        for point_id in ['1', '5']:
            for max_min in ['Max', 'Min']:
                # 만들어진 Group List loop 돌리기
                for BR_grouped in BR_grouped_list:
                    if (load_name in BR_grouped[0][0]) & (BR_grouped[0][1] == int(point_id)) & (BR_grouped[0][2] == max_min):
                        # 같은 결과가 2개씩 있어서 drop_duplicates
                        BR_grouped_df = BR_grouped[1].drop_duplicates()
                        # Element Name이 같은 경우, 큰 값만 선택
                        # BR_grouped_df = BR_grouped_df
                        BR_grouped_df = pd.merge(prop_name, BR_grouped_df, how='left')
                        BR_grouped_df.reset_index(inplace=True, drop=True)
                        BR_output = pd.concat([BR_output, BR_grouped_df['Major Rotation(rad)']], axis=1)
                    
                # 해당 지진하중의 해석결과가 없는 경우 Blank Column 생성
                if load_name not in seismic_load_name_list: 
                    blank_col = pd.Series([''] * len(prop_name))
                    BR_output = pd.concat([BR_output, blank_col], axis=1)    
                    
#%% 결과 정리 후 Input Sheets에 넣기

# 출력용 Dataframe 만들기
    # Design_C.Beam 시트
    steel_design_df = beam_info.iloc[:,21:31]
    beam_output = pd.concat([beam_info, steel_design_df], axis=1)
    
    # Table_C.Beam_DE 시트
    beam_info[['Beam Name', 'Beam Number', 'Floor']] = beam_info['Name'].str.split('_', expand=True)
    # 벽체 이름, 번호에 따라 grouping
    beam_name_list = list(beam_info.groupby(['Beam Name', 'Beam Number'], sort=False))
    # 55 row짜리 empty dataframe 만들기
    name_empty = pd.DataFrame(np.nan, index=range(55), columns=range(len(beam_name_list)))
    # dataframe에 이름 채워넣기
    count = 0
    while True:
        name_iter = beam_name_list[count][0][0]
        num_iter = beam_name_list[count][0][1]
        total_iter = beam_info['Name'][(beam_info['Beam Name'] == name_iter) 
                                       & (beam_info['Beam Number'] == num_iter)]
        name_empty.iloc[range(len(total_iter)), count] = total_iter
        
        count += 1
        if count == len(beam_name_list):
            break
    # dataframe을 1열로 만들기
    name_output_arr = np.array(name_empty)
    name_output_arr = np.reshape(name_output_arr, (-1, 1), order='F')
    name_output = pd.DataFrame(name_output_arr)
    
    # Plot_C.Beam_DE 시트
    plot_num = pd.DataFrame()
    # 첫번째 부재(Beam)의 이름 추출
    plot_num['Name'] = beam_info['Name'][(beam_info['Beam Name'] == beam_name_list[0][0][0]) 
                                 & (beam_info['Beam Number'] == beam_name_list[0][0][1])]    
    plot_num[['Beam Name', 'Beam Number', 'Story Name']] = plot_num['Name'].str.split('_', expand=True)
    # story_info의 Index열을 1부터 시작하도록 재지정
    story_info['Index'] = range(story_info.shape[0], 0, -1)
    # Ground Level(0mm, 1F)에 가장 가까운 층의 row index get
    ground_level_idx = story_info['Height(mm)'].abs().idxmin()
    # Ground Level(0mm, 1F)에 가장 가까운 층에 59번 row 배정
    plot_place_arr = np.zeros(story_info.shape[0], dtype=int)
    for idx, x in enumerate(plot_place_arr): # 59 + index 는 일정하다는 사실을 이용
        plot_place_arr[idx] = 59 + ground_level_idx - idx
    story_info['Place'] = plot_place_arr
    
    # Story 정보 merge
    plot_num = pd.merge(plot_num, story_info, how='left')
    
    # Reverse the order of row
    plot_num = plot_num.iloc[::-1]

    # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
    BR_output = BR_output.replace(np.nan, '', regex=True)
    beam_output = beam_output.replace(np.nan, '', regex=True)
    name_output = name_output.replace(np.nan, '', regex=True)
    
#%% 조작용 코드
    # 없애고 싶은 부재의 이름 입력(error_beam 확인 후!, DE, MCE에서 다 없어짐)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('PB1-10_1'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('PB1-8_1'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('LB1A_2'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('LB1A_4'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('LB2_1'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB4B_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB3D_'))].index)

#%% 엑셀로 출력(Using win32com)
        
    # Using win32com...
    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게
    
    wb = excel.Workbooks.Open(beam_design_xlsx_path)
    ws1 = wb.Sheets('Results_C.Beam_Rotation')
    ws2 = wb.Sheets('Design_C.Beam')
    ws3 = wb.Sheets('Table_C.Beam_DE')
    ws4 = wb.Sheets('Plot_C.Beam_DE')
    
    startrow, startcol = 5, 1
    
    # Results_C.Beam_Rotation 시트 입력
    ws1.Range('A%s:DI%s' %(startrow, startrow + BR_output.shape[0] - 1)).Value\
        = list(BR_output.itertuples(index=False, name=None))
        
    # Design_C.Beam 시트 입력
    ws2.Range('A%s:AO%s' %(startrow, startrow + beam_output.shape[0] - 1)).Value\
        = list(beam_output.itertuples(index=False, name=None))
    
    # Table_C.Beam_DE 시트 입력
    ws3.Range('B%s:B%s' %(startrow, startrow + name_output.shape[0] - 1)).Value\
        = [[i] for i in name_output[0]] # series -> list 형식만 입력가능
    ws3.Range('A4:A4').Value\
        = len(beam_name_list) # series -> list 형식만 입력가능
        
    # Plot_S.Wall_DE 시트 입력
    ws4.Range('A%s:A%s' %(plot_num['Place'].min(), plot_num['Place'].max())).Value\
        = [[i] for i in plot_num['Index']] # series -> list 형식만 입력가능
    
    wb.Save()
    # wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application 
        
#%% 그래프
    if graph == True:
        # Beam 정보 load
        ws_DE = wb.Sheets('Table_C.Beam_DE')
        ws_MCE = wb.Sheets('Table_C.Beam_MCE')
        
        DE_result = ws_DE.Range('K%s:L%s' %(startrow, startrow + name_output.shape[0] - 1)).Value
        DE_result_arr = np.array(DE_result)[:,[0,1]]
        MCE_result = ws_MCE.Range('K%s:L%s' %(startrow, startrow + name_output.shape[0] - 1)).Value
        MCE_result_arr = np.array(MCE_result)[:,[0,1]]
        perform_lv = ws_DE.Range('M%s:O%s' %(startrow, startrow + name_output.shape[0] - 1)).Value
        perform_lv_arr = np.array(perform_lv)[:,[0,1,2]]
        
        # DCR 계산을 위해 결과값, Performance Level 합쳐서 Dataframe 생성
        BR_plot = np.concatenate((DE_result_arr, MCE_result_arr, perform_lv_arr), axis=1)
        BR_plot = pd.DataFrame(BR_plot)
        BR_plot.columns = ['DE_pos', 'DE_neg', 'MCE_pos', 'MCE_neg', 'IO', 'LS', 'CP']
        # DCR 계산
        BR_plot = BR_plot.apply(pd.to_numeric)
        BR_plot['DCR(DE_pos)'] = BR_plot['DE_pos'] / BR_plot['LS']
        BR_plot['DCR(DE_neg)'] = BR_plot['DE_neg'] / BR_plot['LS'] * (-1)
        BR_plot['DCR(MCE_pos)'] = BR_plot['MCE_pos'] / BR_plot['CP']
        BR_plot['DCR(MCE_neg)'] = BR_plot['MCE_neg'] / BR_plot['CP'] * (-1)
        
        BR_plot['Name'] = name_output.copy()
        
        # 결과 dataframe -> pickle
        BR_result = []
        BR_result.append(BR_plot)
        BR_result.append(story_info)
        BR_result.append(DE_load_name_list)
        BR_result.append(MCE_load_name_list)
        with open('pkl/BR.pkl', 'wb') as f:
            pickle.dump(BR_result, f)

#%% C.Beam SF (DCR)
def BSF(self, input_xlsx_path, beam_design_xlsx_path, graph=True, DCR_criteria=1, yticks=2, xlim=3):
    ''' 

    Perform-3D 해석 결과에서 일반기둥의 축력, 전단력을 불러와 Results_G.Column 엑셀파일을 작성. \n
    result_path : Perform-3D에서 나온 해석 파일의 경로. \n
    result_xlsx : Perform-3D에서 나온 해석 파일의 이름. 해당 파일 이름이 포함된 파일들을 모두 불러온다. \n
    input_path : Data Conversion 엑셀 파일의 경로 \n
    input_xlsx : Data Conversion 엑셀 파일의 이름. result_xlsx와는 달리 확장자명(.xlsx)까지 기입해줘야한다. 하나의 파일만 불러온다. \n
    column_xlsx : Results_E.Column 엑셀 파일의 이름.확장자명(.xlsx)까지 기입해줘야한다. \n
    export_to_pdf : 입력된 값에 따른 각 부재들의 결과 시트를 pdf로 출력. True = pdf 출력, False = pdf 미출력(Results_E.Column 엑셀파일만 작성됨).
    pdf_name = 출력할 pdf 파일 이름.
    
    '''
#%% Load Data
    # Data Conversion Sheets        
    story_info = self.story_info
    deform_cap = self.beam_deform_cap

    # Analysis Result Sheets
    node_data = self.node_data
    element_data = self.frame_data
    SF_info_data = self.beam_shear_force_data

    # Seismic Loads List
    load_name_list = self.load_name_list
    gravity_load_name = self.gravity_load_name
    seismic_load_name_list = self.seismic_load_name_list
    DE_load_name_list = self.DE_load_name_list
    MCE_load_name_list = self.MCE_load_name_list

#%% Process Data
    # 필요한 부재만 선별
    prop_name = deform_cap.iloc[:,0]
    prop_name.name = 'Property Name'
    element_data = element_data[element_data['Property Name'].isin(prop_name)]

    element_data = element_data.drop_duplicates()
    
    # Analysis Result에 Element, Node 정보 매칭
    element_data = pd.merge(element_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
    SF_ongoing = pd.merge(element_data, SF_info_data, how='left')
    SF_ongoing.reset_index(inplace=True, drop=True)

#%% V값의 절대값, 최대값, 평균값 뽑기

    # 절대값
    SF_ongoing['V2 I-End'] = SF_ongoing['V2 I-End'].abs()

    # V2의 최대값을 저장하기 위해 필요한 데이터 slice
    SF_ongoing_max = SF_ongoing.iloc[[2*x for x in range(int(SF_ongoing.shape[0]/2))]]
    SF_ongoing_max = SF_ongoing_max.loc[:, ['Element Name', 'Property Name', 'V', 'Load Case']]                       
    # [2*x for x in range(int(SF_ongoing.shape[0]/2] -> [짝수 index]
    
    # V2, V3의 최대값을 저장
    SF_ongoing_max['V2 max'] = SF_ongoing.groupby(SF_ongoing.index // 2)['V2 I-End'].max().tolist()

    # 필요한 하중만 포함된 데이터 slice (MCE)
    SF_ongoing_max_MCE = SF_ongoing_max[SF_ongoing_max['Load Case']\
                                        .str.contains('|'.join(MCE_load_name_list))]
    SF_ongoing_max_G = SF_ongoing_max[SF_ongoing_max['Load Case']\
                                      .str.contains('|'.join(gravity_load_name))]
    # function equivalent of a combination of df.isin() and df.str.contains()
    
    # 부재별(Element Name) 평균값을 저장하기 위해 필요한 데이터프레임 생성
    SF_ongoing_max_avg = SF_ongoing_max_MCE.iloc[:,[0,1,2]]
    SF_ongoing_max_avg = SF_ongoing_max_avg.drop_duplicates()
    SF_ongoing_max_avg.set_index('Element Name', inplace=True)    
    # 부재별(Element Name) 평균값 뽑기
    SF_ongoing_max_avg['V2 max(MCE)'] = SF_ongoing_max_MCE.groupby(['Element Name'])['V2 max'].mean()
    SF_ongoing_max_avg['V2 max(G)'] = SF_ongoing_max_G.groupby(['Element Name'])['V2 max'].mean()
    
    # 이름별(Property Name) 최대값을 저장하기 위해 필요한 데이터프레임 생성
    SF_ongoing_max_avg_max = SF_ongoing_max_avg.copy()
    SF_ongoing_max_avg_max = SF_ongoing_max_avg_max.drop_duplicates(subset=['Property Name'], ignore_index=True)
    SF_ongoing_max_avg_max.set_index('Property Name', inplace=True) 
    # 같은 부재(그러나 잘려있는) 경우(Property Name) 최대값 뽑기
    SF_ongoing_max_avg_max = pd.merge(SF_ongoing_max_avg_max
                                      , SF_ongoing_max_avg.groupby(['Property Name'])['V2 max(MCE)'].max()
                                      , left_on='Property Name', right_index=True, suffixes=('_before', '_after'))
    SF_ongoing_max_avg_max = pd.merge(SF_ongoing_max_avg_max
                                      , SF_ongoing_max_avg.groupby(['Property Name'])['V2 max(G)'].max()
                                      , left_on='Property Name', right_index=True, suffixes=('_before', '_after'))
    
    # MCE에 대해 1.2배, G에 대해 0.2배
    SF_ongoing_max_avg_max['V2 max(MCE)_after'] = SF_ongoing_max_avg_max['V2 max(MCE)_after'] * 1.2
    SF_ongoing_max_avg_max['V2 max(G)_after'] = SF_ongoing_max_avg_max['V2 max(G)_after'] * 0.2
    
    SF_ongoing_max_avg_max.reset_index(inplace=True, drop=False)

#%% 결과값 정리
    
    SF_output = pd.merge(prop_name, SF_ongoing_max_avg_max, how='left')
        
    SF_output = SF_output.dropna()
    SF_output.reset_index(inplace=True, drop=True)
        
    # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
    SF_output = SF_output.replace(np.nan, '', regex=True)
    
    # 기존 시트에 V값 넣기
    SF_output1 = SF_output.iloc[:,0]
    SF_output2 = SF_output.iloc[:,[4,5]]

#%% 출력 (Using win32com...)
    
    # Using win32com...
    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게

    wb = excel.Workbooks.Open(input_xlsx_path)
    ws = wb.Sheets('Results_C.Beam')
    
    startrow, startcol = 5, 1    
    ws.Range(ws.Cells(startrow, startcol),\
              ws.Cells(startrow+SF_output1.shape[0]-1, startcol)).Value\
    = [[i] for i in SF_output1]
    
    startrow, startcol = 5, 20    
    ws.Range(ws.Cells(startrow, startcol),\
              ws.Cells(startrow+SF_output2.shape[0]-1,\
                      startcol+SF_output2.shape[1]-1)).Value\
    = list(SF_output2.itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능
    
    wb.Save()            
    # wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application

#%% Elastic Beam SF (DCR)

def E_BSF(input_xlsx_path, result_xlsx_path, contour=True):

#%% Input Sheet 정보 load
        
    story_info = pd.DataFrame()
    transfer_element_info = pd.DataFrame()

    input_xlsx_sheet = 'Output_E.Beam Properties'
    input_data_raw = pd.ExcelFile(input_xlsx_path)
    input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', input_xlsx_sheet], skiprows=3)
    input_data_raw.close()

    story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
    transfer_element_info = input_data_sheets[input_xlsx_sheet].iloc[:,0]
    story_info = story_info[::-1]
    story_info.reset_index(inplace=True, drop=True)

    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    transfer_element_info.name = 'Name'

    #%% Analysis Result 불러오기

    to_load_list = result_xlsx_path

    # 전단력 Data
    SF_info_data = pd.DataFrame()
    for i in to_load_list:
        SF_info_data_temp = pd.read_excel(i, sheet_name='Frame Results - End Forces'
                                          , skiprows=[0, 2], header=0
                                          , usecols=[0,2,5,7,10,11,17,18]) # usecols로 원하는 열만 불러오기
        SF_info_data = pd.concat([SF_info_data, SF_info_data_temp])

    SF_info_data = SF_info_data.sort_values(by=['Load Case', 'Element Name', 'Step Type']) # 지진파 순서가 섞여있을 때 sort

    # 부재 이름 Matching을 위한 Element 정보
    element_info_data = pd.DataFrame()
    for i in to_load_list:
        element_info_data_temp = pd.read_excel(i, sheet_name='Element Data - Frame Types'
                                               , skiprows=[0, 2], header=0, usecols=[0,2,5,7,9]) # usecols로 원하는 열만 불러오기
        element_info_data = pd.concat([element_info_data, element_info_data_temp])

    # 필요한 부재만 선별
    element_info_data = element_info_data[element_info_data['Property Name'].isin(transfer_element_info)]
    
    # 기둥과 겹치는 등 평가에 반영하지 않을 부재 제거
    element_to_remove = ['E880','E26229','E555','E671','E658','E525','E528','E932','E914','E1256','E1165','E585']
    element_info_data = element_info_data[~element_info_data['Element Name'].isin(element_to_remove)]
    
    # 층 정보 Matching을 위한 Node 정보
    node_info_data = pd.DataFrame()    
    for i in to_load_list:
        node_info_data_temp = pd.read_excel(i, sheet_name='Node Coordinate Data'
                                            , skiprows=[0, 2], header=0, usecols=[1,2,3,4]) # usecols로 원하는 열만 불러오기
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

    #%% (중력하중에 대한) V, M값에 절대값, 최대값 뽑기
    
    # 절대값 0.2배
    SF_ongoing_G = SF_ongoing.copy()
    SF_ongoing_G.iloc[:,[5,6,7,8]] = SF_ongoing_G.iloc[:,[5,6,7,8]].abs() * 0.2
    
    # i, j 노드 중 최대값 뽑기
    SF_ongoing_G['V2 max(G)'] = SF_ongoing_G[['V2 I-End', 'V2 J-End']].max(axis = 1)
    SF_ongoing_G['M3 max(G)'] = SF_ongoing_G[['M3 I-End', 'M3 J-End']].max(axis = 1)

    # max, min 중 최대값 뽑기
    SF_ongoing_G_max = SF_ongoing_G.loc[SF_ongoing_G.groupby(SF_ongoing_G.index // 2)['V2 max(G)'].idxmax()]
    SF_ongoing_G_max['M3 max(G)'] = SF_ongoing_G.groupby(SF_ongoing_G.index // 2)['M3 max(G)'].max().tolist()

    # 필요한 하중만 포함된 데이터 slice (MCE, G)
    SF_ongoing_G_max = SF_ongoing_G_max[SF_ongoing_G_max['Load Case']\
                                    .str.contains('|'.join(gravity_load_name))] # function equivalent of a combination of df.isin() and df.str.contains()

    # 부재별 최대값 뽑기
    SF_ongoing_elements = SF_ongoing_G_max.iloc[:,[0,1,2]]
    SF_ongoing_elements= SF_ongoing_elements.drop_duplicates()
    SF_ongoing_elements.set_index('Element Name', inplace=True)
    
    SF_ongoing_G_max_avg = SF_ongoing_elements.copy() # 평균값을 뽑진 않지만, 아래의 SF_ongoing_max_avg와 형태 맞춰주기위해 이렇게 naming됨 
    SF_ongoing_G_max_avg['V2 max(G)'] = SF_ongoing_G_max.groupby(['Element Name'])['V2 max(G)'].max()
    SF_ongoing_G_max_avg['M3 max(G)'] = SF_ongoing_G_max.groupby(['Element Name'])['M3 max(G)'].max()

    # 같은 부재(그러나 잘려있는) 경우 최대값 뽑기
    SF_ongoing_G_max_max = SF_ongoing_G_max_avg.loc[SF_ongoing_G_max_avg.groupby(['Property Name'])['V2 max(G)'].idxmax()]
    SF_ongoing_G_max_max['M3 max(G)'] = SF_ongoing_G_max_avg.groupby(['Property Name'])['M3 max(G)'].max().tolist()
    
    SF_ongoing_G_max_max.reset_index(inplace=True, drop=True) 

    #%% V, M값에 절대값, 최대값, 평균값 뽑기

    # 절대값, 1.2배
    SF_ongoing.iloc[:,[5,6,7,8]] = SF_ongoing.iloc[:,[5,6,7,8]].abs() * 1.2

    # i, j 노드 중 최대값 뽑기
    SF_ongoing['V2 max'] = SF_ongoing[['V2 I-End', 'V2 J-End']].max(axis = 1)
    SF_ongoing['M3 max'] = SF_ongoing[['M3 I-End', 'M3 J-End']].max(axis = 1)

    # max, min 중 최대값 뽑기
    SF_ongoing_max = SF_ongoing.loc[SF_ongoing.groupby(SF_ongoing.index // 2)['V2 max'].idxmax()]
    SF_ongoing_max['M3 max'] = SF_ongoing.groupby(SF_ongoing.index // 2)['M3 max'].max().tolist()

    # 필요한 하중만 포함된 데이터 slice (MCE, G)
    SF_ongoing_max = SF_ongoing_max[SF_ongoing_max['Load Case']\
                                    .str.contains('|'.join(MCE_load_name_list))] # function equivalent of a combination of df.isin() and df.str.contains()

    # 부재별 평균값 뽑기
    SF_ongoing_max_avg = SF_ongoing_elements.copy()    
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
    
    # 중력하중, 지진하중에 대한 V,M값 합치기
    SF_output = pd.merge(SF_ongoing_max_avg_max, SF_ongoing_G_max_max, how='left')
    
    # SF_ongoing_max_avg 재정렬
    SF_output = SF_output.iloc[:,[0,2,4,3,5]] 
    
    # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
    SF_output = SF_output.replace(np.nan, '', regex=True)
    
    # Using win32com...
    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게

    wb = excel.Workbooks.Open(input_xlsx_path)
    ws = wb.Sheets('Results_E.Beam')
    
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
    = list(SF_output.iloc[:,[1,2,3,4]].itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능
    
    wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application
    
#%% 부재의 위치별  V, M 값 확인을 위한 도면 작성
    # DCR 구하기    
    # 1.2Vus, 1.2Mus, 0.2Vuns, 0.2Muns 불러오기
    SF_ongoing_combined = pd.concat([SF_ongoing_G_max_avg, SF_ongoing_max_avg.iloc[:,[2,3]]]
                                   , axis=1)
    SF_ongoing_combined.reset_index(inplace=True, drop=False)
    
    # Vu, Mu 구하기
    SF_ongoing_combined['Vu'] = SF_ongoing_combined['V2 max'] - SF_ongoing_combined['V2 max(G)']
    SF_ongoing_combined['Mu'] = SF_ongoing_combined['M3 max'] - SF_ongoing_combined['M3 max(G)']
    
    # Mu unit 변경 (mm -> m)
    SF_ongoing_combined['Mu'] = SF_ongoing_combined['Mu'] / 1000
    
    # Vn, Mn값 불러오기
    e_beam_result = pd.read_excel(input_xlsx_path, sheet_name='Results_E.Beam'
                                  , skiprows=3, header=0)
    e_beam_result = e_beam_result.iloc[:,[0,29,32]]
    e_beam_result.columns = ['Property Name', 'Vn', 'Mn']
    e_beam_result.reset_index(inplace=True, drop=True)
    
    # DCR 구하기
    SF_ongoing_combined = pd.merge(SF_ongoing_combined, e_beam_result, how='left')
    SF_ongoing_combined['DCR(V)'] = SF_ongoing_combined['Vu'] / SF_ongoing_combined['Vn']
    SF_ongoing_combined['DCR(M)'] = SF_ongoing_combined['Mu'] / SF_ongoing_combined['Mn']
    

    # 도면을 그리기 위한 Node List 만들기    
    node_map_z = SF_ongoing_max_avg['i-V'].drop_duplicates()
    node_map_z.sort_values(ascending=False, inplace=True)
    node_map_list = node_info_data[node_info_data['V'].isin(node_map_z)]
    
    # 도면을 그리기 위한 Element List 만들기
    element_map_list = pd.merge(SF_ongoing_combined.iloc[:,[0,1,2,11,12]]
                                , element_info_data.iloc[:,[1,5,6,8,9]]
                                , how='left', on='Element Name')
    
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
        norm_V = plt.Normalize(vmin = element_map_list_extracted['DCR(V)'].min()\
                             , vmax = element_map_list_extracted['DCR(V)'].max())
        cmap_V_elem = cmap_V(norm_V(element_map_list_extracted['DCR(V)']))
        scalar_map_V = mpl.cm.ScalarMappable(norm_V, cmap_V)
        
        norm_M = plt.Normalize(vmin = element_map_list_extracted['DCR(M)'].min()\
                             , vmax = element_map_list_extracted['DCR(M)'].max())
        cmap_M_elem = cmap_M(norm_M(element_map_list_extracted['DCR(M)']))
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
        plt.colorbar(scalar_map_V, shrink=0.7, label='DCR (V)')
    
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
        plt.colorbar(scalar_map_M, shrink=0.7, label='DCR (M)')
    
        # 기타
        plt.axis('off')
        plt.title(story_info['Story Name'][story_info['Height(mm)'] == i].iloc[0])

        plt.tight_layout()   
        plt.close()
        count += 1
        yield fig2

#%% C.Beam SF (elementwise)

def BSF_each(input_xlsx_path, retrofit_sheet=None): 
    ''' 

    완성된 Results_Wall 시트에서 보강이 필요한 부재들이 OK될 때까지 자동으로 배근함. \n
    
       
    세로 생성되는 Results_Wall_보강 시트에 보강 결과 출력 (철근 type 변경, 해결 안될 시 spacing은 10mm 간격으로 down)
    
    Parameters
    ----------
    input_path : str
                 Data Conversion 엑셀 파일의 경로.
                 
    input_xlsx : str
                 Data Conversion 엑셀 파일의 이름. result_xlsx와는 달리 확장자명(.xlsx)까지 기입해줘야한다. 하나의 파일만 불러온다.

    Yields
    -------
    Min, Max값 모두 출력됨. 
    
    fig1 : matplotlib.pyplot.figure or None
           DE(설계지진) 발생 시 벽체 회전각 DCR 그래프                                      
    
    Raises
    -------
    
    References
    -------
    .. [1] "철근콘크리트 건축구조물의 성능기반 내진설계 지침", 대한건축학회, p.79, 2021
    
    '''
#%% Input Sheet
        
    # Input Sheets 불러오기
    input_xlsx_sheet = 'Results_C.Beam'
    input_data_raw = pd.ExcelFile(input_xlsx_path)
    input_data_sheets = pd.read_excel(input_data_raw, [input_xlsx_sheet, retrofit_sheet], skiprows=3
                                 , usecols=[0,10,15,16,29])
    input_data_raw.close()
    
    beam_before = input_data_sheets[input_xlsx_sheet]
    beam_after = input_data_sheets[retrofit_sheet]

    beam_before.columns = ['Name', 'Rebar Type(before)', 'Rebar EA(before)', 'Rebar Spacing(before)', 'Results(before)']
    beam_after.columns = ['Name', 'Rebar Type(after)', 'Rebar EA(after)', 'Rebar Spacing(after)', 'Results(after)']
    
#%% 보강 전,후 Wall dataframe 정리
    
    # DCR 열 반올림하기(소수점 5자리까지)
    beam_before['Results(before)'] = beam_before['Results(before)'].round(5)
    beam_after['Results(after)'] = beam_after['Results(after)'].round(5)

    # 필요한 열만 추출
    beam_output = pd.merge(beam_before, beam_after, how='left')

    # 이름 분리(벽체 이름, 번호, 층)
    beam_output['Property Name'] = beam_output['Name'].str.split('_', expand=True)[0]
    beam_output['Number'] = beam_output['Name'].str.split('_', expand=True)[1]
    beam_output['Story'] = beam_output['Name'].str.split('_', expand=True)[2]

    # 벽체 이름과 번호(W1_1)이 같은 부재들끼리 groupby로 묶고, list of dataframes 생성
    beam_output_list = list(beam_output.groupby(['Property Name', 'Number'], sort=False))
    
    yield beam_output_list

#%% Plastic Hinge Detector(Beam, Column)

def p_hinge(input_xlsx_path, result_xlsx_path, beam_group='C.Beam'
            , col_group='G.Column'):

#%% Input Sheets 정보 load
    story_info = pd.DataFrame()
    deformation_cap = pd.DataFrame()
    
    input_data_raw = pd.ExcelFile(input_xlsx_path)
    input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', 'Output_C.Beam Properties'
                                                       , 'Output_G.Column Properties'], skiprows=3)
    input_data_raw.close()
    
    story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
    deformation_cap_beam = input_data_sheets['Output_C.Beam Properties'].iloc[:,[0,54]]
    deformation_cap_col = input_data_sheets['Output_G.Column Properties'].iloc[:,[0,62]]
    
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    deformation_cap_beam.columns = ['Property Name', 'Performance Level 2']
    deformation_cap_col.columns = ['Property Name', 'Performance Level 2']
    
#%% Analysis Result 불러오기
    to_load_list = result_xlsx_path
    
    rot_data = pd.DataFrame()
    
    for i in to_load_list:
        result_data_raw = pd.ExcelFile(i)
        result_data_sheets = pd.read_excel(result_data_raw, ['Frame Results - Bending Deform', 'Node Coordinate Data',\
                                                         'Element Data - Frame Types'], skiprows=2)
        
        rot_data_temp = result_data_sheets['Frame Results - Bending Deform'].iloc[:,[0,2,5,7,10,13,14]]
        rot_data = pd.concat([rot_data, rot_data_temp])
        
    node_data = result_data_sheets['Node Coordinate Data'].iloc[:,[1,4]]
    
    element_data = result_data_sheets['Element Data - Frame Types'].iloc[:,[2,5,7]] # beam의 양 nodes중 한 node에서의 rotation * 2
    
                
    rot_data.columns = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'Distance from I-End', 'H2 Rotation(rad)', 'H3 Rotation(rad)']
    node_data.columns = ['Node ID', 'V']
    element_data.columns = ['Element Name', 'Property Name', 'I-Node ID']
    
#%% temporary ((L), (R) 등 지우기)
    element_data.loc[:, 'Property Name'] = element_data.loc[:, 'Property Name'].str.split('(').str[0]
    
    rot_data = rot_data[rot_data['Distance from I-End'] == 0]
    
#%% Analysis Result에 Element, Node 정보 매칭
    
    element_data = element_data.drop_duplicates()
    node_data = node_data.drop_duplicates()
    
    rot_data = pd.merge(rot_data, element_data, how='left')
    rot_data = pd.merge(rot_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
    
    rot_data = rot_data[rot_data['Property Name'].notna()]
    rot_data.reset_index(inplace=True, drop=True)

    #%% Beam, Column 정보만 추출

    beam_rot_data = rot_data[rot_data['Group Name'] == beam_group]
    col_rot_data = rot_data[rot_data['Group Name'] == col_group]

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
    
#%% beam_rot_data의 값 수정(H1, H2 방향 중 major한 방향의 rotation값만 추출, 그리고 2배)
    major_rot = []
    for i, j in zip(beam_rot_data['H2 Rotation(rad)'], beam_rot_data['H3 Rotation(rad)']):
        if abs(i) >= abs(j):
            major_rot.append(i)
        else: major_rot.append(j)
    
    beam_rot_data['Major Rotation(rad)'] = major_rot
     
    # 필요한 정보들만 다시 모아서 new dataframe
    beam_rot_data = beam_rot_data.iloc[:, [0,1,7,10,2,3,11]]

#%% col_rot_data에서 필요한 정보들만 다시 모아서 new dataframe
    col_rot_data = col_rot_data.iloc[:, [0,1,7,10,2,3,5,6]]
    
#%% DE, MCE 각각의 load case에 대해 max, min 값 추출 / 지진파별 avg 값 계산
    # DE
    if len(DE_load_name_list) != 0:
        # 지진파별 Rotation avg 값을 저장할 dataframe
        beam_rot_data_DE = pd.DataFrame()
        # 지진파별로 for loop
        for load_name in DE_load_name_list:
            # 지진파별로 max, min 값 추출        
            df_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                   & (beam_rot_data['Step Type'] == 'Max')]                
            df_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                   & (beam_rot_data['Step Type'] == 'Min')]
            # max, min 값별로 Rotation avg 값 계산  
            df_max_avg = df_max.groupby(['Element Name', 'Property Name', 'V'])\
                         ['Major Rotation(rad)'].agg(**{'Rotation avg':'mean'})['Rotation avg']
            df_min_avg = df_min.groupby(['Element Name', 'Property Name', 'V'])\
                         ['Major Rotation(rad)'].agg(**{'Rotation avg':'mean'})['Rotation avg']
            # Rotation avg 값을 beam_rot_data_DE에 저장       
            beam_rot_data_DE['{}_max'.format(load_name)] = df_max_avg.tolist()                          
            beam_rot_data_DE['{}_min'.format(load_name)] = df_min_avg.tolist()
        # df_max_avg의 index에 assign된 beam 정보 추출
        beam_info = df_max_avg.index.to_frame()
        # beam_info에 있는 부재 이름을 beam_rot_data_DE에 match
        beam_rot_data_DE['Property Name'] = beam_info['Property Name'].tolist()        
        # 각 지진파에서의 Rotation에 대한 평균값 계산        
        beam_rot_data_DE['DE Max avg'] = beam_rot_data_DE.iloc[:,list(range(0,len(DE_load_name_list)*2,2))].mean(axis=1)
        beam_rot_data_DE['DE Min avg'] = beam_rot_data_DE.iloc[:,list(range(1,len(DE_load_name_list)*2,2))].mean(axis=1).abs()
        # max, min 값 중 큰 값 선택
        beam_rot_data_DE['DE avg'] = beam_rot_data_DE[['DE Max avg', 'DE Min avg']].max(axis=1)
          
        # Input sheet의 Performance Level 2 정보에 Rotation 값 match
        beam_plastic_hinge = pd.merge(deformation_cap_beam
                                      , beam_rot_data_DE[['Property Name', 'DE avg']], how='left')
       
    # MCE
    if len(MCE_load_name_list) != 0:
        # 지진파별 Rotation avg 값을 저장할 dataframe
        beam_rot_data_MCE = pd.DataFrame()
        # 지진파별로 for loop
        for load_name in MCE_load_name_list:
            # 지진파별로 max, min 값 추출        
            df_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                   & (beam_rot_data['Step Type'] == 'Max')]                
            df_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                   & (beam_rot_data['Step Type'] == 'Min')]
            # max, min 값별로 Rotation avg 값 계산  
            df_max_avg = df_max.groupby(['Element Name', 'Property Name', 'V'])\
                         ['Major Rotation(rad)'].agg(**{'Rotation avg':'mean'})['Rotation avg']
            df_min_avg = df_min.groupby(['Element Name', 'Property Name', 'V'])\
                         ['Major Rotation(rad)'].agg(**{'Rotation avg':'mean'})['Rotation avg']
            # Rotation avg 값을 beam_rot_data_MCE에 저장       
            beam_rot_data_MCE['{}_max'.format(load_name)] = df_max_avg.tolist()                          
            beam_rot_data_MCE['{}_min'.format(load_name)] = df_min_avg.tolist()
        # df_max_avg의 index에 assign된 beam 정보 추출
        beam_info = df_max_avg.index.to_frame()
        # beam_info에 있는 부재 이름을 beam_rot_data_MCE에 match
        beam_rot_data_MCE['Property Name'] = beam_info['Property Name'].tolist()        
        # 각 지진파에서의 Rotation에 대한 평균값 계산        
        beam_rot_data_MCE['MCE Max avg'] = beam_rot_data_MCE.iloc[:,list(range(0,len(MCE_load_name_list)*2,2))].mean(axis=1)
        beam_rot_data_MCE['MCE Min avg'] = beam_rot_data_MCE.iloc[:,list(range(1,len(MCE_load_name_list)*2,2))].mean(axis=1).abs()
        # max, min 값 중 큰 값 선택
        beam_rot_data_MCE['MCE avg'] = beam_rot_data_MCE[['MCE Max avg', 'MCE Min avg']].max(axis=1)
          
        # Input sheet의 Performance Level 2 정보에 Rotation 값 match
        beam_plastic_hinge = pd.merge(beam_plastic_hinge
                                      , beam_rot_data_MCE[['Property Name', 'MCE avg']], how='left')
        
        beam_plastic_hinge = beam_plastic_hinge[beam_plastic_hinge['Property Name'].notna()]
        beam_plastic_hinge = beam_plastic_hinge[beam_plastic_hinge['DE avg'].notna()]
        beam_plastic_hinge = beam_plastic_hinge[beam_plastic_hinge['MCE avg'].notna()]

#%% DE, MCE 각각의 load case에 대해 max, min 값 추출 / 지진파별 avg 값 계산
    # DE
    if len(DE_load_name_list) != 0:
        # 지진파별 Rotation avg 값을 저장할 dataframe
        col_rot_data_DE = pd.DataFrame()
        # 지진파별로 for loop
        for load_name in DE_load_name_list:
            # 지진파별로 max, min 값 추출        
            df_max = col_rot_data[(col_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                   & (col_rot_data['Step Type'] == 'Max')]                
            df_min = col_rot_data[(col_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                   & (col_rot_data['Step Type'] == 'Min')]
            # max, min 값별로 Rotation avg 값 계산  
            df_max_avg_x = df_max.groupby(['Element Name', 'Property Name', 'V'])\
                         ['H2 Rotation(rad)'].agg(**{'Rotation avg':'mean'})['Rotation avg']
            df_min_avg_x = df_min.groupby(['Element Name', 'Property Name', 'V'])\
                         ['H2 Rotation(rad)'].agg(**{'Rotation avg':'mean'})['Rotation avg']
            # Y 방향(H3 Rotation)에 대해서도
            df_max_avg_y = df_max.groupby(['Element Name', 'Property Name', 'V'])\
                         ['H3 Rotation(rad)'].agg(**{'Rotation avg':'mean'})['Rotation avg']
            df_min_avg_y = df_min.groupby(['Element Name', 'Property Name', 'V'])\
                         ['H3 Rotation(rad)'].agg(**{'Rotation avg':'mean'})['Rotation avg']
            # Rotation avg 값을 col_rot_data_DE에 저장       
            col_rot_data_DE['{}_max_x'.format(load_name)] = df_max_avg_x.tolist()                          
            col_rot_data_DE['{}_min_x'.format(load_name)] = df_min_avg_x.tolist()
            col_rot_data_DE['{}_max_y'.format(load_name)] = df_max_avg_y.tolist()                          
            col_rot_data_DE['{}_min_y'.format(load_name)] = df_min_avg_y.tolist()
        # df_max_avg의 index에 assign된 col 정보 추출
        col_info = df_max_avg_x.index.to_frame()
        # col_info에 있는 부재 이름을 col_rot_data_DE에 match
        col_rot_data_DE['Property Name'] = col_info['Property Name'].tolist()        
        # 각 지진파에서의 Rotation에 대한 평균값 계산        
        col_rot_data_DE['DE Max avg X'] = col_rot_data_DE.iloc[:,list(range(0,len(DE_load_name_list)*4,4))].mean(axis=1)
        col_rot_data_DE['DE Min avg X'] = col_rot_data_DE.iloc[:,list(range(1,len(DE_load_name_list)*4,4))].mean(axis=1).abs()
        col_rot_data_DE['DE Max avg Y'] = col_rot_data_DE.iloc[:,list(range(2,len(DE_load_name_list)*4,4))].mean(axis=1)
        col_rot_data_DE['DE Min avg Y'] = col_rot_data_DE.iloc[:,list(range(3,len(DE_load_name_list)*4,4))].mean(axis=1).abs()
        # max, min 값 중 큰 값 선택
        col_rot_data_DE['DE avg'] = col_rot_data_DE[['DE Max avg X', 'DE Min avg X'
                                                     , 'DE Max avg Y', 'DE Min avg Y']].max(axis=1)
          
        # Input sheet의 Performance Level 2 정보에 Rotation 값 match
        col_plastic_hinge = pd.merge(deformation_cap_col
                                      , col_rot_data_DE[['Property Name', 'DE avg']], how='left')
       
    # MCE
    if len(MCE_load_name_list) != 0:
        # 지진파별 Rotation avg 값을 저장할 dataframe
        col_rot_data_MCE = pd.DataFrame()
        # 지진파별로 for loop
        for load_name in MCE_load_name_list:
            # 지진파별로 max, min 값 추출        
            df_max = col_rot_data[(col_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                   & (col_rot_data['Step Type'] == 'Max')]                
            df_min = col_rot_data[(col_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                   & (col_rot_data['Step Type'] == 'Min')]
            # max, min 값별로 Rotation avg 값 계산  
            df_max_avg_x = df_max.groupby(['Element Name', 'Property Name', 'V'])\
                         ['H2 Rotation(rad)'].agg(**{'Rotation avg':'mean'})['Rotation avg']
            df_min_avg_x = df_min.groupby(['Element Name', 'Property Name', 'V'])\
                         ['H2 Rotation(rad)'].agg(**{'Rotation avg':'mean'})['Rotation avg']
            # Y 방향(H3 Rotation)에 대해서도
            df_max_avg_y = df_max.groupby(['Element Name', 'Property Name', 'V'])\
                         ['H3 Rotation(rad)'].agg(**{'Rotation avg':'mean'})['Rotation avg']
            df_min_avg_y = df_min.groupby(['Element Name', 'Property Name', 'V'])\
                         ['H3 Rotation(rad)'].agg(**{'Rotation avg':'mean'})['Rotation avg']
            # Rotation avg 값을 col_rot_data_MCE에 저장       
            col_rot_data_MCE['{}_max_x'.format(load_name)] = df_max_avg_x.tolist()                          
            col_rot_data_MCE['{}_min_x'.format(load_name)] = df_min_avg_x.tolist()
            col_rot_data_MCE['{}_max_y'.format(load_name)] = df_max_avg_y.tolist()                          
            col_rot_data_MCE['{}_min_y'.format(load_name)] = df_min_avg_y.tolist()
        # df_max_avg의 index에 assign된 col 정보 추출
        col_info = df_max_avg_x.index.to_frame()
        # col_info에 있는 부재 이름을 col_rot_data_MCE에 match
        col_rot_data_MCE['Property Name'] = col_info['Property Name'].tolist()        
        # 각 지진파에서의 Rotation에 대한 평균값 계산        
        col_rot_data_MCE['MCE Max avg X'] = col_rot_data_MCE.iloc[:,list(range(0,len(MCE_load_name_list)*4,4))].mean(axis=1)
        col_rot_data_MCE['MCE Min avg X'] = col_rot_data_MCE.iloc[:,list(range(1,len(MCE_load_name_list)*4,4))].mean(axis=1).abs()
        col_rot_data_MCE['MCE Max avg Y'] = col_rot_data_MCE.iloc[:,list(range(2,len(MCE_load_name_list)*4,4))].mean(axis=1)
        col_rot_data_MCE['MCE Min avg Y'] = col_rot_data_MCE.iloc[:,list(range(3,len(MCE_load_name_list)*4,4))].mean(axis=1).abs()
        # max, min 값 중 큰 값 선택
        col_rot_data_MCE['MCE avg'] = col_rot_data_MCE[['MCE Max avg X', 'MCE Min avg X'
                                                     , 'MCE Max avg Y', 'MCE Min avg Y']].max(axis=1)
          
        # Input sheet의 Performance Level 2 정보에 Rotation 값 match
        col_plastic_hinge = pd.merge(col_plastic_hinge
                                      , col_rot_data_MCE[['Property Name', 'MCE avg']], how='left')
        
        col_plastic_hinge = col_plastic_hinge[col_plastic_hinge['Property Name'].notna()]
        col_plastic_hinge = col_plastic_hinge[col_plastic_hinge['DE avg'].notna()]
        col_plastic_hinge = col_plastic_hinge[col_plastic_hinge['MCE avg'].notna()]
    
#%% 엑셀로 출력(Using win32com)
    
    # Using win32com...
    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게
    
    wb = excel.Workbooks.Open(input_xlsx_path)
    ws1 = wb.Sheets('Results_C.Beam')
    ws2 = wb.Sheets('Results_G.Column')
    
    startrow, startcol = 5, 1
    
    # C.Beam의 소성회전각(Performance Level 2), 회전각(Rotation) 입력
    # ws1.Range('AE%s:AG%s' %(startrow, startrow+beam_plastic_hinge.shape[0]-1)).Value\
    # = list(beam_plastic_hinge.iloc[:,[1,2,3]].itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능
    
    # G.Column의 소성회전각(Performance Level 2), 회전각(Rotation) 입력
    # ws2.Range('AI%s:AK%s' %(startrow, startrow+col_plastic_hinge.shape[0]-1)).Value\
    # = list(col_plastic_hinge.iloc[:,[1,2,3]].itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능
    
    # (docx 출력을 위해) C.Beam의 회전각, DCR, 소성힌지/ 부재 정보 읽기
    beam_result_output = pd.DataFrame(ws1.Range('AE%s:AJ%s' %(startrow, startrow+beam_plastic_hinge.shape[0]-1)).Value)
    beam_info_output_1 = pd.DataFrame(ws1.Range('A%s:A%s' %(startrow, startrow+beam_plastic_hinge.shape[0]-1)).Value)
    beam_info_output_2 = pd.DataFrame(ws1.Range('C%s:D%s' %(startrow, startrow+beam_plastic_hinge.shape[0]-1)).Value)
    beam_info_output_3 = pd.DataFrame(ws1.Range('J%s:K%s' %(startrow, startrow+beam_plastic_hinge.shape[0]-1)).Value)
    beam_info_output_4 = pd.DataFrame(ws1.Range('M%s:N%s' %(startrow, startrow+beam_plastic_hinge.shape[0]-1)).Value)
    beam_info_output_5 = pd.DataFrame(ws1.Range('P%s:Q%s' %(startrow, startrow+beam_plastic_hinge.shape[0]-1)).Value)
    
    # Dataframe 정리    
    # dataframe 생성 후, 이름 분리(벽체 이름, 번호, 층)
    beam_output = pd.DataFrame()
    beam_output['Property Name'] = beam_info_output_1.iloc[:,0].str.split('_', expand=True)[0]
    beam_output['Number'] = beam_info_output_1.iloc[:,0].str.split('_', expand=True)[1]
    beam_output['Story'] = beam_info_output_1.iloc[:,0].str.split('_', expand=True)[2]

    # width와 height 정보 합치기 (geometry)
    beam_info_output_2 = beam_info_output_2.astype(int) # 데이터프레임에 있는 숫자의 default=float
    beam_output['Geometry'] = beam_info_output_2.iloc[:,0].astype(str) + ' X ' + beam_info_output_2.iloc[:,1].astype(str)
    
    # Top,Bot,Stirrup Rebar 정보 합치고 정리하기
    beam_info_output_4 = beam_info_output_4.astype(int, errors='ignore') # Bot Bar의 셀이 비어있는 경우, 에러 무시하고 ''값 그대로.
    beam_info_output_5 = beam_info_output_5.astype(int)
    beam_output['Top Bar'] = beam_info_output_4.iloc[:,0].astype(str) + '-' + beam_info_output_3.iloc[:,0]
    beam_output['Stirrup'] = beam_info_output_5.iloc[:,0].astype(str) + '-' + beam_info_output_3.iloc[:,1]\
        + '@' + beam_info_output_5.iloc[:,1].astype(str)
        
    # Bot Bar의 경우, 빈 셀일 수 있으므로, 별도 처리
    # beam_info_output_4.iloc[:,1] = [int(i) for i in beam_info_output_4.iloc[:,1] if (i != '') | (i != '0')]
    bot_bar = beam_info_output_4.iloc[:,1].astype(str) + '-' + beam_info_output_3.iloc[:,0]
    bot_bar[bot_bar.str.startswith(('-', '0'))] = '' # '-' 또는 '0'로 시작하는 셀은 빈 칸으로 두기
    beam_output['Bot Bar'] = bot_bar
        
    # Rotation 및 소성힌지 정보 합치기
    beam_result_output.columns = ['Plastic Rotational Capacity', 'Rotation(DE)'
                                  , 'Rotation(MCE)', 'DCR(DE)', 'DCR(MCE)', 'Plastic Hinge']
    beam_output = pd.concat([beam_output, beam_result_output], axis=1)    
    # 반올림(5쨰 자리)
    beam_output[['Plastic Rotational Capacity', 'Rotation(DE)', 'Rotation(MCE)', 'DCR(DE)', 'DCR(MCE)']]\
        = beam_output[['Plastic Rotational Capacity', 'Rotation(DE)', 'Rotation(MCE)', 'DCR(DE)', 'DCR(MCE)']].round(5)

    # 보 이름과 번호(W1_1)이 같은 부재들끼리 groupby로 묶고, list of dataframes 생성
    beam_output_list = list(beam_output.groupby(['Property Name', 'Number'], sort=False))
    
    
    # (docx 출력을 위해) G.Column의 회전각, DCR, 소성힌지 여부 읽기
    col_result_output = pd.DataFrame(ws2.Range('AI%s:AN%s' %(startrow, startrow+col_plastic_hinge.shape[0]-1)).Value)
    col_info_output_1 = pd.DataFrame(ws2.Range('A%s:A%s' %(startrow, startrow+col_plastic_hinge.shape[0]-1)).Value)
    col_info_output_2 = pd.DataFrame(ws2.Range('B%s:C%s' %(startrow, startrow+col_plastic_hinge.shape[0]-1)).Value)
    col_info_output_3 = pd.DataFrame(ws2.Range('G%s:G%s' %(startrow, startrow+col_plastic_hinge.shape[0]-1)).Value)
    col_info_output_4 = pd.DataFrame(ws2.Range('I%s:I%s' %(startrow, startrow+col_plastic_hinge.shape[0]-1)).Value)
    col_info_output_5 = pd.DataFrame(ws2.Range('J%s:J%s' %(startrow, startrow+col_plastic_hinge.shape[0]-1)).Value)
    col_info_output_6 = pd.DataFrame(ws2.Range('L%s:L%s' %(startrow, startrow+col_plastic_hinge.shape[0]-1)).Value)
    col_info_output_7 = pd.DataFrame(ws2.Range('P%s:P%s' %(startrow, startrow+col_plastic_hinge.shape[0]-1)).Value)
    
    # Dataframe 정리    
    # dataframe 생성 후, 이름 분리(벽체 이름, 번호, 층)
    col_output = pd.DataFrame()
    col_output['Property Name'] = col_info_output_1.iloc[:,0].str.split('_', expand=True)[0]
    col_output['Number'] = col_info_output_1.iloc[:,0].str.split('_', expand=True)[1]
    col_output['Story'] = col_info_output_1.iloc[:,0].str.split('_', expand=True)[2]

    # width와 height 정보 합치기 (geometry)
    col_info_output_2 = col_info_output_2.astype(int) # 데이터프레임에 있는 숫자의 default=float
    col_output['Geometry'] = col_info_output_2.iloc[:,0].astype(str) + ' X ' + col_info_output_2.iloc[:,1].astype(str)
    
    # Main,Hoop Rebar 정보 합치고 정리하기
    col_info_output_5 = col_info_output_5.astype(int)    
    col_info_output_7 = col_info_output_7.astype(int)
    col_output['Main Bar-1'] = col_info_output_5.iloc[:,0].astype(str) + '-' + col_info_output_3.iloc[:,0]    
    col_output['Hoop'] = col_info_output_4.iloc[:,0] + '@' + col_info_output_7.iloc[:,0].astype(str)
    
    # Main Bar-2의 경우, 빈 셀일 수 있으므로, 별도 처리
    col_info_output_6 = col_info_output_6.astype(int, errors='ignore') # Layer2의 셀이 비어있는 경우, 에러 무시하고 ''값 그대로. 
    main_bar_2 = col_info_output_6.iloc[:,0].astype(str) + '-' + col_info_output_3.iloc[:,0]
    main_bar_2[main_bar_2.str.startswith(('-', '0'))] = '' # '-' 또는 '0'로 시작하는 셀은 빈 칸으로 두기
    col_output['Main Bar-2'] = main_bar_2
    
    # Rotation 및 소성힌지 정보 합치기
    col_result_output.columns = ['Plastic Rotational Capacity', 'Rotation(DE)'
                                  , 'Rotation(MCE)', 'DCR(DE)', 'DCR(MCE)', 'Plastic Hinge']
    col_output = pd.concat([col_output, col_result_output], axis=1)
    # 반올림(5쨰 자리)
    col_output[['Plastic Rotational Capacity', 'Rotation(DE)', 'Rotation(MCE)', 'DCR(DE)', 'DCR(MCE)']]\
        = col_output[['Plastic Rotational Capacity', 'Rotation(DE)', 'Rotation(MCE)', 'DCR(DE)', 'DCR(MCE)']].round(5)

    # 기둥 이름과 번호(W1_1)이 같은 부재들끼리 groupby로 묶고, list of dataframes 생성
    col_output_list = list(col_output.groupby(['Property Name', 'Number'], sort=False))
    
    # return beam_output_list, col_output_list
    
    
    wb.Save()
    # wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application 

#%% C.Beam SF - 허무원 박사
def BSF_HMW(input_xlsx_path, result_xlsx_path):
    ''' 

    Perform-3D 해석 결과에서 일반기둥의 축력, 전단력을 불러와 Results_G.Column 엑셀파일을 작성. \n
    result_path : Perform-3D에서 나온 해석 파일의 경로. \n
    result_xlsx : Perform-3D에서 나온 해석 파일의 이름. 해당 파일 이름이 포함된 파일들을 모두 불러온다. \n
    input_path : Data Conversion 엑셀 파일의 경로 \n
    input_xlsx : Data Conversion 엑셀 파일의 이름. result_xlsx와는 달리 확장자명(.xlsx)까지 기입해줘야한다. 하나의 파일만 불러온다. \n
    column_xlsx : Results_E.Column 엑셀 파일의 이름.확장자명(.xlsx)까지 기입해줘야한다. \n
    export_to_pdf : 입력된 값에 따른 각 부재들의 결과 시트를 pdf로 출력. True = pdf 출력, False = pdf 미출력(Results_E.Column 엑셀파일만 작성됨).
    pdf_name = 출력할 pdf 파일 이름.
    
    '''
#%% Input Sheet 정보 load
        
    story_info = pd.DataFrame()
    element_name = pd.DataFrame()

    input_xlsx_sheet = 'Output_C.Beam Properties'
    input_data_raw = pd.ExcelFile(input_xlsx_path)
    input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', input_xlsx_sheet], skiprows=3)
    input_data_raw.close()

    story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
    element_name = input_data_sheets[input_xlsx_sheet].iloc[:,0]

    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    element_name.name = 'Property Name'

#%% Analysis Result 불러오기
    to_load_list = result_xlsx_path

    # 전단력 Data
    SF_info_data = pd.DataFrame()
    for i in to_load_list:
        result_data_raw = pd.ExcelFile(i)
        result_data_sheets = pd.read_excel(result_data_raw, ['Frame Results - End Forces'
                                           , 'Node Coordinate Data', 'Element Data - Frame Types']
                                          , skiprows=[0, 2]) # usecols로 원하는 열만 불러오기
        
        SF_info_data_temp = result_data_sheets['Frame Results - End Forces'].iloc[:,[0,2,5,7,10,12]]
        SF_info_data = pd.concat([SF_info_data, SF_info_data_temp])

    node_data = result_data_sheets['Node Coordinate Data'].iloc[:,[1,4]]
    element_data = result_data_sheets['Element Data - Frame Types'].iloc[:,[0,2,5,7]] # beam의 양 nodes중 한 node에서의 rotation * 2

    # 필요한 부재만 선별
    # 필요한 부재만 선별
    element_data = element_data[element_data['Group Name'] == 'BEAM']
    element_data['Property Name'] = element_data['Property Name'] + '_1_'

#%% element 이름 재명명(101동 부재 섞어서 쓰심)     ########## 허무원 ##########
    node_data_101 = result_data_sheets['Node Coordinate Data'].iloc[:,[1,2,3,4]]
    element_data_101 = pd.merge(element_data, node_data_101, how='left', left_on='I-Node ID', right_on='Node ID')
    
    list_101 = []    
    
    # for idx, row in element_data_101.iterrows():
    #     if (row['Property Name'] == 'LB4_1_') & (row['H1'] == 2172.5):
    #         list_101.append('LB104_1_')
    #     elif (row['Property Name'] == 'LB4_1_') & (row['H2'] == -930.5):
    #         list_101.append('LB5_1_')
    #     elif (row['Property Name'] == 'LB7_1_') & (row['H1'] == 1982):
    #         list_101.append('LB101_1_')
    #     elif (row['Property Name'] == 'LB102_1_') & (row['H2'] == -465):
    #         list_101.append('LB103_1_')
    #     else:    
    #         list_101.append(row['Property Name'])
    
    for idx, row in element_data_101.iterrows():
        if (row['Property Name'] == 'LB4LB5LB104_1_'):
            list_101.append('LB104_1_')
        elif (row['Property Name'] == 'LB102LB103_1_'):
            list_101.append('LB103_1_')
        elif (row['Property Name'] == 'LB7LB101_1_'):
            list_101.append('LB101_1_')

        else:    
            list_101.append(row['Property Name'])
            
    element_data['Property Name'] = list_101

#%% Analysis Result에 Element, Node 정보 매칭

    element_data = element_data.drop_duplicates()
    
    element_data = pd.merge(element_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
    SF_ongoing = pd.merge(element_data.iloc[:, [1,2,5]], SF_info_data.iloc[:, 1:], how='left')
    SF_ongoing.reset_index(inplace=True, drop=True)
    
#%% 허무원    
    # 이름에 층정보 붙이기
    SF_ongoing_copy = pd.merge(SF_ongoing, story_info, how='left', left_on = 'V', right_on = 'Height(mm)')
    new_name = SF_ongoing_copy['Property Name'] + SF_ongoing_copy['Story Name']
    SF_ongoing['Property Name'] = new_name    

#%% 지진파 이름 list 만들기 ########## 허무원 ##########

    load_name_list = []
    for i in SF_ongoing['Load Case'].drop_duplicates():
        new_i = i.split('+')[1]
        new_i = new_i.strip()
        load_name_list.append(new_i)

    gravity_load_name = [x for x in load_name_list if 'GL' in x]
    seismic_load_name_list = [x for x in load_name_list if 'GL' not in x]

    seismic_load_name_list.sort()

    DE_load_name_list = [x for x in load_name_list if ('GL' not in x) & ('MCE' not in x)]
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]

#%% V, M값에 절대값, 최대값, 평균값 뽑기

    # 절대값
    SF_ongoing.iloc[:,5] = SF_ongoing.iloc[:,5].abs()

    # V2의 최대값을 저장하기 위해 필요한 데이터 slice
    SF_ongoing_max = SF_ongoing.iloc[[2*x for x in range(int(SF_ongoing.shape[0]/2))],[0,1,2,3]] 
    # [2*x for x in range(int(SF_ongoing.shape[0]/2] -> [짝수 index]
    
    # V2, V3의 최대값을 저장
    SF_ongoing_max['V2 max'] = SF_ongoing.groupby(SF_ongoing.index // 2)['V2 I-End'].max().tolist()

    # 필요한 하중만 포함된 데이터 slice (MCE)
    SF_ongoing_max_MCE = SF_ongoing_max[SF_ongoing_max['Load Case']\
                                        .str.contains('|'.join(MCE_load_name_list))]
    SF_ongoing_max_G = SF_ongoing_max[SF_ongoing_max['Load Case']\
                                      .str.contains('|'.join(gravity_load_name))]
    # function equivalent of a combination of df.isin() and df.str.contains()
    
    # 부재별(Element Name) 평균값을 저장하기 위해 필요한 데이터프레임 생성
    SF_ongoing_max_avg = SF_ongoing_max_MCE.iloc[:,[0,1,2]]
    SF_ongoing_max_avg = SF_ongoing_max_avg.drop_duplicates()
    SF_ongoing_max_avg.set_index('Element Name', inplace=True)    
    # 부재별(Element Name) 평균값 뽑기
    SF_ongoing_max_avg['V2 max(MCE)'] = SF_ongoing_max_MCE.groupby(['Element Name'])['V2 max'].mean()
    SF_ongoing_max_avg['V2 max(G)'] = SF_ongoing_max_G.groupby(['Element Name'])['V2 max'].mean()
    
    # 이름별(Property Name) 최대값을 저장하기 위해 필요한 데이터프레임 생성
    SF_ongoing_max_avg_max = SF_ongoing_max_avg.copy()
    SF_ongoing_max_avg_max = SF_ongoing_max_avg_max.drop_duplicates(subset=['Property Name'], ignore_index=True)
    SF_ongoing_max_avg_max.set_index('Property Name', inplace=True) 
    # 같은 부재(그러나 잘려있는) 경우(Property Name) 최대값 뽑기
    SF_ongoing_max_avg_max = pd.merge(SF_ongoing_max_avg_max
                                      , SF_ongoing_max_avg.groupby(['Property Name'])['V2 max(MCE)'].max()
                                      , left_on='Property Name', right_index=True, suffixes=('_before', '_after'))
    SF_ongoing_max_avg_max = pd.merge(SF_ongoing_max_avg_max
                                      , SF_ongoing_max_avg.groupby(['Property Name'])['V2 max(G)'].max()
                                      , left_on='Property Name', right_index=True, suffixes=('_before', '_after'))
    
    # MCE에 대해 1.2배, G에 대해 0.2배
    SF_ongoing_max_avg_max['V2 max(MCE)_after'] = SF_ongoing_max_avg_max['V2 max(MCE)_after'] * 1.2
    SF_ongoing_max_avg_max['V2 max(G)_after'] = SF_ongoing_max_avg_max['V2 max(G)_after'] * 0.2
    
    SF_ongoing_max_avg_max.reset_index(inplace=True, drop=False)

#%% 결과값 정리
    
    SF_output = pd.merge(element_name, SF_ongoing_max_avg_max, how='left')
        
    SF_output = SF_output.dropna()
    SF_output.reset_index(inplace=True, drop=True)
        
    # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
    SF_output = SF_output.replace(np.nan, '', regex=True)
    
    # 기존 시트에 V값 넣기
    SF_output1 = SF_output.iloc[:,0]
    SF_output2 = SF_output.iloc[:,[4,5]]

#%% 출력 (Using win32com...)
    
    # Using win32com...
    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게

    wb = excel.Workbooks.Open(input_xlsx_path)
    ws = wb.Sheets('Results_C.Beam')
    
    startrow, startcol = 5, 1    
    ws.Range(ws.Cells(startrow, startcol),\
              ws.Cells(startrow+SF_output1.shape[0]-1, startcol)).Value\
    = [[i] for i in SF_output1]
    
    startrow, startcol = 5, 20    
    ws.Range(ws.Cells(startrow, startcol),\
              ws.Cells(startrow+SF_output2.shape[0]-1,\
                      startcol+SF_output2.shape[1]-1)).Value\
    = list(SF_output2.itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능
    
    wb.Save()            
    # wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application
    
#%% Beam Rotation - 허무원 박사
def BR_HMW(input_xlsx_path, result_xlsx_path
           , c_beam_group='C.Beam', DCR_criteria=1, yticks=3, xlim=3):

#%% Input Sheets 정보 load
    story_info = pd.DataFrame()
    deformation_cap = pd.DataFrame()
    
    input_data_raw = pd.ExcelFile(input_xlsx_path)
    input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', 'Output_C.Beam Properties'], skiprows=3)
    input_data_raw.close()
    
    story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
    deformation_cap = input_data_sheets['Output_C.Beam Properties'].iloc[:,[0,80,81]]
    
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    deformation_cap.columns = ['Name', 'LS', 'CP']
    
#%% Analysis Result 불러오기
    to_load_list = result_xlsx_path
    
    beam_rot_data = pd.DataFrame()
    
    for i in to_load_list:
        result_data_raw = pd.ExcelFile(i)
        result_data_sheets = pd.read_excel(result_data_raw, ['Frame Results - Bending Deform', 'Node Coordinate Data',\
                                                         'Element Data - Frame Types'], skiprows=[0,2])
        
        beam_rot_data_temp = result_data_sheets['Frame Results - Bending Deform'].iloc[:,[0,2,5,7,10,13,14]]
        beam_rot_data = pd.concat([beam_rot_data, beam_rot_data_temp])
        
    node_data = result_data_sheets['Node Coordinate Data'].iloc[:,[1,4]]
    
    element_data = result_data_sheets['Element Data - Frame Types'].iloc[:,[2,5,7]] # beam의 양 nodes중 한 node에서의 rotation * 2
    
                
    beam_rot_data.columns = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'Distance from I-End', 'H2 Rotation(rad)', 'H3 Rotation(rad)']
    node_data.columns = ['Node ID', 'V']
    element_data.columns = ['Element Name', 'Property Name', 'I-Node ID']
    
    #%% 필요없는 부재 빼기, 필요한 부재만 추출
    beam_rot_data = beam_rot_data[beam_rot_data['Group Name'] == beam_group]
    beam_rot_data = beam_rot_data[beam_rot_data['Distance from I-End'] == 0]
    
#%% element 이름 재명명(101동 부재 섞어서 쓰심)     ########## 허무원 ##########
    element_data['Property Name'] = element_data['Property Name'] + '_1_'
    
    node_data_101 = result_data_sheets['Node Coordinate Data'].iloc[:,[1,2,3,4]]
    element_data_101 = pd.merge(element_data, node_data_101, how='left', left_on='I-Node ID', right_on='Node ID')
    
    list_101 = []    
    
    # for idx, row in element_data_101.iterrows():
    #     if (row['Property Name'] == 'LB4_1_') & (row['H1'] == 2172.5):
    #         list_101.append('LB104_1_')
    #     elif (row['Property Name'] == 'LB4_1_') & (row['H2'] == -930.5):
    #         list_101.append('LB5_1_')
    #     elif (row['Property Name'] == 'LB7_1_') & (row['H1'] == 1982):
    #         list_101.append('LB101_1_')
    #     elif (row['Property Name'] == 'LB102_1_') & (row['H2'] == -465):
    #         list_101.append('LB103_1_')
    #     else:    
    #         list_101.append(row['Property Name'])
    
    for idx, row in element_data_101.iterrows():
        if (row['Property Name'] == 'LB2LB3_1_'):
            list_101.append('LB3_1_')
        elif (row['Property Name'] == 'LB101LB105_1_'):
            list_101.append('LB101_1_')
        elif (row['Property Name'] == 'LB103LB104_1_'):
            list_101.append('LB103_1_')

        else:    
            list_101.append(row['Property Name'])
            
    element_data['Property Name'] = list_101
    
    #%% Analysis Result에 Element, Node 정보 매칭    
    element_data = element_data.drop_duplicates()
    node_data = node_data.drop_duplicates()
    
    beam_rot_data = pd.merge(beam_rot_data, element_data, how='left')
    beam_rot_data = pd.merge(beam_rot_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
    
    beam_rot_data = beam_rot_data[beam_rot_data['Property Name'].notna()]
    
    beam_rot_data.reset_index(inplace=True, drop=True)
    
    
    # 이름에 층정보 붙이기
    beam_rot_data_copy = pd.merge(beam_rot_data, story_info, how='left', left_on = 'V', right_on = 'Height(mm)')
    new_name = beam_rot_data_copy['Property Name'] + beam_rot_data_copy['Story Name']
    beam_rot_data['Property Name'] = new_name  

#%% 지진파 이름 list 만들기 ########## 허무원 ##########

    ################## 허무원 박사님용 지진파 이름 변경 #########################
    existing = list(range(14,0,-1)) + ['MCE-14', 'MCE-13', 'MCE-12', 'MCE-11'
                                       , 'MCE-10', 'MCE-09', 'MCE-08', 'MCE-07'
                                       , 'MCE-06', 'MCE-05', 'MCE-04', 'MCE-03'
                                       , 'MCE-02', 'MCE-01']
    renewed = ['DE72', 'DE71', 'DE62', 'DE61', 'DE52', 'DE51', 'DE42', 'DE41'
               , 'DE32', 'DE31', 'DE22', 'DE21', 'DE12', 'DE11', 'MCE72', 'MCE71'
               , 'MCE62', 'MCE61', 'MCE52', 'MCE51', 'MCE42', 'MCE41', 'MCE32'
               , 'MCE31', 'MCE22', 'MCE21', 'MCE12', 'MCE11']
    for i, j in zip(existing, renewed):
        beam_rot_data['Load Case'] = beam_rot_data['Load Case'].str.replace('[1] + %s'%i, '[1] + %s'%j, regex=False)
    ###########################################################################

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
    
#%% beam_rot_data의 값 수정(H1, H2 방향 중 major한 방향의 rotation값만 추출, 그리고 2배)
    major_rot = []
    for i, j in zip(beam_rot_data['H2 Rotation(rad)'], beam_rot_data['H3 Rotation(rad)']):
        if abs(i) >= abs(j):
            major_rot.append(i)
        else: major_rot.append(j)
    
    beam_rot_data['Major Rotation(rad)'] = major_rot
     
    # 필요한 정보들만 다시 모아서 new dataframe
    beam_rot_data = beam_rot_data.iloc[:, [0,1,2,3,7,9,10,11]]
    
#%% 성능기준(LS, CP) 정리해서 merge
    
    beam_rot_data = pd.merge(beam_rot_data, deformation_cap, how='left', left_on='Property Name', right_on='Name')
    
    beam_rot_data['DE Rotation(rad)'] = beam_rot_data['Major Rotation(rad)'].abs() / beam_rot_data['LS']
    beam_rot_data['MCE Rotation(rad)'] = beam_rot_data['Major Rotation(rad)'].abs() / beam_rot_data['CP']
    
    beam_rot_data = beam_rot_data[beam_rot_data['Name'].notna()]
    
    # beam_rot_data = pd.merge(deformation_cap['Name'], beam_rot_data, how='left', left_on='Name', right_on='Property Name')
    
#%% 조작용 코드
    # 없애고 싶은 부재의 이름 입력(error_beam 확인 후!, DE, MCE에서 다 없어짐)
    beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('PB'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('PB1-8_1'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('LB1A_2'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('LB1A_4'))].index)
    beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('LB101_1'))].index)
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
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, story_info, how='left', left_on='V', right_on='Height(mm)')
        beam_rot_data_total_DE.sort_values('Height(mm)', inplace=True)
        # beam_rot_data_total_DE.reset_index(inplace=True, drop=True)
        
    # 평균 열 생성
        
        beam_rot_data_total_DE['DE Max avg'] = beam_rot_data_total_DE.iloc[:,list(range(0,len(DE_load_name_list)*2,2))].mean(axis=1)
        beam_rot_data_total_DE['DE Min avg'] = beam_rot_data_total_DE.iloc[:,list(range(1,len(DE_load_name_list)*2,2))].mean(axis=1)
        
    # 전체 Plot
            
        ### DE 
        fig1 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        plt.scatter(beam_rot_data_total_DE['DE Max avg'], beam_rot_data_total_DE.loc[:,'V'], color='k', s=1)
        plt.scatter(beam_rot_data_total_DE['DE Min avg'], beam_rot_data_total_DE.loc[:,'V'], color='k', s=1)
        
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
        yield 'DE' # Marker 출력
        
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
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, story_info, how='left', left_on='V', right_on='Height(mm)')
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
        plt.scatter(beam_rot_data_total_MCE['MCE Max avg'], beam_rot_data_total_MCE.loc[:,'V'], color='k', s=1)
        plt.scatter(beam_rot_data_total_MCE['MCE Min avg'], beam_rot_data_total_MCE.loc[:,'V'], color='k', s=1)
        
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
        yield 'MCE' # Marker 출력