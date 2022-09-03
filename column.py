import pandas as pd
import numpy as np
import os
from collections import deque # Double-ended Queue : 자료의 앞, 뒤 양 방향에서 자료를 추가하거나 제거가능
# import openpyxl
import win32com.client

#%% Transfer Column SF (DCR)

def trans_column_SF(result_path, result_xlsx, input_path, input_xlsx, column_xlsx\
                    , export_to_pdf=True, pdf_name='Transfer Column Results'):
    ''' 

    Perform-3D 해석 결과에서 기둥의 축력, 전단력, 모멘트를 불러와 Results_E.Column 엑셀파일을 작성. \n
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
    transfer_element_name = pd.DataFrame()

    input_xlsx_sheet = 'Output_E.Column Properties'
    input_data_raw = pd.ExcelFile(input_path + '\\' + input_xlsx)
    input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', 'ETC', input_xlsx_sheet], skiprows=3)

    story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
    transfer_element_info = input_data_sheets[input_xlsx_sheet].iloc[:,0:17]
    rebar_info = input_data_sheets['ETC'].iloc[:,[0,3,4]]
    
    story_info = story_info[::-1]
    story_info.reset_index(inplace=True, drop=True)

    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    rebar_info.columns = ['Type', '일반용', '내진용']

#%% Result_E.Column에 입력될 Input sheets의 내용들 정리하기
    
    transfer_element_info.iloc[:,3] = transfer_element_info.iloc[:,3].fillna(method='ffill')

    transfer_element_info.columns = ['Name', 'b(mm)', 'h(mm)', 'c(mm)', 'Concrete Grade', 'Main Bar Type', 'Main Bar Diameter',\
                                     'Hoop Bar Type', 'Hoop Bar Diameter', 'Layer 1 EA', 'Layer 1 Row', 'Layer 2 EA',\
                                     'Layer 2 Row', 'Hoop X', 'Hoop Y', 'Hoop Spacing(mm)', 'Direction']
    
    # 철근 강도 불러오기
    main_bar_info = transfer_element_info.iloc[:,[5,6]]
    hoop_bar_info = transfer_element_info.iloc[:,[7,8]]
    
    main_bar_info = pd.merge(main_bar_info, rebar_info,\
                                           how='left', left_on='Main Bar Diameter', right_on='Type')
    hoop_bar_info = pd.merge(hoop_bar_info, rebar_info,\
                                           how='left', left_on='Hoop Bar Diameter', right_on='Type')
    
    # Main Bar 강도 리스트 만들기
    main_bar_strength = []
    for idx, row in main_bar_info.iterrows():
        if row[0] == '일반용':
            main_bar_strength.append(row[3])
        elif row[0] == '내진용':
            main_bar_strength.append(row[4])
    
    # Hoop Bar 강도 리스트 만들기
    hoop_bar_strength = []
    for idx, row in hoop_bar_info.iterrows():
        if row[0] == '일반용':
            hoop_bar_strength.append(row[3])
        elif row[0] == '내진용':
            hoop_bar_strength.append(row[4])
    
    transfer_element_info['Main Bar Strength'] = main_bar_strength
    transfer_element_info['Hoop Bar Strength'] = hoop_bar_strength  
    
    # 부재 이름 리스트    
    transfer_element_name = input_data_sheets[input_xlsx_sheet].iloc[:,0]

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
                                   sheet_name='Frame Results - End Forces', skiprows=[0, 2], header=0, usecols=[0,2,5,7,8,10,12,15,16,17,18]) # usecols로 원하는 열만 불러오기
        SF_info_data = pd.concat([SF_info_data, SF_info_data_temp])

    SF_info_data = SF_info_data.sort_values(by=['Load Case', 'Element Name', 'Step Type']) # 지진파 순서가 섞여있을 때 sort

    # 부재 이름 Matching을 위한 Element 정보
    element_info_data = pd.DataFrame()
    for i in to_load_list:
        element_info_data_temp = pd.read_excel(result_path + '\\' + i,
                                   sheet_name='Element Data - Frame Types', skiprows=[0, 2], header=0, usecols=[0, 2, 5, 7]) # usecols로 원하는 열만 불러오기
        element_info_data = pd.concat([element_info_data, element_info_data_temp])

    # 필요한 부재만 선별
    element_info_data = element_info_data[element_info_data['Property Name'].isin(transfer_element_name)]
    
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
    SF_ongoing.iloc[:,[5,6,7,8,9,10,11]] = SF_ongoing.iloc[:,[5,6,7,8,9,10,11]].abs() * 1.2

    # i, j 노드 중 최대값 뽑기
    SF_ongoing['M2 max'] = SF_ongoing[['M2 I-End', 'M2 J-End']].max(axis = 1)
    SF_ongoing['M3 max'] = SF_ongoing[['M3 I-End', 'M3 J-End']].max(axis = 1)

    # max, min 중 최대값 뽑기
    SF_ongoing_max = SF_ongoing.loc[SF_ongoing.groupby(SF_ongoing.index // 2)['P I-End'].idxmax()]
    SF_ongoing_max['V2 max'] = SF_ongoing.groupby(SF_ongoing.index // 2)['V2 I-End'].max().tolist()
    SF_ongoing_max['V3 max'] = SF_ongoing.groupby(SF_ongoing.index // 2)['V3 I-End'].max().tolist()
    SF_ongoing_max['M2 max'] = SF_ongoing.groupby(SF_ongoing.index // 2)['M2 max'].max().tolist()
    SF_ongoing_max['M3 max'] = SF_ongoing.groupby(SF_ongoing.index // 2)['M3 max'].max().tolist()

    # 필요한 하중만 포함된 데이터 slice (MCE)
    SF_ongoing_max = SF_ongoing_max[SF_ongoing_max['Load Case']\
                                    .str.contains('|'.join(MCE_load_name_list))] # function equivalent of a combination of df.isin() and df.str.contains()
    
    # 부재별 평균값 뽑기
    SF_ongoing_max_avg = SF_ongoing_max.iloc[:,[0,1,2]]
    SF_ongoing_max_avg = SF_ongoing_max_avg.drop_duplicates()
    SF_ongoing_max_avg.set_index('Element Name', inplace=True)
    
    SF_ongoing_max_avg['P'] = SF_ongoing_max.groupby(['Element Name'])['P I-End'].mean()
    SF_ongoing_max_avg['V2 max'] = SF_ongoing_max.groupby(['Element Name'])['V2 max'].mean()
    SF_ongoing_max_avg['V3 max'] = SF_ongoing_max.groupby(['Element Name'])['V3 max'].mean()
    SF_ongoing_max_avg['M2 max'] = SF_ongoing_max.groupby(['Element Name'])['M2 max'].mean()
    SF_ongoing_max_avg['M3 max'] = SF_ongoing_max.groupby(['Element Name'])['M3 max'].mean()
 
    # 같은 부재(그러나 잘려있는) 경우 최대값 뽑기
    SF_ongoing_max_avg_max = SF_ongoing_max_avg.loc[SF_ongoing_max_avg.groupby(['Property Name'])['P'].idxmax()]
    SF_ongoing_max_avg_max['V2 max'] = SF_ongoing_max_avg.groupby(['Property Name'])['V2 max'].max().tolist()
    SF_ongoing_max_avg_max['V3 max'] = SF_ongoing_max_avg.groupby(['Property Name'])['V3 max'].max().tolist()
    SF_ongoing_max_avg_max['M2 max'] = SF_ongoing_max_avg.groupby(['Property Name'])['M2 max'].max().tolist()
    SF_ongoing_max_avg_max['M3 max'] = SF_ongoing_max_avg.groupby(['Property Name'])['M3 max'].max().tolist()
    
    SF_ongoing_max_avg_max.reset_index(inplace=True, drop=True)

#%% 결과값 정리
    
    SF_ongoing_max_avg_max = pd.merge(transfer_element_name.rename('Property Name'),\
                                      SF_ongoing_max_avg_max, how='left')
        
    SF_ongoing_max_avg_max = SF_ongoing_max_avg_max.dropna()
    SF_ongoing_max_avg_max.reset_index(inplace=True, drop=True)
    
    SF_output = pd.merge(SF_ongoing_max_avg_max, transfer_element_info,\
                         how='left', left_on='Property Name', right_on='Name')
    
    # 기존 시트에 V, M 값 넣기
    # wb = openpyxl.load_workbook(input_path + '\\' + column_xlsx, keep_links=True)
    
    # for idx, row in SF_output.iterrows():
        
    #     wb['Results'].cell(column=2, row=idx+5, value=row[0]) # name
    #     wb['Results'].cell(column=3, row=idx+5, value=row[8]) # b
    #     wb['Results'].cell(column=4, row=idx+5, value=row[9]) # h
    #     wb['Results'].cell(column=5, row=idx+5, value=row[23]) # direction
    #     wb['Results'].cell(column=6, row=idx+5, value=row[10]) # cover thickness
    #     wb['Results'].cell(column=7, row=idx+5, value=row[11]) # concrete grade
        
    #     wb['Results'].cell(column=8, row=idx+5, value=row[13]) # main bar diameter
    #     wb['Results'].cell(column=9, row=idx+5, value=row[24]) # main bar strength
    #     wb['Results'].cell(column=10, row=idx+5, value=row[15]) # hoop bar diameter
    #     wb['Results'].cell(column=11, row=idx+5, value=row[25]) # hoop bar strength
        
    #     wb['Results'].cell(column=12, row=idx+5, value=row[16]) # layer 1 ea
    #     wb['Results'].cell(column=13, row=idx+5, value=row[17]) # layer 1 row
    #     wb['Results'].cell(column=14, row=idx+5, value=row[18]) # layer 2 ea
    #     wb['Results'].cell(column=15, row=idx+5, value=row[19]) # layer 2 row
        
    #     wb['Results'].cell(column=16, row=idx+5, value=row[20]) # hoop X
    #     wb['Results'].cell(column=17, row=idx+5, value=row[21]) # hoop Y
    #     wb['Results'].cell(column=18, row=idx+5, value=row[22]) # hoop spacing        
        
    #     wb['Results'].cell(column=19, row=idx+5, value=row[2]) # P
    #     wb['Results'].cell(column=20, row=idx+5, value=row[5]) # M2
    #     wb['Results'].cell(column=21, row=idx+5, value=row[6]) # M3
    #     wb['Results'].cell(column=22, row=idx+5, value=row[4]) # V3
    #     wb['Results'].cell(column=23, row=idx+5, value=row[3]) # V2
        
    # wb.save(input_path + '\\' + column_xlsx)
    # wb.close()
    
    # 기존 시트에 V, M 값 넣기(alt2)
    SF_output = SF_output.iloc[:,[0,8,9,23,10,11,13,24,15,25,16,17,\
                                  18,19,20,21,22,2,5,6,4,3]] # SF_output 재정렬
    
# nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)

    SF_output = SF_output.replace(np.nan, '', regex=True)

#%% 출력 (Using win32com...)
    
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application') # 엑셀 실행
    excel.Visible = False # 엑셀창 안보이게

    wb = excel.Workbooks.Open(input_path + '\\' + column_xlsx)
    ws = wb.Sheets('Results')
    
    startrow, startcol = 5, 2
    
    ws.Range(ws.Cells(startrow, startcol),\
              ws.Cells(startrow+SF_output.shape[0]-1,\
                      startcol+SF_output.shape[1]-1)).Value\
    = list(SF_output.itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능
    
    if export_to_pdf == True:
    
    # try:
        for i in range(SF_output.shape[0]):
                   
            wb.Worksheets(2).Select()
            
            wb.Worksheets(2).Name = '({})'.format(i+1)
            
            xlTypePDF = 0
            xlQualityStandard = 0
            
            wb.ActiveSheet.ExportAsFixedFormat(xlTypePDF\
                                               ,os.path.join(input_path, pdf_name+'({}).pdf'.format(i+1))\
                                               , xlQualityStandard, True, False)
                
    # except Exception as e:
    #     print(e)
        
    # finally:
    wb.Close(SaveChanges=1) # Closing the workbook
    excel.Quit() # Closing the application

#%% Done!

    print('Done!')
    
#%% Transfer Column SF (DCR) PDF export

def trans_column_SF_pdf(input_path, column_xlsx, pdf_name='Transfer Column Results'):
    ''' 

    Perform-3D 해석 결과에서 기둥의 축력, 전단력, 모멘트를 불러와 Results_E.Column 엑셀파일을 작성. \n
    result_path : Perform-3D에서 나온 해석 파일의 경로. \n
    result_xlsx : Perform-3D에서 나온 해석 파일의 이름. 해당 파일 이름이 포함된 파일들을 모두 불러온다. \n
    input_path : Data Conversion 엑셀 파일의 경로 \n
    input_xlsx : Data Conversion 엑셀 파일의 이름. result_xlsx와는 달리 확장자명(.xlsx)까지 기입해줘야한다. 하나의 파일만 불러온다. \n
    column_xlsx : Results_E.Column 엑셀 파일의 이름.확장자명(.xlsx)까지 기입해줘야한다. \n
    export_to_pdf : 입력된 값에 따른 각 부재들의 결과 시트를 pdf로 출력. True = pdf 출력, False = pdf 미출력(Results_E.Column 엑셀파일만 작성됨).
    pdf_name = 출력할 pdf 파일 이름.
    
    '''

#%% Column 엑셀시트 불러오기

    SF_output = pd.read_excel(input_path + '\\' + column_xlsx, sheet_name='Results', usecols=[0], skiprows=3)
    SF_output = SF_output.iloc[:,0]
    SF_output = SF_output[SF_output.str.contains('\(', na=False)]

#%% 출력 (Using win32com...)
    
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application') # 엑셀 실행
    excel.Visible = False # 엑셀창 안보이게

    wb = excel.Workbooks.Open(input_path + '\\' + column_xlsx)
    ws = wb.Sheets('Results')
    
    for i in range(SF_output.shape[0]):
               
        wb.Worksheets(2).Select()
        
        wb.Worksheets(2).Name = '({})'.format(i+1)
        
        xlTypePDF = 0
        xlQualityStandard = 0
        
        wb.ActiveSheet.ExportAsFixedFormat(xlTypePDF\
                                           ,os.path.join(input_path, pdf_name+'({}).pdf'.format(i+1))\
                                           , xlQualityStandard, True, False)
                
    wb.Close(SaveChanges=1) # Closing the workbook
    excel.Quit() # Closing the application

#%% Done!

    print('Done!')