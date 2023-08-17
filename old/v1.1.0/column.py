import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from collections import deque # Double-ended Queue : 자료의 앞, 뒤 양 방향에서 자료를 추가하거나 제거가능
import win32com.client
import pythoncom
from PyPDF2 import PdfMerger, PdfFileReader

#%% Elastic Column SF (DCR)
def E_CSF(input_xlsx_path, result_xlsx_path, column_xlsx_path\
          , export_to_pdf=True, pdf_name='E.Column Results'):
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
    input_data_raw = pd.ExcelFile(input_xlsx_path)
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
    to_load_list = result_xlsx_path

    # 전단력 Data
    SF_info_data = pd.DataFrame()
    for i in to_load_list:
        SF_info_data_temp = pd.read_excel(i, sheet_name='Frame Results - End Forces'
                                          , skiprows=[0, 2], header=0
                                          , usecols=[0,2,5,7,8,10,12,15,16,17,18]) # usecols로 원하는 열만 불러오기
        SF_info_data = pd.concat([SF_info_data, SF_info_data_temp])

    SF_info_data = SF_info_data.sort_values(by=['Load Case', 'Element Name', 'Step Type']) # 지진파 순서가 섞여있을 때 sort

    # 부재 이름 Matching을 위한 Element 정보
    element_info_data = pd.DataFrame()
    for i in to_load_list:
        element_info_data_temp = pd.read_excel(i, sheet_name='Element Data - Frame Types'
                                               , skiprows=[0, 2], header=0, usecols=[0, 2, 5, 7]) # usecols로 원하는 열만 불러오기
        element_info_data = pd.concat([element_info_data, element_info_data_temp])

    # 필요한 부재만 선별
    element_info_data = element_info_data[element_info_data['Property Name'].isin(transfer_element_name)]
    
    # 층 정보 Matching을 위한 Node 정보
    height_info_data = pd.DataFrame()    
    for i in to_load_list:
        height_info_data_temp = pd.read_excel(i, sheet_name='Node Coordinate Data'
                                              , skiprows=[0, 2], header=0, usecols=[1, 4]) # usecols로 원하는 열만 불러오기
        height_info_data = pd.concat([height_info_data, height_info_data_temp])

    element_info_data = pd.merge(element_info_data, height_info_data, how='left', left_on='I-Node ID', right_on='Node ID')

    element_info_data = element_info_data.drop_duplicates()

    # 전단력, 부재 이름 Matching (by Element Name)
    SF_ongoing = pd.merge(element_info_data.iloc[:, [1,2,5]], SF_info_data.iloc[:, 1:], how='left')

    SF_ongoing = SF_ongoing.sort_values(by=['Element Name', 'Load Case', 'Step Type'])

    SF_ongoing.reset_index(inplace=True, drop=True)

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
        SF_ongoing['Load Case'] = SF_ongoing['Load Case'].str.replace('[1] + %s'%i, '[1] + %s'%j, regex=False)
    ###########################################################################

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
    
    # 기존 시트에 V, M 값 넣기(alt2)
    SF_output = SF_output.iloc[:,[0,8,9,23,10,11,13,24,15,25,16,17,\
                                  18,19,20,21,22,2,5,6,4,3]] # SF_output 재정렬
    
# nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)

    SF_output = SF_output.replace(np.nan, '', regex=True)

#%% 출력 (Using win32com...)    
    # Using win32com...
    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게

    wb = excel.Workbooks.Open(os.path.dirname(result_xlsx_path[0]) + '\\' + column_xlsx_path)
    ws = wb.Sheets('Results')
    
    startrow, startcol = 5, 2
    
    ws.Range(ws.Cells(startrow, startcol),\
              ws.Cells(startrow+SF_output.shape[0]-1,\
                      startcol+SF_output.shape[1]-1)).Value\
    = list(SF_output.itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능
    
    if export_to_pdf == True:
        # pdf Merge를 위한 PdfMerger 클래스 생성
        merger = PdfMerger()

        result_path = os.path.dirname(result_xlsx_path[0])
        # print(os.getcwd() + '\\template\\Results_E.Column_Ver.1.3.xlsx')
        # print(result_path)

        for i in range(SF_output.shape[0]):

            pdf_file_path = os.path.join(result_path, pdf_name+'({}).pdf'.format(i+1))
            
            wb.Worksheets(2).Select()            
            wb.Worksheets(2).Name = '({})'.format(i+1)
            
            xlTypePDF = 0
            xlQualityStandard = 0
            
            wb.ActiveSheet.ExportAsFixedFormat(xlTypePDF, pdf_file_path\
                                               , xlQualityStandard, True, False)    
            merger.append(pdf_file_path)
            
        merger.write(result_path+'\\'+'{}.pdf'.format(pdf_name))
        merger.close()
        # Merge한 후 개별 파일들 지우기    
        for i in range(SF_output.shape[0]):
            pdf_file_path = os.path.join(result_path, pdf_name+'({}).pdf'.format(i+1))
            os.remove(pdf_file_path)

    wb.Save()
    # wb.SaveAs(Filename= column_xlsx_path)
    # wb.Close()
    # wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application
    
#%% Transfer Column SF (DCR) PDF export
def E_CSF_pdf(column_xlsx_path, pdf_name='E.Column Results'):
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
    SF_output = pd.read_excel(column_xlsx_path, sheet_name='Results', usecols=[0], skiprows=3)
    SF_output = SF_output.iloc[:,0]
    SF_output = SF_output[SF_output.str.contains('\(', na=False)]

#%% 출력 (Using win32com...)
    
    # Using win32com...
    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게

    wb = excel.Workbooks.Open(column_xlsx_path)
    ws = wb.Sheets('Results')
    
    # pdf Merge를 위한 PdfMerger 클래스 생성
    merger = PdfMerger()
    
    input_path = os.path.dirname(column_xlsx_path)

    for i in range(SF_output.shape[0]):
        
        pdf_file_path = os.path.join(input_path, pdf_name+'({}).pdf'.format(i+1))
               
        wb.Worksheets(2).Select()        
        wb.Worksheets(2).Name = '({})'.format(i+1)
        
        xlTypePDF = 0
        xlQualityStandard = 0
        
        wb.ActiveSheet.ExportAsFixedFormat(xlTypePDF, pdf_file_path
                                           , xlQualityStandard, True, False)
        
        merger.append(pdf_file_path)
    
    merger.write(input_path+'\\'+'{}.pdf'.format(pdf_name))
    merger.close()
    
    # Merge한 후 개별 파일들 지우기    
    for i in range(SF_output.shape[0]):
        pdf_file_path = os.path.join(input_path, pdf_name+'({}).pdf'.format(i+1))
        os.remove(pdf_file_path)
            
    wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application
    
#%% General Column SF (DCR)
def CSF(input_xlsx_path, result_xlsx_path, DCR_criteria=1, yticks=2, xlim=3):
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

    input_xlsx_sheet = 'Output_G.Column Properties'
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
    element_data = element_data[element_data['Property Name'].isin(element_name)]

#%% Analysis Result에 Element, Node 정보 매칭

    element_data = element_data.drop_duplicates()
    
    element_data = pd.merge(element_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
    SF_ongoing = pd.merge(element_data.iloc[:, [1,2,5]], SF_info_data.iloc[:, 1:], how='left')
    SF_ongoing.reset_index(inplace=True, drop=True)

#%% 지진파 이름 list 만들기

    load_name_list = []
    for i in SF_ongoing['Load Case'].drop_duplicates():
        new_i = i.split('+')[1]
        new_i = new_i.strip()
        load_name_list.append(new_i)

    gravity_load_name = [x for x in load_name_list if 'GL' in x]
    seismic_load_name_list = [x for x in load_name_list if'GL' not in x]

    seismic_load_name_list.sort()

    DE_load_name_list = [x for x in load_name_list if 'MCE' not in x]
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]
    
#%% V, M값에 절대값, 최대값, 평균값 뽑기

    # 절대값
    SF_ongoing.iloc[:,[5,6]] = SF_ongoing.iloc[:,[5,6]].abs()

    # V2, V3의 최대값을 저장하기 위해 필요한 데이터 slice
    SF_ongoing_max = SF_ongoing.iloc[[2*x for x in range(int(SF_ongoing.shape[0]/2))],[0,1,2,3]] 
    # [2*x for x in range(int(SF_ongoing.shape[0]/2] -> [짝수 index]
    
    # V2, V3의 최대값을 저장
    SF_ongoing_max['V2 max'] = SF_ongoing.groupby(SF_ongoing.index // 2)['V2 I-End'].max().tolist()
    SF_ongoing_max['V3 max'] = SF_ongoing.groupby(SF_ongoing.index // 2)['V3 I-End'].max().tolist()

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
    SF_ongoing_max_avg['V3 max(MCE)'] = SF_ongoing_max_MCE.groupby(['Element Name'])['V3 max'].mean()
    SF_ongoing_max_avg['V3 max(G)'] = SF_ongoing_max_G.groupby(['Element Name'])['V3 max'].mean()
 
    
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
    SF_ongoing_max_avg_max = pd.merge(SF_ongoing_max_avg_max
                                      , SF_ongoing_max_avg.groupby(['Property Name'])['V3 max(MCE)'].max()
                                      , left_on='Property Name', right_index=True, suffixes=('_before', '_after'))
    SF_ongoing_max_avg_max = pd.merge(SF_ongoing_max_avg_max
                                      , SF_ongoing_max_avg.groupby(['Property Name'])['V3 max(G)'].max()
                                      , left_on='Property Name', right_index=True, suffixes=('_before', '_after'))
    
    # MCE에 대해 1.2배, G에 대해 0.2배
    SF_ongoing_max_avg_max['V2 max(MCE)_after'] = SF_ongoing_max_avg_max['V2 max(MCE)_after'] * 1.2
    SF_ongoing_max_avg_max['V2 max(G)_after'] = SF_ongoing_max_avg_max['V2 max(G)_after'] * 0.2
    SF_ongoing_max_avg_max['V3 max(MCE)_after'] = SF_ongoing_max_avg_max['V3 max(MCE)_after'] * 1.2
    SF_ongoing_max_avg_max['V3 max(G)_after'] = SF_ongoing_max_avg_max['V3 max(G)_after'] * 0.2
    
    SF_ongoing_max_avg_max.reset_index(inplace=True, drop=False)

#%% 결과값 정리
    
    SF_output = pd.merge(element_name, SF_ongoing_max_avg_max, how='left')
        
    SF_output = SF_output.dropna()
    SF_output.reset_index(inplace=True, drop=True)
        
    # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
    SF_output = SF_output.replace(np.nan, '', regex=True)
    
    # 기존 시트에 V값 넣기
    SF_output1 = SF_output.iloc[:,0]
    SF_output2 = SF_output.iloc[:,[6,7,8,9]]

#%% 출력 (Using win32com...)
    
    # Using win32com...
    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게

    wb = excel.Workbooks.Open(input_xlsx_path)
    ws = wb.Sheets('Results_G.Column')
    
    startrow, startcol = 5, 1    
    ws.Range(ws.Cells(startrow, startcol),\
              ws.Cells(startrow+SF_output1.shape[0]-1, startcol)).Value\
    = [[i] for i in SF_output1]
    
    startrow, startcol = 5, 18    
    ws.Range(ws.Cells(startrow, startcol),\
              ws.Cells(startrow+SF_output2.shape[0]-1,\
                      startcol+SF_output2.shape[1]-1)).Value\
    = list(SF_output2.itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능
    
    wb.Save()            
    # wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application
    

#%% Column Shear Force 결과(DCR) 불러오기

    input_data_raw = pd.ExcelFile(input_xlsx_path)
    input_data_sheets = pd.read_excel(input_data_raw, ['Results_G.Column'], skiprows=3)
    input_data_raw.close()
    
    SF_DCR = input_data_sheets['Results_G.Column'].iloc[:,[0,31,33]]
    
    SF_DCR.columns = ['Property Name', 'DCR_MCE(X)', 'DCR_MCE(Y)']
    
    ########### 허무원
    # story_name_splitted = SF_DCR['Property Name'].str.split('_', expand=True)
    # SF_DCR['Story Name'] = story_name_splitted.iloc[:,2]
    # SF_output = pd.merge(SF_DCR, story_info, how='left')
    # SF_output.columns = ['Property Name', 'DCR_MCE(X)', 'DCR_MCE(Y)', 'Story Name', 'Index', 'V']
    ############

    # SF_output에 DCR값 merge(그래프 그릴 때 height(mm) 정보가 필요함)
    SF_output = pd.merge(SF_output, SF_DCR, how='left') 

#%% 그래프
    count = 1
    
    # ### DE 그래프
    # if len(DE_load_name_list) != 0:
        
    #     fig1 = plt.figure(count, dpi=150, figsize=(5,6))
    #     plt.xlim(0, xlim)
        
    #     plt.scatter(SWR_avg_total['DCR_DE_min'], SWR_avg_total['Height'], color='k', s=1)
    #     plt.scatter(SWR_avg_total['DCR_DE_max'], SWR_avg_total['Height'], color='k', s=1)
    #     plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
    #     plt.axvline(x = DCR_criteria, color='r', linestyle='--')
        
    #     # 기타
    #     plt.grid(linestyle='-.')
    #     plt.xlabel('D/C Ratios')
    #     plt.ylabel('Story')
    #     plt.title('Wall Rotation (DE)')
        
    #     plt.close()
    #     count += 1
                            
    #     yield fig1
    #     yield 'DE' # Marker 출력
        
    ### MCE 그래프
    if len(MCE_load_name_list) != 0:
        
        ### H1 MCE 그래프 ###
        fig3 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        plt.scatter(SF_output['DCR_MCE(X)'], SF_output['V'], color='k', s=1)
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        plt.axvline(x = DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Shear Strength (X MCE)')
        
        plt.close()
        count += 1
        
        yield fig3
        
        ### H2 MCE 그래프 ###
        fig4 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        plt.scatter(SF_output['DCR_MCE(Y)'], SF_output['V'], color='k', s=1)
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        plt.axvline(x = DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Shear Strength (Y MCE)')
        
        plt.close()
        count += 1
        
        yield fig4
        yield 'MCE' # Marker 출력
        
#%% Column Rotation (DCR)
def CR_DCR(input_xlsx_path, result_xlsx_path
           , col_group='G.Column', DCR_criteria=1, yticks=2, xlim=3):

#%% Input Sheets 정보 load
    story_info = pd.DataFrame()
    deformation_cap = pd.DataFrame()
    
    input_data_raw = pd.ExcelFile(input_xlsx_path)
    input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', 'Output_G.Column Properties'], skiprows=3)
    input_data_raw.close()
    
    story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
    deformation_cap = input_data_sheets['Output_G.Column Properties'].iloc[:,[0,80,81,82,83]]
    
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    deformation_cap.columns = ['Name', 'LS(X)', 'LS(Y)', 'CP(X)', 'CP(Y)']
    
#%% Analysis Result 불러오기   
    to_load_list = result_xlsx_path
    
    beam_rot_data = pd.DataFrame()
    
    for i in to_load_list:
        result_data_raw = pd.ExcelFile(i)
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
    beam_rot_data = beam_rot_data[beam_rot_data['Group Name'] == col_group]
    beam_rot_data = beam_rot_data[beam_rot_data['Distance from I-End'] == 0]
    
#%% Analysis Result에 Element, Node 정보 매칭
    
    element_data = element_data.drop_duplicates()
    node_data = node_data.drop_duplicates()
    
    beam_rot_data = pd.merge(beam_rot_data, element_data, how='left')
    beam_rot_data = pd.merge(beam_rot_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
    
    beam_rot_data = beam_rot_data[beam_rot_data['Property Name'].notna()]
     
    # 필요한 정보들만 다시 모아서 new dataframe
    beam_rot_data = beam_rot_data.iloc[:, [0,1,7,10,2,3,5,6]]
    
    beam_rot_data.reset_index(inplace=True, drop=True)
    
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
    
    beam_rot_data['DE_X Rotation(rad)'] = beam_rot_data['H2 Rotation(rad)'].abs() / beam_rot_data['LS(X)']
    beam_rot_data['DE_Y Rotation(rad)'] = beam_rot_data['H3 Rotation(rad)'].abs() / beam_rot_data['LS(Y)']
    beam_rot_data['MCE_X Rotation(rad)'] = beam_rot_data['H2 Rotation(rad)'].abs() / beam_rot_data['CP(X)']
    beam_rot_data['MCE_Y Rotation(rad)'] = beam_rot_data['H3 Rotation(rad)'].abs() / beam_rot_data['CP(Y)']
    
    beam_rot_data = beam_rot_data[beam_rot_data['Name'].notna()]
    
#%% 조작용 코드
    # 없애고 싶은 부재의 이름 입력(error_beam 확인 후!, DE, MCE에서 다 없어짐)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('AC405_1_17F'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('B15_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB4A_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB4B_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB3D_'))].index)

#%% DE 결과 Plot
    count = 1
    if len(DE_load_name_list) != 0:
        
        beam_rot_data_total_DE = pd.DataFrame()    
        
        for load_name in DE_load_name_list:
        
            temp_df_X_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['DE_X Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_DE['{}_X_max'.format(load_name)] = temp_df_X_max.tolist()
            
            temp_df_X_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['DE_X Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_DE['{}_X_min'.format(load_name)] = temp_df_X_min.tolist()
            
            temp_df_Y_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['DE_Y Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_DE['{}_Y_max'.format(load_name)] = temp_df_Y_max.tolist()
            
            temp_df_Y_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['DE_Y Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_DE['{}_Y_min'.format(load_name)] = temp_df_Y_min.tolist()
            
        beam_rot_data_total_DE['Element Name'] = temp_df_X_max.index
        
        beam_rot_data_total_DE.reset_index(inplace=True, drop=True)
        
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, element_data, how='left')
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, story_info, how='left', left_on='V(mm)', right_on='Height(mm)')
        beam_rot_data_total_DE.sort_values('Height(mm)', inplace=True)
        # beam_rot_data_total_DE.reset_index(inplace=True, drop=True)
        
        # 평균 열 생성        
        beam_rot_data_total_DE['DE_X Max avg'] = beam_rot_data_total_DE.iloc[:,list(range(0,len(DE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_DE['DE_X Min avg'] = beam_rot_data_total_DE.iloc[:,list(range(1,len(DE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_DE['DE_Y Max avg'] = beam_rot_data_total_DE.iloc[:,list(range(2,len(DE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_DE['DE_Y Min avg'] = beam_rot_data_total_DE.iloc[:,list(range(3,len(DE_load_name_list)*4,4))].mean(axis=1)
    
        # 전체 Plot            
        ### DE X ###
        fig1 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        plt.scatter(beam_rot_data_total_DE['DE_X Max avg'], beam_rot_data_total_DE.loc[:,'V(mm)'], color='k', s=1)
        plt.scatter(beam_rot_data_total_DE['DE_X Min avg'], beam_rot_data_total_DE.loc[:,'V(mm)'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Column Rotation (X DE)')
        
        plt.tight_layout()   
        plt.close()

        # 기준 넘는 점 확인
        error_beam_DE_X = beam_rot_data_total_DE[['Element Name', 'Property Name', 'Story Name', 'DE_X Max avg', 'DE_X Min avg']]\
                      [(beam_rot_data_total_DE['DE_X Max avg'] >= DCR_criteria) | (beam_rot_data_total_DE['DE_X Min avg'] >= DCR_criteria)]
        
        count += 1
        
        yield fig1
        yield error_beam_DE_X
        
        ### DE Y ###
        fig2 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        plt.scatter(beam_rot_data_total_DE['DE_Y Max avg'], beam_rot_data_total_DE.loc[:,'V(mm)'], color='k', s=1)
        plt.scatter(beam_rot_data_total_DE['DE_Y Min avg'], beam_rot_data_total_DE.loc[:,'V(mm)'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Column Rotation (Y DE)')
        
        plt.tight_layout()   
        plt.close()

        # 기준 넘는 점 확인
        error_beam_DE_Y = beam_rot_data_total_DE[['Element Name', 'Property Name', 'Story Name', 'DE_Y Max avg', 'DE_Y Min avg']]\
                      [(beam_rot_data_total_DE['DE_Y Max avg'] >= DCR_criteria) | (beam_rot_data_total_DE['DE_Y Min avg'] >= DCR_criteria)]
        
        count += 1
        
        yield fig2
        yield error_beam_DE_Y
        yield 'DE' # Marker 출력
        
#%% MCE 결과 Plot
    
    if len(MCE_load_name_list) != 0:
        
        beam_rot_data_total_MCE = pd.DataFrame()    
        
        for load_name in MCE_load_name_list:
        
            temp_df_X_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['MCE_X Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_MCE['{}_X_max'.format(load_name)] = temp_df_X_max.tolist()
            
            temp_df_X_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['MCE_X Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_MCE['{}_X_min'.format(load_name)] = temp_df_X_min.tolist()
            
            temp_df_Y_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['MCE_Y Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_MCE['{}_Y_max'.format(load_name)] = temp_df_Y_max.tolist()
            
            temp_df_Y_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['MCE_Y Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_MCE['{}_Y_min'.format(load_name)] = temp_df_Y_min.tolist()
            
        beam_rot_data_total_MCE['Element Name'] = temp_df_X_max.index
        
        beam_rot_data_total_MCE.reset_index(inplace=True, drop=True)
        
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, element_data, how='left')
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, story_info, how='left', left_on='V(mm)', right_on='Height(mm)')
        beam_rot_data_total_MCE.sort_values('Height(mm)', inplace=True)
        # beam_rot_data_total_MCE.reset_index(inplace=True, drop=True)
        
        # 평균 열 생성        
        beam_rot_data_total_MCE['MCE_X Max avg'] = beam_rot_data_total_MCE.iloc[:,list(range(0,len(MCE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_MCE['MCE_X Min avg'] = beam_rot_data_total_MCE.iloc[:,list(range(1,len(MCE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_MCE['MCE_Y Max avg'] = beam_rot_data_total_MCE.iloc[:,list(range(2,len(MCE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_MCE['MCE_Y Min avg'] = beam_rot_data_total_MCE.iloc[:,list(range(3,len(MCE_load_name_list)*4,4))].mean(axis=1)     

        # 전체 Plot 
        ### MCE X ###
        fig3 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        # 평균 plot
        plt.scatter(beam_rot_data_total_MCE['MCE_X Max avg'], beam_rot_data_total_MCE.loc[:,'V(mm)'], color='k', s=1)
        plt.scatter(beam_rot_data_total_MCE['MCE_X Min avg'], beam_rot_data_total_MCE.loc[:,'V(mm)'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Column Rotation (X MCE)')
        
        plt.tight_layout()
        plt.close()
    
        # 기준 넘는 점 확인    
        error_beam_MCE_X = beam_rot_data_total_MCE[['Element Name', 'Property Name', 'Story Name', 'MCE_X Max avg', 'MCE_X Min avg']]\
                      [(beam_rot_data_total_MCE['MCE_X Max avg'] >= DCR_criteria) | (beam_rot_data_total_MCE['MCE_X Min avg'] >= DCR_criteria)]
        
        count += 1
        
        yield fig3
        yield error_beam_MCE_X
        
        ### MCE X ###
        fig4 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        # 평균 plot
        plt.scatter(beam_rot_data_total_MCE['MCE_Y Max avg'], beam_rot_data_total_MCE.loc[:,'V(mm)'], color='k', s=1)
        plt.scatter(beam_rot_data_total_MCE['MCE_Y Min avg'], beam_rot_data_total_MCE.loc[:,'V(mm)'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Column Rotation (Y MCE)')
        
        plt.tight_layout()
        plt.close()
    
        # 기준 넘는 점 확인    
        error_beam_MCE_Y = beam_rot_data_total_MCE[['Element Name', 'Property Name', 'Story Name', 'MCE_Y Max avg', 'MCE_Y Min avg']]\
                      [(beam_rot_data_total_MCE['MCE_Y Max avg'] >= DCR_criteria) | (beam_rot_data_total_MCE['MCE_Y Min avg'] >= DCR_criteria)]
        
        count += 1
        
        yield fig4
        yield error_beam_MCE_Y
        yield 'MCE' # Marker 출력
        
#%% G.Column SF (elementwise)

def CSF_each(input_xlsx_path, retrofit_sheet=None): 
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
    input_xlsx_sheet = 'Results_G.Column'
    input_data_raw = pd.ExcelFile(input_xlsx_path)
    input_data_sheets = pd.read_excel(input_data_raw, [input_xlsx_sheet, retrofit_sheet], skiprows=3
                                 , usecols=[0,8,15,31,33])
    input_data_raw.close()
    
    col_before = input_data_sheets[input_xlsx_sheet]
    col_after = input_data_sheets[retrofit_sheet]

    col_before.columns = ['Name', 'Rebar Type(before)', 'Rebar Spacing(before)', 'MCE(H1)', 'MCE(H2)']
    col_after.columns = ['Name', 'Rebar Type(after)', 'Rebar Spacing(after)', 'MCE(H1)', 'MCE(H2)']
    
#%% 보강 전,후 Column dataframe 정리

    # 4개의 DCR 열에서 max값을 추출한 열 만들기
    col_before['DCR_max(before)'] = col_before[['MCE(H1)', 'MCE(H2)']].max(axis=1)
    col_after['DCR_max(after)'] = col_after[['MCE(H1)', 'MCE(H2)']].max(axis=1)

    # DCR 열 반올림하기(소수점 5자리까지)
    col_before['DCR_max(before)'] = col_before['DCR_max(before)'].round(5)
    col_after['DCR_max(after)'] = col_after['DCR_max(after)'].round(5)

    # 필요한 열만 추출
    col_output = pd.merge(col_before[['Name', 'Rebar Type(before)', 'Rebar Spacing(before)', 'DCR_max(before)']]
                           , col_after[['Name', 'Rebar Type(after)', 'Rebar Spacing(after)', 'DCR_max(after)']], how='left')

    # 이름 분리(부재 이름, 번호, 층)
    col_output['Property Name'] = col_output['Name'].str.split('_', expand=True)[0]
    col_output['Number'] = col_output['Name'].str.split('_', expand=True)[1]
    col_output['Story'] = col_output['Name'].str.split('_', expand=True)[2]

    # 부재 이름과 번호(C1_1)이 같은 부재들끼리 groupby로 묶고, list of dataframes 생성
    col_output_list = list(col_output.groupby(['Property Name', 'Number'], sort=False))
    
    yield col_output_list

#%% General Column SF - 허무원 박사

def CSF_HMW(input_xlsx_path, result_xlsx_path, DCR_criteria=1, yticks=2, xlim=3):
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

    input_xlsx_sheet = 'Output_G.Column Properties'
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
        
        SF_info_data_temp = result_data_sheets['Frame Results - End Forces'].iloc[:,[0,2,5,7,8,10,12]]
        SF_info_data = pd.concat([SF_info_data, SF_info_data_temp])

    node_data = result_data_sheets['Node Coordinate Data'].iloc[:,[1,4]]
    element_data = result_data_sheets['Element Data - Frame Types'].iloc[:,[0,2,5,7]] # beam의 양 nodes중 한 node에서의 rotation * 2

    # 필요한 부재만 선별
    element_data = element_data[element_data['Group Name'] == 'COLUMN']
    element_data['Property Name'] = element_data['Property Name'].str[:-1] + '_1_'

#%% 101동 element 이름 재명명(101동 부재 섞어서 쓰심)     ########## 허무원 ##########
    node_data_101 = result_data_sheets['Node Coordinate Data'].iloc[:,[1,2,3,4]]
    element_data_101 = pd.merge(element_data, node_data_101, how='left', left_on='I-Node ID', right_on='Node ID')
    
    list_101 = []    
    
    for idx, row in element_data_101.iterrows():
        if (row['Property Name'] == 'AC404_1_') & (row['H1'] == 3521.5):
            list_101.append('AC403_1_')
        else: 
            list_101.append(row['Property Name'])
            
    element_data['Property Name'] = list_101
    
#%% Analysis Result에 Element, Node 정보 매칭

    element_data = element_data.drop_duplicates()    
    element_data = pd.merge(element_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
    
#%%
    SF_ongoing = pd.merge(element_data.iloc[:, [1,2,5]], SF_info_data.iloc[:, 1:], how='left')
    SF_ongoing.reset_index(inplace=True, drop=True)
    
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

    gravity_load_name = [x for x in load_name_list if 'gl' in x]
    seismic_load_name_list = [x for x in load_name_list if '1.0D' not in x]

    seismic_load_name_list.sort()

    DE_load_name_list = [x for x in load_name_list if ('gl' not in x) & ('MCE' not in x)]
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]
    
#%% V, M값에 절대값, 최대값, 평균값 뽑기

    # 절대값
    SF_ongoing.iloc[:,[5,6,7]] = SF_ongoing.iloc[:,[5,6,7]].abs()

    # V2, V3의 최대값을 저장하기 위해 필요한 데이터 slice
    SF_ongoing_max = SF_ongoing.iloc[[2*x for x in range(int(SF_ongoing.shape[0]/2))],[0,1,2,3]] 
    # [2*x for x in range(int(SF_ongoing.shape[0]/2] -> [짝수 index]
    
    # V2, V3의 최대값을 저장
    SF_ongoing_max['V2 max'] = SF_ongoing.groupby(SF_ongoing.index // 2)['V2 I-End'].max().tolist()
    SF_ongoing_max['V3 max'] = SF_ongoing.groupby(SF_ongoing.index // 2)['V3 I-End'].max().tolist()
    SF_ongoing_max['P I-End'] = SF_ongoing.groupby(SF_ongoing.index // 2)['P I-End'].max().tolist()

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
    SF_ongoing_max_avg['V3 max(MCE)'] = SF_ongoing_max_MCE.groupby(['Element Name'])['V3 max'].mean()
    SF_ongoing_max_avg['V3 max(G)'] = SF_ongoing_max_G.groupby(['Element Name'])['V3 max'].mean()
    SF_ongoing_max_avg['P(G)'] = SF_ongoing_max_G.groupby(['Element Name'])['P I-End'].mean()
 
    
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
    SF_ongoing_max_avg_max = pd.merge(SF_ongoing_max_avg_max
                                      , SF_ongoing_max_avg.groupby(['Property Name'])['V3 max(MCE)'].max()
                                      , left_on='Property Name', right_index=True, suffixes=('_before', '_after'))
    SF_ongoing_max_avg_max = pd.merge(SF_ongoing_max_avg_max
                                      , SF_ongoing_max_avg.groupby(['Property Name'])['V3 max(G)'].max()
                                      , left_on='Property Name', right_index=True, suffixes=('_before', '_after'))
    SF_ongoing_max_avg_max = pd.merge(SF_ongoing_max_avg_max
                                      , SF_ongoing_max_avg.groupby(['Property Name'])['P(G)'].max()
                                      , left_on='Property Name', right_index=True, suffixes=('_before', '_after'))
    
    # MCE에 대해 1.2배, G에 대해 0.2배
    SF_ongoing_max_avg_max['V2 max(MCE)_after'] = SF_ongoing_max_avg_max['V2 max(MCE)_after'] * 1.2
    SF_ongoing_max_avg_max['V2 max(G)_after'] = SF_ongoing_max_avg_max['V2 max(G)_after'] * 0.2
    SF_ongoing_max_avg_max['V3 max(MCE)_after'] = SF_ongoing_max_avg_max['V3 max(MCE)_after'] * 1.2
    SF_ongoing_max_avg_max['V3 max(G)_after'] = SF_ongoing_max_avg_max['V3 max(G)_after'] * 0.2
    SF_ongoing_max_avg_max['P(G)_after'] = SF_ongoing_max_avg_max['P(G)_after'] * 1.0
    
    SF_ongoing_max_avg_max.reset_index(inplace=True, drop=False)

#%% 결과값 정리
    
    SF_output = pd.merge(element_name, SF_ongoing_max_avg_max, how='left')
        
    SF_output = SF_output.dropna()
    SF_output.reset_index(inplace=True, drop=True)
        
    # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
    SF_output = SF_output.replace(np.nan, '', regex=True)
    
    # 기존 시트에 V값 넣기
    SF_output1 = SF_output.iloc[:,0]
    SF_output2 = SF_output.iloc[:,[7,8,9,10]]
    SF_output3 = SF_output.iloc[:,11]

#%% 출력 (Using win32com...)
    
    # Using win32com...
    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게

    wb = excel.Workbooks.Open(input_xlsx_path)
    ws = wb.Sheets('Results_G.Column')
    ws2 = wb.Sheets('Output_G.Column Properties')
    
    startrow, startcol = 5, 1    
    ws.Range(ws.Cells(startrow, startcol),\
              ws.Cells(startrow+SF_output1.shape[0]-1, startcol)).Value\
    = [[i] for i in SF_output1]
    
    startrow, startcol = 5, 18    
    ws.Range(ws.Cells(startrow, startcol),\
              ws.Cells(startrow+SF_output2.shape[0]-1,\
                      startcol+SF_output2.shape[1]-1)).Value\
    = list(SF_output2.itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능
    
    ws2.Range('Q%s:Q%s' %(startrow, startrow+SF_output3.shape[0]-1)).Value\
        = [[i] for i in SF_output3]
    
    wb.Save()            
    # wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application

#%% Column Rotation (DCR) - 허무원 박사

def CR_HMW(input_xlsx_path, result_xlsx_path
           , col_group='G.Column', DCR_criteria=1, yticks=2, xlim=3):
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
#%% Input Sheets 정보 load
    story_info = pd.DataFrame()
    deformation_cap = pd.DataFrame()
    
    input_data_raw = pd.ExcelFile(input_xlsx_path)
    input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', 'Output_G.Column Properties'], skiprows=3)
    input_data_raw.close()
    
    story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
    deformation_cap = input_data_sheets['Output_G.Column Properties'].iloc[:,[0,80,81,82,83]]
    
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    deformation_cap.columns = ['Name', 'LS(X)', 'LS(Y)', 'CP(X)', 'CP(Y)']

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
    
    element_data = result_data_sheets['Element Data - Frame Types'].iloc[:,[0,2,5,7]] # beam의 양 nodes중 한 node에서의 rotation * 2

    beam_rot_data.columns = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'Distance from I-End', 'H2 Rotation(rad)', 'H3 Rotation(rad)']
    node_data.columns = ['Node ID', 'V(mm)']
    element_data.columns = ['Group Name', 'Element Name', 'Property Name', 'I-Node ID']
    
    # 필요한 부재만 선별
    element_data = element_data[element_data['Group Name'] == 'COLUMN']
    element_data['Property Name'] = element_data['Property Name'].str[:-1] + '_1_'

#%% 101동 element 이름 재명명(101동 부재 섞어서 쓰심)     ########## 허무원 ##########
    # node_data_101 = result_data_sheets['Node Coordinate Data'].iloc[:,[1,2,3,4]]
    # element_data_101 = pd.merge(element_data, node_data_101, how='left', left_on='I-Node ID', right_on='Node ID')
    
    # list_101 = []    
    
    # for idx, row in element_data_101.iterrows():
    #     if (row['Property Name'] == 'AC404_1_') & (row['H1'] == 3521.5):
    #         list_101.append('AC403_1_')
    #     else: 
    #         list_101.append(row['Property Name'])
            
    # element_data['Property Name'] = list_101
    
#%% Analysis Result에 Element, Node 정보 매칭

    element_data = element_data.drop_duplicates()    
    element_data = pd.merge(element_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
    
    beam_rot_data = beam_rot_data[beam_rot_data['Group Name'] == 'COLUMN']
    beam_rot_data = beam_rot_data[beam_rot_data['Distance from I-End'] == 0]
    
    beam_rot_data = pd.merge(beam_rot_data, element_data, how='left')
    beam_rot_data = beam_rot_data[beam_rot_data['Property Name'].notna()]
         
    beam_rot_data.reset_index(inplace=True, drop=True)
    
    # 이름에 층정보 붙이기
    beam_rot_data_copy = pd.merge(beam_rot_data, story_info, how='left', left_on = 'V(mm)', right_on = 'Height(mm)')
    new_name = beam_rot_data_copy['Property Name'] + beam_rot_data_copy['Story Name']
    beam_rot_data['Property Name'] = new_name
    
#%% Story info update (story z좌표 알아내기, 개별실행 후 엑셀에 붙여넣기)    ########## 허무원 ##########
    # story_updated = SF_ongoing['V'].drop_duplicates().sort_values(ascending=False)
    # story_updated.reset_index(inplace=True, drop=True)

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
    
#%% 성능기준(LS, CP) 정리해서 merge
    
    beam_rot_data = pd.merge(beam_rot_data, deformation_cap, how='left', left_on='Property Name', right_on='Name')
    
    beam_rot_data['DE_X Rotation(rad)'] = beam_rot_data['H2 Rotation(rad)'].abs() / beam_rot_data['LS(X)']
    beam_rot_data['DE_Y Rotation(rad)'] = beam_rot_data['H3 Rotation(rad)'].abs() / beam_rot_data['LS(Y)']
    beam_rot_data['MCE_X Rotation(rad)'] = beam_rot_data['H2 Rotation(rad)'].abs() / beam_rot_data['CP(X)']
    beam_rot_data['MCE_Y Rotation(rad)'] = beam_rot_data['H3 Rotation(rad)'].abs() / beam_rot_data['CP(Y)']
    
    beam_rot_data = beam_rot_data[beam_rot_data['Name'].notna()]
    
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
#%% 조작용 코드
    # 없애고 싶은 부재의 이름 입력(error_beam 확인 후!, DE, MCE에서 다 없어짐)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('AC405_1_17F'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('B15_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB4A_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB4B_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB3D_'))].index)

#%% DE 결과 Plot
    count = 1
    if len(DE_load_name_list) != 0:
        
        beam_rot_data_total_DE = pd.DataFrame()    
        
        for load_name in DE_load_name_list:
        
            temp_df_X_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['DE_X Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_DE['{}_X_max'.format(load_name)] = temp_df_X_max.tolist()
            
            temp_df_X_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['DE_X Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_DE['{}_X_min'.format(load_name)] = temp_df_X_min.tolist()
            
            temp_df_Y_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['DE_Y Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_DE['{}_Y_max'.format(load_name)] = temp_df_Y_max.tolist()
            
            temp_df_Y_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['DE_Y Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_DE['{}_Y_min'.format(load_name)] = temp_df_Y_min.tolist()
            
        beam_rot_data_total_DE['Element Name'] = temp_df_X_max.index
        
        beam_rot_data_total_DE.reset_index(inplace=True, drop=True)
        
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, element_data, how='left')
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, story_info, how='left', left_on='V(mm)_x', right_on='Height(mm)')
        beam_rot_data_total_DE.sort_values('Height(mm)', inplace=True)
        # beam_rot_data_total_DE.reset_index(inplace=True, drop=True)
        
        # 평균 열 생성        
        beam_rot_data_total_DE['DE_X Max avg'] = beam_rot_data_total_DE.iloc[:,list(range(0,len(DE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_DE['DE_X Min avg'] = beam_rot_data_total_DE.iloc[:,list(range(1,len(DE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_DE['DE_Y Max avg'] = beam_rot_data_total_DE.iloc[:,list(range(2,len(DE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_DE['DE_Y Min avg'] = beam_rot_data_total_DE.iloc[:,list(range(3,len(DE_load_name_list)*4,4))].mean(axis=1)
    
        # 전체 Plot            
        ### DE X ###
        fig1 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, 3)
        
        plt.scatter(beam_rot_data_total_DE['DE_X Max avg'], beam_rot_data_total_DE.loc[:,'V(mm)_x'], color='k', s=1)
        plt.scatter(beam_rot_data_total_DE['DE_X Min avg'], beam_rot_data_total_DE.loc[:,'V(mm)_x'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Column Rotation (X DE)')
        
        plt.tight_layout()   
        plt.close()

        # 기준 넘는 점 확인
        error_beam_DE_X = beam_rot_data_total_DE[['Element Name', 'Property Name', 'Story Name', 'DE_X Max avg', 'DE_X Min avg']]\
                      [(beam_rot_data_total_DE['DE_X Max avg'] >= DCR_criteria) | (beam_rot_data_total_DE['DE_X Min avg'] >= DCR_criteria)]
        
        count += 1
        
        yield fig1
        yield error_beam_DE_X
        
        ### DE Y ###
        fig2 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        plt.scatter(beam_rot_data_total_DE['DE_Y Max avg'], beam_rot_data_total_DE.loc[:,'V(mm)_x'], color='k', s=1)
        plt.scatter(beam_rot_data_total_DE['DE_Y Min avg'], beam_rot_data_total_DE.loc[:,'V(mm)_x'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Column Rotation (Y DE)')
        
        plt.tight_layout()   
        plt.close()

        # 기준 넘는 점 확인
        error_beam_DE_Y = beam_rot_data_total_DE[['Element Name', 'Property Name', 'Story Name', 'DE_Y Max avg', 'DE_Y Min avg']]\
                      [(beam_rot_data_total_DE['DE_Y Max avg'] >= DCR_criteria) | (beam_rot_data_total_DE['DE_Y Min avg'] >= DCR_criteria)]
        
        count += 1
        
        yield fig2
        yield error_beam_DE_Y
        yield 'DE' # Marker 출력
        
#%% MCE 결과 Plot
    
    if len(MCE_load_name_list) != 0:
        
        beam_rot_data_total_MCE = pd.DataFrame()    
        
        for load_name in MCE_load_name_list:
        
            temp_df_X_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['MCE_X Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_MCE['{}_X_max'.format(load_name)] = temp_df_X_max.tolist()
            
            temp_df_X_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['MCE_X Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_MCE['{}_X_min'.format(load_name)] = temp_df_X_min.tolist()
            
            temp_df_Y_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['MCE_Y Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_MCE['{}_Y_max'.format(load_name)] = temp_df_Y_max.tolist()
            
            temp_df_Y_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['MCE_Y Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_MCE['{}_Y_min'.format(load_name)] = temp_df_Y_min.tolist()
            
        beam_rot_data_total_MCE['Element Name'] = temp_df_X_max.index
        
        beam_rot_data_total_MCE.reset_index(inplace=True, drop=True)
        
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, element_data, how='left')
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, story_info, how='left', left_on='V(mm)_x', right_on='Height(mm)')
        beam_rot_data_total_MCE.sort_values('Height(mm)', inplace=True)
        # beam_rot_data_total_MCE.reset_index(inplace=True, drop=True)
        
        # 평균 열 생성        
        beam_rot_data_total_MCE['MCE_X Max avg'] = beam_rot_data_total_MCE.iloc[:,list(range(0,len(MCE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_MCE['MCE_X Min avg'] = beam_rot_data_total_MCE.iloc[:,list(range(1,len(MCE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_MCE['MCE_Y Max avg'] = beam_rot_data_total_MCE.iloc[:,list(range(2,len(MCE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_MCE['MCE_Y Min avg'] = beam_rot_data_total_MCE.iloc[:,list(range(3,len(MCE_load_name_list)*4,4))].mean(axis=1)     

        # 전체 Plot 
        ### MCE X ###
        fig3 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        # 평균 plot
        plt.scatter(beam_rot_data_total_MCE['MCE_X Max avg'], beam_rot_data_total_MCE.loc[:,'V(mm)_x'], color='k', s=1)
        plt.scatter(beam_rot_data_total_MCE['MCE_X Min avg'], beam_rot_data_total_MCE.loc[:,'V(mm)_x'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Column Rotation (X MCE)')
        
        plt.tight_layout()
        plt.close()
    
        # 기준 넘는 점 확인    
        error_beam_MCE_X = beam_rot_data_total_MCE[['Element Name', 'Property Name', 'Story Name', 'MCE_X Max avg', 'MCE_X Min avg']]\
                      [(beam_rot_data_total_MCE['MCE_X Max avg'] >= DCR_criteria) | (beam_rot_data_total_MCE['MCE_X Min avg'] >= DCR_criteria)]
        
        count += 1
        
        yield fig3
        yield error_beam_MCE_X
        
        ### MCE X ###
        fig4 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        # 평균 plot
        plt.scatter(beam_rot_data_total_MCE['MCE_Y Max avg'], beam_rot_data_total_MCE.loc[:,'V(mm)_x'], color='k', s=1)
        plt.scatter(beam_rot_data_total_MCE['MCE_Y Min avg'], beam_rot_data_total_MCE.loc[:,'V(mm)_x'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Column Rotation (Y MCE)')
        
        plt.tight_layout()
        plt.close()
    
        # 기준 넘는 점 확인    
        error_beam_MCE_Y = beam_rot_data_total_MCE[['Element Name', 'Property Name', 'Story Name', 'MCE_Y Max avg', 'MCE_Y Min avg']]\
                      [(beam_rot_data_total_MCE['MCE_Y Max avg'] >= DCR_criteria) | (beam_rot_data_total_MCE['MCE_Y Min avg'] >= DCR_criteria)]
        
        count += 1
        
        yield fig4
        yield error_beam_MCE_Y
        yield 'MCE' # Marker 출력