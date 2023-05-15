import pandas as pd
import numpy as np
import os
from collections import deque # Double-ended Queue : 자료의 앞, 뒤 양 방향에서 자료를 추가하거나 제거가능
import win32com.client
import pythoncom

###############################################################################
###############################################################################
#%% User Input

# Data Conversion Sheet, Column Sheet, Beam Sheet
input_xlsx_path = r'K:\2105-이형우\성능기반 내진설계\KHSM\107\KHSM_107_Data Conversion_Ver.1.3M.xlsx'
# input_xlsx_path = r'K:\2105-이형우\KHSM_106_Data Conversion_Ver.1.3M.xlsx'
# input_xlsx_path = r'D:\이형우\성능기반 내진설계\21-GR-222 광명 4R구역 주택재개발사업 성능기반내진설계\101\101D_N3_Data Conversion_Ver.1.2.xlsx'
input_xlsx_sheet = 'Output_G.Column Properties (2)'
result_xlsx_path = r'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\Results\107\KHSM_107_1_Pre-Analysis.xlsx'

###############################################################################
###############################################################################

transfer_element_name = pd.DataFrame()
input_data_raw = pd.ExcelFile(input_xlsx_path)
input_data_sheets = pd.read_excel(input_data_raw, [input_xlsx_sheet], skiprows=3)

transfer_element_name = input_data_sheets[input_xlsx_sheet].iloc[:,0]
transfer_element_name.name = 'Property Name'

#%% Analysis Result 불러오기
to_load_list = result_xlsx_path

# 전단력 Data
SF_info_data = pd.read_excel(to_load_list, sheet_name='Frame Results - End Forces'
                                  , skiprows=[0, 2], header=0
                                  , usecols=[0,2,5,7,8]) # usecols로 원하는 열만 불러오기

SF_info_data = SF_info_data.sort_values(by=['Load Case', 'Element Name', 'Step Type']) # 지진파 순서가 섞여있을 때 sort

# 부재 이름 Matching을 위한 Element 정보
element_info_data = pd.read_excel(to_load_list, sheet_name='Element Data - Frame Types'
                                  , skiprows=[0, 2], header=0, usecols=[0, 2, 5]) # usecols로 원하는 열만 불러오기

# 필요한 부재만 선별
element_info_data = element_info_data[element_info_data['Property Name'].isin(transfer_element_name)]

# 전단력, 부재 이름 Matching (by Element Name)
SF_ongoing = pd.merge(element_info_data.iloc[:, [1,2]], SF_info_data.iloc[:, 1:], how='left')

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
SF_ongoing['P I-End'] = SF_ongoing['P I-End'].abs() * 1.0


# max, min 중 최대값 뽑기
SF_ongoing_max = SF_ongoing.loc[SF_ongoing.groupby(SF_ongoing.index // 2)['P I-End'].idxmax()]


# 필요한 하중만 포함된 데이터 slice (MCE)
SF_ongoing_max = SF_ongoing_max[SF_ongoing_max['Load Case']\
                                .str.contains('|'.join(MCE_load_name_list))] # function equivalent of a combination of df.isin() and df.str.contains()
 
# 같은 부재(그러나 잘려있는) 경우 최대값 뽑기
SF_ongoing_max_max = SF_ongoing_max.loc[SF_ongoing_max.groupby(['Property Name'])['P I-End'].idxmax()]

SF_ongoing_max_max.reset_index(inplace=True, drop=True)

#%% 결과값 정렬
SF_ongoing_max_max = pd.merge(transfer_element_name, SF_ongoing_max_max, how='left')

#%% 출력 (Using win32com...)

# Using win32com...
# Call CoInitialize function before using any COM object
excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
excel.Visible = True # 엑셀창 안보이게

wb = excel.Workbooks.Open(input_xlsx_path)
ws = wb.Sheets(input_xlsx_sheet)

startrow, startcol = 5, 17    
ws.Range(ws.Cells(startrow, startcol),\
          ws.Cells(startrow+SF_ongoing_max_max.shape[0]-1, startcol)).Value\
= [[i] for i in SF_ongoing_max_max['P I-End']]

#%% NG인 부재의 Hoop 간격 줄이기(-10mm every iteration)
# Check Vy <= Vn
while True:    
    
    # Vy, Vn 비교 결과 읽기
    vy_vn = ws.Range('DF%s:DG%s' %(startrow, startrow+SF_ongoing_max_max.shape[0]-1)).Value # list of tuples
    # Hoop 간격 읽기
    h_space = ws.Range('P%s:P%s' %(startrow, startrow+SF_ongoing_max_max.shape[0]-1)).Value # list of tuples
    h_space_array = np.array(h_space)[:,0] # list of tuples -> np.array    
    # vy_vn의 결과에 따른 np,array 생성   (NG가 있는 경우 = 1, NG가 없는 경우 = 0)
    vy_vn_array = np.array([1 if 'N.G' in row else 0 for row in vy_vn])
    
    # NG 부재가 없거나, Hoop 간격이 0이하가 되는 경우 break
    if (np.all(vy_vn_array == 0)) | (np.any(h_space_array <= 0)):
        break
    
    # NG인 부재의 Hoop 간격 줄이기
    h_space_array = np.where(vy_vn_array == 1, h_space_array-10, h_space_array)
    
    # Hoop 간격의 변경된 값을 Excel에 다시 입력
    ws.Range('P%s:P%s' %(startrow, startrow+SF_ongoing_max_max.shape[0]-1)).Value\
    = [[i] for i in h_space_array]    
    
    # Hoop 간격이 변경된(Vy, Vn 비교 결과, NG가 난) index의 list 만들기
    h_space_changed_idx = np.where(vy_vn_array == 1)[0]
    for j in h_space_changed_idx:
        ws.Range('P%s' %str(startrow+int(j))).Font.ColorIndex = 3

element_name = ws.Range('A%s:A%s' %(startrow, startrow+SF_ongoing_max_max.shape[0]-1)).Value
element_name = pd.DataFrame(element_name)
element_name_splitted = pd.DataFrame(element_name.iloc[:,0].str.split('_', expand=True))

element_info = pd.concat([element_name, pd.Series(h_space_array)], axis=1)
element_info['Name'] = element_name_splitted.iloc[:,0]
element_info['Story'] = element_name_splitted.iloc[:,2]

#%%
wb.Save()            
# wb.Close(SaveChanges=1) # Closing the workbook
# excel.Quit() # Closing the application



