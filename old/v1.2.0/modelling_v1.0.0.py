'''
###########################    주의할 점     ##################################
* DL, LL : 해당 이름을 포함한 모든 Loads를 Import함
* drift_position : 예시. (2,5,7,11), (2,5,7,11,total), (NE,NW,SE,SW)
* 반드시 쉼표(,)로 구분할 것\n* 띄어쓰기, 밑줄(_) 사용 금지
* Output_Naming 시트의 결과를 모두 지우고 실행할 것
* Data Conversion 시트를 반드시 저장하고 실행할 것
###############################################################################
'''

import time
import PBD_p3d as pbd

#%% 시간 측정(START)
time_start = time.time()

###############################################################################
###############################################################################
#%% User Input

# Data Conversion Sheet, Column Sheet, Beam Sheet
input_xlsx_path = r'C:\Users\hwlee\Documents\하이웍스 받은파일\Data Conversion_Ver.2.0.xlsx'
# result_path = r'K:\2105-이형우\from 박재성\Results_E.Column'
# result_xlsx = 'Analysis Result'

# Analysis Result
result_xlsx_1 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\비선형해석모델\Results\105\KHSM_105_8_Analysis Result_1.xlsx'"
result_xlsx_2 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\비선형해석모델\Results\105\KHSM_105_8_Analysis Result_2.xlsx'"
result_xlsx_3 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\비선형해석모델\Results\105\KHSM_105_8_Analysis Result_3.xlsx'"
result_xlsx_path = result_xlsx_1 + ',' + result_xlsx_2 + ',' + result_xlsx_3  # + ',' + result_xlsx_4 + ',' + result_xlsx_5
result_xlsx_path = result_xlsx_path.split(',')
result_xlsx_path = [i.strip("'") for i in result_xlsx_path]
result_xlsx_path = [i.strip('"') for i in result_xlsx_path]
to_load_list = result_xlsx_path

DL_name = 'DL' # Midas Gen에서 Import해올 때 만든 하중 이름
LL_name = 'LL'

# Naming Option
drift_position = [2,5,7,11]

g_col_group_name = 'G.Column'

# Properties Assign 매크로 Option
# drag_duration = 0.15 # drag 하는 속도(너무 빨리하면 팅길 수 있으므로 적당한 속도 권장)
# offset = 2 # 픽셀 오차 방지용 여유치, 단위 : pixel
# start_material_index = 761 # wall_material_data_repeat 에서 시작하고자 하는 material name의 index 입력, 처음부터일때는 0 입력

###############################################################################
###############################################################################
#%% Import (MIDAS Gen -> Perform-3D)
# pbd.import_midas(input_xlsx_path)

#%% 이름 출력
# pbd.naming(input_xlsx_path)

#%% Properties 변환
pbd.convert_property(input_xlsx_path, get_beam=False, get_column=True, get_wall=False)

#%% Properties 변환 (기둥, Nu값 포함)
# pbd.convert_property_col_Nu(input_xlsx_path, result_path, result_xlsx=result_xlsx)

#%% Properties Assign 매크로

# pbd.property_assign_macro(drag_duration=0.15, offset=2, start_material_index=761)

#%% 시간 측정(END)
time_end = time.time()
time_run = (time_end-time_start)/60
print('\n', 'total time = %0.7f min' %(time_run))