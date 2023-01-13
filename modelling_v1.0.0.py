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
input_xlsx_path = r'C:\Users\hwlee\Desktop\Python\내진성능설계\Data Conversion_Ver.1.0_220930.xlsx'

DL = 'DL' # Midas Gen에서 Import해올 때 만든 하중 이름
LL = 'LL'

# Naming Option
drift_position = [2,5,7,11]

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

#%% Properties Assign 매크로

# pbd.property_assign_macro(drag_duration=0.15, offset=2, start_material_index=761)

#%% 시간 측정(END)
time_end = time.time()
time_run = (time_end-time_start)/60
print('\n', 'total time = %0.7f min' %(time_run))