import time
import PBD_p3d as pbd

#%% 시간 측정(START)
time_start = time.time()

###############################################################################
###############################################################################
#%% User Input

# Data Conversion Sheet, Column Sheet, Beam Sheet
input_path = r'C:\Users\hwlee\Desktop\Python\내진성능설계'
input_xlsx = 'Data Conversion_Ver.1.0_220930.xlsx'

DL = 'DL' # Midas Gen에서 Import해올 때 만든 하중 이름
LL = 'LL'

# Naming Option
drift_position = [2,5,7,11]

# Properties Assign 매크로 Option
drag_duration = 0.15 # drag 하는 속도(너무 빨리하면 팅길 수 있으므로 적당한 속도 권장)
offset = 2 # 픽셀 오차 방지용 여유치, 단위 : pixel
start_material_index = 761 # wall_material_data_repeat 에서 시작하고자 하는 material name의 index 입력, 처음부터일때는 0 입력

###############################################################################
###############################################################################
#%% Import (MIDAS Gen -> Perform-3D)
pbd.import_midas(input_path, input_xlsx)

#%% 이름 출력
pbd.naming(input_path, input_xlsx)

#%% Properties 변환
# pbd.convert_property(input_path, input_xlsx, get_beam=False, get_wall=True)
pbd.convert_property_reverse(input_path, input_xlsx, get_beam=False, get_wall=True)


#%% Properties Assign 매크로

pbd.property_assign_macro(drag_duration=0.15, offset=2, start_material_index=761)

#%% 시간 측정(END)
time_end = time.time()
time_run = (time_end-time_start)/60
print('\n', 'total time = %0.7f min' %(time_run))