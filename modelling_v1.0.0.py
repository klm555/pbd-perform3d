import time
import PBD_p3d as pbd

#%% 시간 측정(START)
time_start = time.time()

###############################################################################
###############################################################################
#%% User Input

# Data Conversion Sheet, Column Sheet, Beam Sheet
input_path = r'D:\이형우\내진성능평가\광명 4R\103'
input_xlsx = 'Input Sheets(103_15)_v.1.8.xlsx'

DL = 'DL' # Midas Gen에서 Import해올 때 만든 하중 이름
LL = 'LL'

###############################################################################
###############################################################################
#%% Import (MIDAS Gen -> Perform-3D)

# pbd.import_midas(input_path, input_xlsx)

#%% 이름 출력


#%% Properties 변환

# pbd.convert_property(input_path, input_xlsx, get_beam=False, get_wall=True)
pbd.convert_property_reverse(input_path, input_xlsx, get_beam=False, get_wall=True)

#%% 시간 측정(END)
time_end = time.time()
time_run = (time_end-time_start)/60
print('\n', 'total time = %0.7f min' %(time_run))