import time
import PBD_p3d as pbd

#%% 시간 측정(START)
time_start = time.time()

###############################################################################
###############################################################################
#%% User Input

# Data Conversion Sheet, Column Sheet, Beam Sheet
input_path = r'D:\이형우\내진성능평가\광명 4R\102\test'
input_xlsx = '102_9_Data Conversion_Shear Wall Type_Ver.1.7.xlsx'

DL = 'DL' # Midas Gen에서 Import해올 때 만든 하중 이름
LL = 'LL'

###############################################################################
###############################################################################
#%% Import (MIDAS Gen -> Perform-3D)

pbd.import_midas(input_path, input_xlsx, DL_name=DL, LL_name=LL\
                       , import_mass=False, import_AS_gage=False, import_plate=False)

#%% 이름 출력

    

#%% 시간 측정(END)
time_end = time.time()
time_run = (time_end-time_start)/60
print('\n', 'total time = %0.7f min' %(time_run))