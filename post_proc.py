import pandas as pd
import os
from collections import deque  # Double-ended Queue : 자료의 앞, 뒤 양 방향에서 자료를 추가하거나 제거가능
import matplotlib.pyplot as plt
from decimal import Decimal, ROUND_UP
import io
import pickle
from collections import deque

import PBD_p3d as pbd
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QSettings, QCoreApplication, QThread, QObject, Qt

#%% Base SF
class 

    
def base_SF(result_xlsx_path, ylim=70000):
    ''' 

    Perform-3D 해석 결과에서 각 지진파에 대한 Base층의 전단력을 막대그래프 형식으로 출력. (kN)
    
    Parameters
    ----------
    result_path : str
                  Perform-3D에서 나온 해석 파일의 경로.
                  
    result_xlsx : str, optional, default='Analysis Result'
                  Perform-3D에서 나온 해석 파일의 이름. 해당 파일 이름이 포함된 파일들을 모두 불러온다.
                  
    ylim : int, optional, default=70000
           그래프의 y축 limit 값. y축 limit 안의 값만 표기되므로, limit를 넘어가는 값을 확인하고 싶을 시에는 ylim 값을 더 크게 설정하면 된다.
    
    Returns
    -------
    '''
#%% Analysis Result 불러오기
    to_load_list = result_xlsx_path
    
    # 전단력 불러오기
    shear_force_data = pd.DataFrame()
    
    for i in to_load_list:
        result_data_raw = pd.ExcelFile(i)
        result_data_sheets = pd.read_excel(result_data_raw, ['Structure Section Forces'], skiprows=[0,2])
        
        column_name_to_slice = ['StrucSec Name', 'Load Case', 'Step Type', 'FH1', 'FH2']
        shear_force_data_temp = result_data_sheets['Structure Section Forces'].loc[:,column_name_to_slice]
        shear_force_data = pd.concat([shear_force_data, shear_force_data_temp])
        
    shear_force_data.columns = ['Name', 'Load Case', 'Step Type', 'H1(kN)', 'H2(kN)']
    
    # Base 전단력 추출
    shear_force_data = shear_force_data[shear_force_data['Name'].str.contains('base', case=False)]
      
    shear_force_data.reset_index(inplace=True, drop=True)
    
#%% 지진파 이름 list 만들기
    load_name_list = []
    for i in shear_force_data['Load Case'].drop_duplicates():
        new_i = i.split('+')[1]
        new_i = new_i.strip()
        load_name_list.append(new_i)
    
    gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
    seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]
    
    seismic_load_name_list.sort()
    
    DE_load_name_list = [x for x in load_name_list if 'DE' in x] # base shear로 사용할 지진파 개수 산정을 위함
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]

    # Marker 생성
    markers = []
    if len(DE_load_name_list) != 0:
        markers.append('DE')
    if len(MCE_load_name_list) != 0:
        markers.append('MCE')

#%% 데이터 Grouping
    shear_force_H1_data_grouped = pd.DataFrame()
    shear_force_H2_data_grouped = pd.DataFrame()
    
    for load_name in seismic_load_name_list:
        shear_force_H1_data_grouped['{}_H1_max'.format(load_name)] = shear_force_data[(shear_force_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (shear_force_data['Step Type'] == 'Max')]['H1(kN)'].values
            
        shear_force_H1_data_grouped['{}_H1_min'.format(load_name)] = shear_force_data[(shear_force_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (shear_force_data['Step Type'] == 'Min')]['H1(kN)'].values
    
    for load_name in seismic_load_name_list:
        shear_force_H2_data_grouped['{}_H2_max'.format(load_name)] = shear_force_data[(shear_force_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (shear_force_data['Step Type'] == 'Max')]['H2(kN)'].values
            
        shear_force_H2_data_grouped['{}_H2_min'.format(load_name)] = shear_force_data[(shear_force_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (shear_force_data['Step Type'] == 'Min')]['H2(kN)'].values   
    
    # all 절대값
    shear_force_H1_abs = shear_force_H1_data_grouped.abs()
    shear_force_H2_abs = shear_force_H2_data_grouped.abs()
    
    # Min, Max 중 최대값
    shear_force_H1_max = shear_force_H1_abs.groupby([[i//2 for i in range(0,len(seismic_load_name_list)*2)]], axis=1).max()
    shear_force_H2_max = shear_force_H2_abs.groupby([[i//2 for i in range(0,len(seismic_load_name_list)*2)]], axis=1).max()
    
    shear_force_H1_max.columns = seismic_load_name_list
    shear_force_H2_max.columns = seismic_load_name_list
    
    shear_force_H1_max.index = shear_force_data['Name'].drop_duplicates()
    shear_force_H2_max.index = shear_force_data['Name'].drop_duplicates()
    
    #%% Base Shear 그래프 그리기
    # Base Shear
    base_shear_H1 = shear_force_H1_max.copy()
    base_shear_H2 = shear_force_H2_max.copy()
    
    result = [base_shear_H1, base_shear_H2, DE_load_name_list, MCE_load_name_list, markers]

    return result