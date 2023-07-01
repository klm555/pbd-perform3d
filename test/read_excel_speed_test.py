#%% Import

import os
import pandas as pd
import time
from io import BytesIO, StringIO # 파일처럼 취급되는 문자열 객체 생성(메모리 낭비 down)
import multiprocessing as mp
from joblib import Parallel, delayed
from openpyxl import load_workbook
from xlsx2csv import Xlsx2csv
import PBD_p3d.output_to_docx as otd
import PBD_p3d as pbd
###########################   FILE 경로    ####################################
# Analysis Result
result_xlsx_1 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\비선형해석모델\Results\103\KHSM_103_2_Analysis Result_1.xlsx'"
result_xlsx_2 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\비선형해석모델\Results\103\KHSM_103_2_Analysis Result_2.xlsx'"
result_xlsx_3 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\비선형해석모델\Results\103\KHSM_103_2_Analysis Result_3.xlsx'"

result_xlsx_path = result_xlsx_1 + ',' + result_xlsx_2 + ',' + result_xlsx_3
result_xlsx_path = result_xlsx_path.split(',')
result_xlsx_path = [i.strip("'") for i in result_xlsx_path]
result_xlsx_path = [i.strip('"') for i in result_xlsx_path]
to_load_list = result_xlsx_path

#%% DATA ANALYSIS
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

#%% DATA ANALYSIS (with parse)
beam_rot_data = pd.DataFrame()

for i in to_load_list:
    result_data_raw = pd.ExcelFile(i)
    result_data_sheets = result_data_raw.parse(['Frame Results - Bending Deform', 'Node Coordinate Data',\
                                                        'Element Data - Frame Types'], skiprows=[0,2])
    
    beam_rot_data_temp = result_data_sheets['Frame Results - Bending Deform'].iloc[:,[0,2,5,7,10,13,14]]
    beam_rot_data = pd.concat([beam_rot_data, beam_rot_data_temp])
    
node_data = result_data_sheets['Node Coordinate Data'].iloc[:,[1,4]]

element_data = result_data_sheets['Element Data - Frame Types'].iloc[:,[0,2,5,7]] # beam의 양 nodes중 한 node에서의 rotation * 2

beam_rot_data.columns = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'Distance from I-End', 'H2 Rotation(rad)', 'H3 Rotation(rad)']
node_data.columns = ['Node ID', 'V(mm)']
element_data.columns = ['Group Name', 'Element Name', 'Property Name', 'I-Node ID']

#%% DATA ANALYSIS (with xlsx2csv & multiprocessing)
beam_rot_data = pd.DataFrame()

def read_excel(path: str, sheet_name: str) -> pd.DataFrame:
    buffer = StringIO()
    Xlsx2csv(path, outputencoding="utf-8").convert(buffer, sheetname=sheet_name)
    buffer.seek(0)
    df = pd.read_csv(buffer, low_memory=False, skiprows=[0,2])
    return df

beam_rot_data  = Parallel(n_jobs=-1, verbose=10)(delayed(read_excel)(file_path, 'Frame Results - Bending Deform') for file_path in to_load_list)
beam_rot_data = pd.concat(beam_rot_data, ignore_index=True)
beam_rot_data = beam_rot_data.iloc[:,[0,2,5,7,10,13,14]]

# for i in to_load_list:
#     beam_rot_data_temp = read_excel(i, 'Frame Results - Bending Deform').iloc[:,[0,2,5,7,10,13,14]]
#     beam_rot_data = pd.concat([beam_rot_data, beam_rot_data_temp])

node_data = read_excel(to_load_list[0], 'Node Coordinate Data').iloc[:,[1,4]]

element_data = read_excel(to_load_list[0], 'Element Data - Frame Types').iloc[:,[0,2,5,7]]

beam_rot_data.columns = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'Distance from I-End', 'H2 Rotation(rad)', 'H3 Rotation(rad)']
node_data.columns = ['Node ID', 'V(mm)']
element_data.columns = ['Group Name', 'Element Name', 'Property Name', 'I-Node ID']

#%% DATA ANALYSIS (with multiprocessing)
beam_rot_data = pd.DataFrame()

def read_xlsx_parallel(file_path):
    result_data_raw = pd.ExcelFile(file_path)
    result_data_sheets = pd.read_excel(result_data_raw, ['Frame Results - Bending Deform', 'Node Coordinate Data',\
                                                        'Element Data - Frame Types'], skiprows=[0,2])
    beam_rot_data = result_data_sheets['Frame Results - Bending Deform'].iloc[:,[0,2,5,7,10,13,14]]
    return result_data_sheets, beam_rot_data


beam_rot_data  = Parallel(n_jobs=-1, verbose=10)(delayed(read_xlsx_parallel)(file_path) for file_path in to_load_list)

# df = pd.concat(df, ignore_index=True)