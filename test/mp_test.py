import pandas as pd
from xlsx2csv import Xlsx2csv
from io import StringIO
# from joblib import Parallel, delayed
import dask.dataframe as dd
import dask.delayed as delayed

import time
import PBD_p3d as pbd

time_start = time.time()

input_xlsx_path = r'K:\2105-이형우\성능기반 내진설계\2_Excel_Sheets\Data Conversion_Ver.1.4 - 복사본.xlsx'

result_xlsx_1 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\비선형해석모델\Results\105\KHSM_105_8_Analysis Result_1.xlsx'"
result_xlsx_2 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\비선형해석모델\Results\105\KHSM_105_8_Analysis Result_2.xlsx'"
result_xlsx_3 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\비선형해석모델\Results\105\KHSM_105_8_Analysis Result_3.xlsx'"
result_xlsx_path = result_xlsx_1 + ',' + result_xlsx_2 + ',' + result_xlsx_3  # + ',' + result_xlsx_4 + ',' + result_xlsx_5
result_xlsx_path = result_xlsx_path.split(',')
result_xlsx_path = [i.strip("'") for i in result_xlsx_path]
result_xlsx_path = [i.strip('"') for i in result_xlsx_path]
to_load_list = result_xlsx_path


##### Excel 파일 읽는 Function (w/ Xlsx2csv & joblib)
def read_excel(path:str, sheet_name:str, skip_rows:list=[0,2]) -> pd.DataFrame:
    data_buffer = StringIO()
    Xlsx2csv(path, outputencoding="utf-8").convert(data_buffer, sheetname=sheet_name)
    data_buffer.seek(0)
    data_df = pd.read_csv(data_buffer, low_memory=False, skiprows=skip_rows)
    return data_df

##### Read Excel Files (Data Conversion Sheets & Analysis Result Sheets)
# Input_G.Beam
gbeam = read_excel(input_xlsx_path, sheet_name='Input_G.Beam', skip_rows=[0,1,2])
gbeam = gbeam.iloc[:,0]
gbeam.dropna(inplace=True, how='all')
gbeam.name = 'Property Name'

# Input_G.Column
gcol = read_excel(input_xlsx_path, sheet_name='Input_G.Column', skip_rows=[0,1,2])
gcol = gcol.iloc[:,0]
gcol.dropna(inplace=True, how='all')
gcol.name = 'Property Name'
# Input_E.Column
ecol = read_excel(input_xlsx_path, sheet_name='Input_E.Column', skip_rows=[0,1,2])
ecol = ecol.iloc[:,0]
ecol.dropna(inplace=True, how='all')
ecol.name = 'Property Name'

# Elements(Frame)
element_data = read_excel(to_load_list[0], 'Element Data - Frame Types')
column_name_to_slice = ['Element Name', 'Property Name', 'I-Node ID', 'J-Node ID']
element_data = element_data.loc[:, column_name_to_slice]    

# Forces (Vu, Nu)
# Using joblib (occurs an error "NoneType has no attribute 'write'")
# beam_force_data = Parallel(n_jobs=-1, verbose=10)(delayed(read_excel)(file_path, 'Frame Results - End Forces') for file_path in to_load_list)
# Using dask
beam_force_parts = [dask.delayed(read_excel)(file_path, 'Frame Results - End Forces') 
         for file_path in to_load_list]
beam_force_data = dd.from_delayed(beam_force_parts, meta=beam_force_parts[0].compute())

beam_force_data = pd.concat(beam_force_data, ignore_index=True)
column_name_to_slice = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'P J-End', 'V2 I-End', 'V2 J-End']
beam_force_data = beam_force_data.loc[:, column_name_to_slice]

##### Create Seismic Loads List
load_name_list = []
for i in beam_force_data['Load Case'].drop_duplicates():
    new_i = i.split('+')[1]
    new_i = new_i.strip()
    load_name_list.append(new_i)
gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]
seismic_load_name_list.sort()        
DE_load_name_list = [x for x in load_name_list if 'DE' in x]
MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]

##### Merge Result Data & Element Data
beam_force_data = pd.merge(beam_force_data, element_data, how='left')
##### Slice only Data from Gravitaional Loads
beam_force_data = beam_force_data[beam_force_data['Load Case'].str.contains(gravity_load_name[0])]
beam_force_data.reset_index(inplace=True, drop=True)




#%%
import pandas as pd
from xlsx2csv import Xlsx2csv
from io import StringIO
from joblib import Parallel, delayed
import dask.dataframe as dd
import dask
import multiprocessing as mp
from itertools import repeat

import time
import PBD_p3d as pbd

time_start = time.time()

input_xlsx_path = r'K:\2105-이형우\성능기반 내진설계\2_Excel_Sheets\Data Conversion_Ver.1.4 - 복사본.xlsx'
result_xlsx_1 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\비선형해석모델\Results\105\KHSM_105_8_Analysis Result_1.xlsx'"
result_xlsx_2 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\비선형해석모델\Results\105\KHSM_105_8_Analysis Result_2.xlsx'"
result_xlsx_3 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\비선형해석모델\Results\105\KHSM_105_8_Analysis Result_3.xlsx'"
result_xlsx_path = result_xlsx_1 + ',' + result_xlsx_2 + ',' + result_xlsx_3  # + ',' + result_xlsx_4 + ',' + result_xlsx_5
result_xlsx_path = result_xlsx_path.split(',')
result_xlsx_path = [i.strip("'") for i in result_xlsx_path]
result_xlsx_path = [i.strip('"') for i in result_xlsx_path]
to_load_list = result_xlsx_path

##### Excel 파일 읽는 Function (w/ Xlsx2csv & joblib)
def read_excel(path:str, sheet_name:str, skip_rows:list=[0,2]) -> pd.DataFrame:
    data_buffer = StringIO()
    Xlsx2csv(path, outputencoding="utf-8").convert(data_buffer, sheetname=sheet_name)
    data_buffer.seek(0)
    data_df = pd.read_csv(data_buffer, low_memory=False, skiprows=skip_rows)
    return data_df

### Joblib
beam_force_data = Parallel(n_jobs=-1, verbose=10)(delayed(read_excel)(file_path, 'Frame Results - End Forces') for file_path in to_load_list)
beam_force_data = pd.concat(beam_force_data, ignore_index=True)
column_name_to_slice = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'P J-End', 'V2 I-End', 'V2 J-End']
beam_force_data = beam_force_data.loc[:, column_name_to_slice]

### Multiprocessing
# cpu_num = mp.cpu_count()
# pool = mp.Pool(processes=cpu_num)
# beam_force_data = pool.map(read_excel, [(file, 'Frame Results - End Forces') for file in to_load_list])
# 메모리 릭 방지 위해 사용

### Dask
# beam_force_parts = [delayed(read_excel)(file_path, 'Frame Results - End Forces').compute() 
#                     for file_path in to_load_list]
# beam_force_data = pd.concat(beam_force_parts, ignore_index=True)
# column_name_to_slice = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'P J-End', 'V2 I-End', 'V2 J-End']
# beam_force_data = beam_force_data.loc[:, column_name_to_slice]

time_end = time.time()
time_run = (time_end-time_start)/60
print('\n', 'total time = %0.7f min' %(time_run))


#%%
import multiprocessing as mp
from itertools import product

def merge_names(a, b):
    return '{} & {}'.format(a, b)

cpu_num = mp.cpu_count()
pool = mp.Pool(processes=cpu_num)
names = ['Brown', 'Wilson', 'Bartlett', 'Rivera', 'Molloy', 'Opie']
results = pool.map(merge_names, [(i, 'James') for i in names])
print(results)
