import pandas as pd
import os
from io import StringIO
from xlsx2csv import Xlsx2csv
from joblib import Parallel, delayed
from collections import deque  # Double-ended Queue : 자료의 앞, 뒤 양 방향에서 자료를 추가하거나 제거가능
import matplotlib.pyplot as plt
from decimal import Decimal, ROUND_UP
import io
import pickle
from collections import deque

import PBD_p3d as pbd
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QSettings, QCoreApplication, QThread, QObject, Qt

#%% Post-Processing Class

class PostProc():
    
    def __init__(self, input_xlsx_path, result_xlsx_path):
        
        ##### Load Excel Files (Analysis Result Sheets)
        to_load_list = result_xlsx_path    

        ##### Excel 파일 읽는 Function (w/ Xlsx2csv & joblib)
        def read_excel(path:str, sheet_name:str, skip_rows:list=[0,2]) -> pd.DataFrame:
            data_buffer = StringIO()
            Xlsx2csv(path, outputencoding="utf-8").convert(data_buffer, sheetname=sheet_name)
            data_buffer.seek(0)
            data_df = pd.read_csv(data_buffer, low_memory=False, skiprows=skip_rows)
            return data_df
        
        ##### Read Excel Files (Data Conversion Sheets & Analysis Result Sheets)
        # Story Info
        self.story_info = read_excel(input_xlsx_path, sheet_name='Story Data', skip_rows=[0,1,2])
        self.story_info = self.story_info.iloc[:,[0,1,2]]
        self.story_info.columns = ['Index', 'Story Name', 'Height(mm)']
        self.story_info.dropna(how='all', inplace=True)
        # Output_Wall Properties
        self.wall_info = read_excel(input_xlsx_path, sheet_name='Output_Wall Properties', skip_rows=[0,1,2])
        self.wall_info = self.wall_info.iloc[:,0:10]
        self.wall_info.columns = ['Name', 'Length(mm)', 'Thickness(mm)', 'Concrete Grade', 'Rebar Type', 'V.Rebar Type'
                                  , 'V.Rebar Spacing(mm)', 'V.Rebar EA', 'H.Rebar Type', 'H.Rebar Spacing(mm)']
        self.wall_info.dropna(how='all', inplace=True)
        # Wall Deformation Capacities
        # self.wall_deform_cap = read_excel(input_xlsx_path, sheet_name='Results_Wall', skip_rows=[0,1,2])
        # self.wall_deform_cap = self.wall_deform_cap.iloc[:,[0,11,12,13,14,48,49,54,55]]
        # self.wall_deform_cap.columns = ['Name', 'Vu_DE_H1', 'Vu_DE_H2', 'Vu_MCE_H1', 'Vu_MCE_H2'
        #                                 , 'LS(H1)', 'LS(H2)', 'CP(H1)', 'CP(H2)']
        # C.Beam Deformation Capacities
        self.beam_deform_cap = read_excel(input_xlsx_path, sheet_name='Output_C.Beam Properties', skip_rows=[0,1,2])
        self.beam_deform_cap = self.beam_deform_cap.iloc[:,[0,80,81]]
        self.beam_deform_cap.columns = ['Name', 'LS', 'CP']
        # G.Column Deformation Capacities
        self.col_deform_cap = read_excel(input_xlsx_path, sheet_name='Output_G.Column Properties', skip_rows=[0,1,2])
        self.col_deform_cap = self.col_deform_cap.iloc[:,[0,80,81,82,83]]
        self.col_deform_cap.columns = ['Name', 'LS(X)', 'LS(Y)', 'CP(X)', 'CP(Y)']

        # Nodes
        self.node_data = read_excel(to_load_list[0], 'Node Coordinate Data')
        column_name_to_slice = ['Node ID', 'H1', 'H2', 'V']
        self.node_data = self.node_data.loc[:, column_name_to_slice]
        # Elements(Wall)
        self.wall_data = read_excel(to_load_list[0], 'Element Data - Shear Wall')
        column_name_to_slice = ['Element Name', 'Property Name', 'I-Node ID', 'J-Node ID', 'K-Node ID', 'L-Node ID']
        self.wall_data = self.wall_data.loc[:, column_name_to_slice]
        # Elements(Frame)
        self.frame_data = read_excel(to_load_list[0], 'Element Data - Frame Types')
        column_name_to_slice = ['Element Name', 'Property Name', 'I-Node ID', 'J-Node ID']
        self.frame_data = self.frame_data.loc[:, column_name_to_slice]
        # Wall Axial Strain Gage
        self.wall_as_gage_data = read_excel(to_load_list[0], 'Gage Data - Bar Type')
        column_name_to_slice = ['Group Name', 'Element Name', 'I-Node ID', 'J-Node ID']
        self.wall_as_gage_data = self.wall_as_gage_data.loc[:, column_name_to_slice]
        # Wall Rotation Gage
        self.wall_rot_gage_data = read_excel(to_load_list[0], 'Gage Data - Wall Type')
        column_name_to_slice = ['Element Name', 'I-Node ID', 'J-Node ID', 'K-Node ID', 'L-Node ID']
        self.wall_rot_gage_data = self.wall_rot_gage_data.loc[:, column_name_to_slice]

        # Wall Shear Force
        self.shear_force_data = Parallel(n_jobs=-1, verbose=10)(delayed(read_excel)(file_path, 'Structure Section Forces') for file_path in to_load_list)
        self.shear_force_data = pd.concat(self.shear_force_data, ignore_index=True)
        column_name_to_slice = ['StrucSec Name', 'Load Case', 'Step Type', 'FH1', 'FH2', 'FV']
        self.shear_force_data = self.shear_force_data.loc[:, column_name_to_slice]
        self.shear_force_data.columns = ['Name', 'Load Case', 'Step Type', 'H1(kN)', 'H2(kN)', 'V(kN)']
        # Inter-Story Drift
        self.drift_data = Parallel(n_jobs=-1, verbose=10)(delayed(read_excel)(file_path, 'Drift Output') for file_path in to_load_list)
        self.drift_data = pd.concat(self.drift_data, ignore_index=True)
        column_name_to_slice = ['Drift Name', 'Drift ID', 'Load Case', 'Step Type', 'Drift']
        self.drift_data = self.drift_data.loc[:, column_name_to_slice]
        # Wall Axial Strain Result
        self.wall_as_result_data = Parallel(n_jobs=-1, verbose=10)(delayed(read_excel)(file_path, 'Gage Results - Bar Type') for file_path in to_load_list)
        self.wall_as_result_data = pd.concat(self.wall_as_result_data, ignore_index=True)
        column_name_to_slice = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'Axial Strain', 'Performance Level']
        self.wall_as_result_data = self.wall_as_result_data.loc[:, column_name_to_slice]
        # Wall Rotation Result
        self.wall_rot_result_data = Parallel(n_jobs=-1, verbose=10)(delayed(read_excel)(file_path, 'Gage Results - Wall Type') for file_path in to_load_list)
        self.wall_rot_result_data = pd.concat(self.wall_rot_result_data, ignore_index=True)
        column_name_to_slice = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'Rotation', 'Performance Level']
        self.wall_rot_result_data = self.wall_rot_result_data.loc[:, column_name_to_slice]
        # Beam Rotation
        self.beam_rot_data = Parallel(n_jobs=-1, verbose=10)(delayed(read_excel)(file_path, 'Frame Results - Bending Deform') for file_path in to_load_list)
        self.beam_rot_data = pd.concat(self.beam_rot_data, ignore_index=True)
        column_name_to_slice = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'Distance from I-End', 'R2', 'R3']
        self.beam_rot_data = self.beam_rot_data.loc[:, column_name_to_slice]
        self.beam_rot_data.columns = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'Distance from I-End', 'H2 Rotation(rad)', 'H3 Rotation(rad)']
        # Beam Shear Force
        self.beam_shear_force_data = Parallel(n_jobs=-1, verbose=10)(delayed(read_excel)(file_path, 'Frame Results - End Forces') for file_path in to_load_list)
        self.beam_shear_force_data = pd.concat(self.beam_shear_force_data, ignore_index=True)
        column_name_to_slice = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'V2 I-End', 'V3 I-End']
        self.beam_shear_force_data = self.beam_shear_force_data.loc[:, column_name_to_slice]

        ##### Create Seismic Loads List
        self.load_name_list = []
        for i in self.drift_data['Load Case'].drop_duplicates():
            new_i = i.split('+')[1]
            new_i = new_i.strip()
            self.load_name_list.append(new_i)
        self.gravity_load_name = [x for x in self.load_name_list if ('DE' not in x) and ('MCE' not in x)]
        self.seismic_load_name_list = [x for x in self.load_name_list if ('DE' in x) or ('MCE' in x)]
        self.seismic_load_name_list.sort()        
        self.DE_load_name_list = [x for x in self.load_name_list if 'DE' in x] # base shear로 사용할 지진파 개수 산정을 위함
        self.MCE_load_name_list = [x for x in self.load_name_list if 'MCE' in x]

    ##### Import Post-Processing Methods
    from .system import base_SF, story_SF, IDR
    from .wall import WAS, WR, WSF
    from .beam import BR, BSF
    from .column import CR, CSF

    # 
    # def write_xlsx()
