import pandas as pd
import numpy as np
import os
import pickle
import multiprocess as mp
import win32com.client
import pythoncom
from PyPDF2 import PdfMerger, PdfFileReader

from PyQt5.QtWidgets import *
from PyQt5.QtCore import QSettings, QCoreApplication, QThread, QObject, Qt

#%% Post-Processing Class

class PostProc():
    
    def __init__(self, input_xlsx_path, result_xlsx_path, get_base_SF=False
                 , get_story_SF=False, get_IDR=False, get_BR=False, get_BSF=False
                 , get_E_BSF=False, get_CR=False, get_CSF=False, get_E_CSF=False
                 , get_WAS=False, get_WR=False, get_WSF=False):
        
        ##### Load Excel Files (Analysis Result Sheets)
        to_load_list = result_xlsx_path    

        ##### Excel 파일 읽는 Function (w/ Xlsx2csv & joblib)
        def read_excel(path:str, sheet_name:str, skip_rows:list=[0,2]) -> pd.DataFrame:
            import pandas as pd
            from io import StringIO # if not import, error occurs when using multiprocessing
            from xlsx2csv import Xlsx2csv
            data_buffer = StringIO()
            Xlsx2csv(path, outputencoding="utf-8", ignore_formats='float').convert(data_buffer, sheetname=sheet_name)
            data_buffer.seek(0)
            data_df = pd.read_csv(data_buffer, low_memory=False, skiprows=skip_rows)
            return data_df
        
        ##### Read Excel Files (Data Conversion Sheets & Analysis Result Sheets)
        # Story Info
        self.story_info = read_excel(input_xlsx_path, sheet_name='Story Data', skip_rows=[0,1,2])
        self.story_info = self.story_info.iloc[:,[0,1,2]]
        self.story_info.dropna(how='all', inplace=True)
        self.story_info.columns = ['Index', 'Story Name', 'Height(mm)']
        # Rebar Info
        if get_E_CSF == True:
            self.rebar_info = read_excel(input_xlsx_path, sheet_name='ETC', skip_rows=[0,1,2])
            self.rebar_info = self.rebar_info.iloc[:,[0,3,4]]
            self.rebar_info.dropna(how='all', inplace=True)
            self.rebar_info.columns = ['Type', '일반용', '내진용']
        # Output_Wall Properties
        if (get_WR == True) | (get_WSF == True):
            self.wall_info = read_excel(input_xlsx_path, sheet_name='Input_S.Wall', skip_rows=[0,1,2])
            self.wall_info = self.wall_info.iloc[:,0:11]
            self.wall_info.dropna(how='all', inplace=True)
            self.wall_info.columns = ['Name', 'Length(mm)', 'Thickness(mm)', 'Concrete Grade', 'V.Rebar Type', 'V.Rebar(DXX)'
                                      , 'V.Rebar Spacing(mm)', 'V.Rebar EA', 'H.Rebar Type', 'H.Rebar(DXX)', 'H.Rebar Spacing(mm)']
        # C.Beam Info
        if (get_BR == True) | (get_BSF == True):
            self.beam_info = read_excel(input_xlsx_path, sheet_name='Input_C.Beam', skip_rows=[0,1,2])
            self.beam_info = self.beam_info.iloc[:,0:31]
            self.beam_info.dropna(how='all', inplace=True)
            self.beam_info.columns = ['Name', 'Length(mm)', 'b(mm)', 'h(mm)', 'd(mm)', 'Concrete Grade'
                                      , 'Arrangement', 'Seismic Detail', 'Main Rebar Type', 'Main Rebar(DXX)'
                                      , 'Stirrup Type', 'Stirrup(DXX)', 'X-Bracing Type', 'X-Bracing(DXX)'
                                      , 'Top(1)', 'Top(2)', 'Top(3)', 'Stirrup EA', 'Stirrup Space(mm)'
                                      , 'X-Bracing EA', 'X-Bracing deg', 'Main Rebar(DXX)_after', 'Stirrup(DXX)_after'
                                      , 'X-Bracing(DXX)_after', 'Top(1)_after', 'Top(2)_after', 'Top(3)_after'
                                      , 'Stirrup EA_after', 'Stirrup Space(mm)_after', 'X-Bracing EA_after', 'X-Bracing deg_after']
        # E.Column Info
        if get_E_CSF == True:
            self.ecol_info = read_excel(input_xlsx_path, sheet_name='Input_E.Column', skip_rows=[0,1,2])
            self.ecol_info = self.ecol_info.iloc[:,np.r_[0:15,18]]
            self.ecol_info.dropna(how='all', inplace=True)
            self.ecol_info.columns = ['Name', 'b(mm)', 'h(mm)', 'Concrete Grade'
                                      , 'Main Rebar Type', 'Main Rebar(DXX)', 'Hoop Type'
                                      , 'Hoop(DXX)', 'Layer1 EA', 'Layer1 Row', 'Layer2 EA'
                                      , 'Layer2 Row', 'Hoop X', 'Hoop Y', 'Hoop Space(mm)', 'Direction']

        # Nodes
        if (get_BR == True) | (get_BSF == True) | (get_WAS == True) | (get_WR == True) | (get_WSF == True)\
            | (get_E_CSF == True):
            self.node_data = read_excel(to_load_list[0], 'Node Coordinate Data')
            column_name_to_slice = ['Node ID', 'H1', 'H2', 'V']
            self.node_data = self.node_data.loc[:, column_name_to_slice]
        # Elements(Wall)
        if (get_WR == True) | (get_WSF == True):
            self.wall_data = read_excel(to_load_list[0], 'Element Data - Shear Wall')
            column_name_to_slice = ['Element Name', 'Property Name', 'I-Node ID', 'J-Node ID', 'K-Node ID', 'L-Node ID']
            self.wall_data = self.wall_data.loc[:, column_name_to_slice]
        # Elements(Frame)
        if (get_BR == True) | (get_BSF == True) | (get_E_CSF == True):
            self.frame_data = read_excel(to_load_list[0], 'Element Data - Frame Types')
            column_name_to_slice = ['Element Name', 'Property Name', 'I-Node ID', 'J-Node ID']
            self.frame_data = self.frame_data.loc[:, column_name_to_slice]
        # Wall Axial Strain Gage
        if get_WAS == True:
            self.wall_as_gage_data = read_excel(to_load_list[0], 'Gage Data - Bar Type')
            column_name_to_slice = ['Group Name', 'Element Name', 'I-Node ID', 'J-Node ID']
            self.wall_as_gage_data = self.wall_as_gage_data.loc[:, column_name_to_slice]
        # Wall Rotation Gage
        if (get_WR == True):
            self.wall_rot_gage_data = read_excel(to_load_list[0], 'Gage Data - Wall Type')
            column_name_to_slice = ['Element Name', 'I-Node ID', 'J-Node ID', 'K-Node ID', 'L-Node ID']
            self.wall_rot_gage_data = self.wall_rot_gage_data.loc[:, column_name_to_slice]

        # Using multiprocess (library which overcomes the issue made ny using 'pickle' in 'multiprocessing' library)
        cpu_num = mp.cpu_count() # Count the # of CPU
        pool = mp.Pool(processes=cpu_num) # Create a pool equal to the # of CPU
        # Inter-Story Drift
        self.drift_data = pool.starmap(read_excel, [[file_path, 'Drift Output'] for file_path in to_load_list])
        self.drift_data = pd.concat(self.drift_data, ignore_index=True)
        column_name_to_slice = ['Drift Name', 'Drift ID', 'Load Case', 'Step Type', 'Drift']
        self.drift_data = self.drift_data.loc[:, column_name_to_slice]
        # Wall Shear Force
        if (get_base_SF == True) | (get_story_SF == True) | (get_WR == True) | (get_WSF == True):
            self.shear_force_data = pool.starmap(read_excel, [[file_path, 'Structure Section Forces'] for file_path in to_load_list])
            self.shear_force_data = pd.concat(self.shear_force_data, ignore_index=True)
            column_name_to_slice = ['StrucSec Name', 'Load Case', 'Step Type', 'FH1', 'FH2', 'FV']
            self.shear_force_data = self.shear_force_data.loc[:, column_name_to_slice]
            self.shear_force_data.columns = ['Name', 'Load Case', 'Step Type', 'H1(kN)', 'H2(kN)', 'V(kN)']
        # Wall Axial Strain Result
        if get_WAS == True:
            self.wall_as_result_data = pool.starmap(read_excel, [[file_path, 'Gage Results - Bar Type'] for file_path in to_load_list])
            self.wall_as_result_data = pd.concat(self.wall_as_result_data, ignore_index=True)
            column_name_to_slice = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'Axial Strain', 'Performance Level']
            self.wall_as_result_data = self.wall_as_result_data.loc[:, column_name_to_slice]
        # Wall Rotation Result
        if (get_WR == True):
            self.wall_rot_result_data= pool.starmap(read_excel, [[file_path, 'Gage Results - Wall Type'] for file_path in to_load_list])
            self.wall_rot_result_data = pd.concat(self.wall_rot_result_data, ignore_index=True)
            column_name_to_slice = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'Rotation', 'Performance Level']
            self.wall_rot_result_data = self.wall_rot_result_data.loc[:, column_name_to_slice]
        # Beam Rotation
        if get_BR == True:
            self.beam_rot_data = pool.starmap(read_excel, [[file_path, 'Frame Results - Bending Deform'] for file_path in to_load_list])
            self.beam_rot_data = pd.concat(self.beam_rot_data, ignore_index=True)
            column_name_to_slice = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'Point ID', 'R2', 'R3']
            self.beam_rot_data = self.beam_rot_data.loc[:, column_name_to_slice]
            self.beam_rot_data.columns = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'Point ID', 'H2 Rotation(rad)', 'H3 Rotation(rad)']
        # Beam Shear Force
        if (get_BSF == True) | (get_E_CSF == True):
            self.beam_shear_force_data = pool.starmap(read_excel, [[file_path, 'Frame Results - End Forces'] for file_path in to_load_list])
            self.beam_shear_force_data = pd.concat(self.beam_shear_force_data, ignore_index=True)
            column_name_to_slice = ['Group Name', 'Element Name', 'Load Case', 'Step Type'
                                    , 'P I-End', 'V2 I-End', 'V2 J-End', 'V3 I-End', 'M2 I-End', 'M3 I-End']
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
        self.DE_load_name_list.sort()
        self.MCE_load_name_list.sort() 

        # pkl 폴더 생성
        def create_dir(directory):
            try:
                if not os.path.exists(directory):
                    os.makedirs(directory)
            except OSError:
                print("Error: Failed to create the directory.")                
        create_dir('pkl')        

    ##### Import Post-Processing Methods
    # from PBD_p3d.system import base_SF, story_SF, IDR
    # from PBD_p3d.wall import WAS, WR, WSF
    # from PBD_p3d.beam import BR, BSF
    # from PBD_p3d.column import CR, CSF
    from .system import base_SF, story_SF, IDR
    from .wall import WAS, WR, WSF
    from .beam import BR, BSF
    from .column import CR, CSF, E_CSF

#%% Function to Print the Result into PDF

def print_pdf(beam_design_xlsx_path, col_design_xlsx_path
              , wall_design_xlsx_path, get_cbeam=False, get_ecol=False, get_wall=False):  

    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = False # 엑셀창 안보이게

    if get_ecol == True:
        wb_ecol = excel.Workbooks.Open(col_design_xlsx_path)
        ws = wb_ecol.Sheets('Design_E.Column')        
        startrow, startcol = 5, 1
        
        # 부재 개수(for iterration) 구하기
        ws_row_num = ws.UsedRange.Rows.Count        
        element_name = ws.Range('B%s:B%s' %(startrow, ws_row_num)).Value
        element_name_df = pd.DataFrame(element_name)

        # Drop NoneType object & Rebar Diameter in the end of Data
        element_name_df.iloc[:,0] = element_name_df[element_name_df.iloc[:,0].str.count('_') == 2]
        element_name_df = element_name_df.dropna()
        
        # 부재개수
        element_num = element_name_df.shape[0]
        
        # Path 지정
        result_path = os.path.splitext(col_design_xlsx_path)[0] # 확장자명(extension) 제거
        
        # pdf Merge를 위한 PdfMerger 클래스 생성
        merger = PdfMerger()

        for i in range(element_num):

            wb_ecol.Worksheets(3).Select()            
            wb_ecol.Worksheets(3).Name = '({})'.format(i+1)

            # wb_ecol.SaveAs(pdf_file_path, FileFormat=57)

            xlTypePDF = 0
            xlQualityStandard = 0
            
            # 왜인지 모르겠지만 result_path에 suffix 붙이면 \가 /로 바뀜... 그래서 다시 바꿔주기
            pdf_file_path = result_path + '({}).pdf'.format(i+1)
            pdf_file_path = pdf_file_path.replace('/', '\\')            
            wb_ecol.ActiveSheet.ExportAsFixedFormat(xlTypePDF, pdf_file_path\
                                                , xlQualityStandard, True, False)    

            merger.append(pdf_file_path)
            
        merger.write(result_path + '.pdf')
        merger.close()
        
        # Merge한 후 개별 파일들 지우기    
        for i in range(element_num):
            pdf_file_path = result_path + '({}).pdf'.format(i+1)
            pdf_file_path = pdf_file_path.replace('/', '\\')
            os.remove(pdf_file_path)

        # wb_ecol.Close(SaveChanges=False)
        # wb_ecol.Save()
        
    if get_wall == True:
        wb_wall = excel.Workbooks.Open(wall_design_xlsx_path)
        ws_DE = wb_wall.Sheets('Plot_S.Wall_DE')
        ws_MCE = wb_wall.Sheets('Plot_S.Wall_MCE')
        ws_row_num = wb_wall.Sheets('Table_S.Wall_DE')
        startrow, startcol = 5, 1
        
        # 부재 개수(for iterration) 구하기
        element_num = ws_row_num.Range('A4:A4').Value
        
        # Path 지정
        result_path = os.path.splitext(wall_design_xlsx_path)[0] # 확장자명(extension) 제거
        
        # pdf Merge를 위한 PdfMerger 클래스 생성
        merger = PdfMerger()

        for i in range(element_num):

            xlTypePDF = 0
            xlQualityStandard = 0
            
            ws_DE.Range('A8:A8').Value = i + 1
            
            # 왜인지 모르겠지만 result_path에 suffix 붙이면 \가 /로 바뀜... 그래서 다시 바꿔주기
            pdf_file_path = result_path + '_DE({}).pdf'.format(i + 1)
            pdf_file_path = pdf_file_path.replace('/', '\\')
            wb_wall.ActiveSheet.ExportAsFixedFormat(xlTypePDF, pdf_file_path\
                                                , xlQualityStandard, True, False)    

            merger.append(pdf_file_path)
            
        for i in range(element_num):

            xlTypePDF = 0
            xlQualityStandard = 0
            
            ws_MCE.Range('A8:A8').Value = i + 1
            
            # 왜인지 모르겠지만 result_path에 suffix 붙이면 \가 /로 바뀜... 그래서 다시 바꿔주기
            pdf_file_path = result_path + '_MCE({}).pdf'.format(i + 1)
            pdf_file_path = pdf_file_path.replace('/', '\\')
            wb_wall.ActiveSheet.ExportAsFixedFormat(xlTypePDF, pdf_file_path\
                                                , xlQualityStandard, True, False)    

            merger.append(pdf_file_path)
            
        merger.write(result_path + '.pdf')
        merger.close()
        
        # Merge한 후 개별 파일들 지우기    
        for i in range(element_num):
            DE_pdf_file_path = result_path + '_DE({}).pdf'.format(i + 1)
            MCE_pdf_file_path = result_path + '_MCE({}).pdf'.format(i + 1)
            DE_pdf_file_path = DE_pdf_file_path.replace('/', '\\')
            MCE_pdf_file_path = MCE_pdf_file_path.replace('/', '\\')
            os.remove(DE_pdf_file_path)
            os.remove(MCE_pdf_file_path)
