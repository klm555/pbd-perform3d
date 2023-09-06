import time
import pickle
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QObject, pyqtSignal, Qt

import PBD_p3d as pbd

#%% Worker object
# 전처리=======================================================================
# Import Worker 만들기
class ImportWorker(QObject):               
    # Create signals
    finished = pyqtSignal()
    msg = pyqtSignal(str)
    
    def __init__(self, *args):
        super().__init__()
        # 변수 정리
        self.input_xlsx_path = args[0]
        self.DL = args[1]
        self.LL = args[2]
        self.import_node = args[3]
        self.import_beam = args[4]
        self.import_col = args[5]
        self.import_wall = args[6]
        self.import_plate = args[7]
        self.import_WR_gage = args[8]
        self.import_WAS_gage = args[9]
        self.import_I_beam = args[10]
        self.import_mass = args[11]
        self.import_nodal_load = args[12]
        self.time_start = args[13]
    
    # Import (MIDAS Gen -> Perform-3D) function    
    def import_midas_fn(self):  
        try:
            # 함수 실행
            pbd.import_midas(self.input_xlsx_path, DL_name=self.DL, LL_name=self.LL
                              , import_node=self.import_node, import_beam=self.import_beam
                              , import_column=self.import_col, import_wall=self.import_wall
                              , import_plate=self.import_plate, import_WR_gage=self.import_WR_gage
                              , import_WAS_gage=self.import_WAS_gage, import_I_beam=self.import_I_beam
                              , import_mass=self.import_mass, import_DL=self.import_nodal_load
                              , import_LL=self.import_nodal_load)
            
            # 실행 시간 계산
            time_end = time.time()
            time_run = (time_end-self.time_start)/60
            
            # Emit
            self.finished.emit()
            self.msg.emit('Completed!' + '  (total time = %0.5f min)' %(time_run))
            
        except Exception as e:
            self.finished.emit()
            self.msg.emit('Error : %s' %e)
        
# Print Name Worker 만들기   
class NameWorker(QObject):   
    # Create signals
    finished = pyqtSignal()
    msg = pyqtSignal(str)
    
    def __init__(self, *args):
        super().__init__()

        self.input_xlsx_path = args[0]
        self.drift_position = args[1]
        self.time_start = args[2]
        
    # Naming function
    def naming_fn(self):
        try:
            # 함수 실행
            pbd.naming(self.input_xlsx_path, drift_position=self.drift_position)            
            # 실행 시간 계산
            time_end = time.time()
            time_run = (time_end-self.time_start)/60            
            # Emit
            self.finished.emit()
            self.msg.emit('Completed!' + '  (total time = %0.5f min)' %(time_run))
        
        except Exception as e:
            self.finished.emit()
            self.msg.emit('Error : %s' %e)
        
# Convert Properties Worker 만들기   
class ConvertWorker(QObject):   
    # Create signals
    finished = pyqtSignal()
    msg = pyqtSignal(str)   
    def __init__(self, *args):
        super().__init__()

        self.input_xlsx_path = args[0]
        self.get_cbeam = args[1]
        self.get_gbeam = args[2]
        self.get_ebeam = args[3]
        self.get_gcol = args[4]
        self.get_ecol = args[5]
        self.get_wall = args[6]
        self.time_start = args[7]

    # Properties 변환 function
    def convert_property_fn(self):   
        try:
            # 함수 실행
            pbd.convert_property(self.input_xlsx_path, get_cbeam=self.get_cbeam
                                 , get_gbeam=self.get_gbeam, get_ebeam=self.get_ebeam
                                 , get_gcol=self.get_gcol, get_ecol=self.get_ecol
                                 , get_wall=self.get_wall)      
            # 실행 시간 계산
            time_end = time.time()
            time_run = (time_end-self.time_start)/60            
            # Emit
            self.finished.emit()
            self.msg.emit('Completed!' + '  (total time = %0.3f min)' %(time_run))
            
        except Exception as e:
            self.finished.emit()
            self.msg.emit('Error : %s' %e)

# Insert Properties Worker 만들기   
class InsertWorker(QObject):   
    # Create signals
    finished = pyqtSignal()
    msg = pyqtSignal(str)   
    def __init__(self, *args):
        super().__init__()

        self.input_xlsx_path = args[0]
        self.result_xlsx_path = args[1]
        self.get_gbeam = args[2]
        self.get_gcol = args[3]
        self.get_ecol = args[4]
        self.time_start = args[5]

    # Properties 변환 function
    def insert_force_fn(self):   
        try:
            # 함수 실행
            pbd.insert_force(self.input_xlsx_path, self.result_xlsx_path
                             , get_gbeam=self.get_gbeam, get_gcol=self.get_gcol
                             , get_ecol=self.get_ecol)      
            # 실행 시간 계산
            time_end = time.time()
            time_run = (time_end-self.time_start)/60            
            # Emit
            self.finished.emit()
            self.msg.emit('Completed!' + '  (total time = %0.3f min)' %(time_run))
            
        except Exception as e:
            self.finished.emit()
            self.msg.emit('Error : %s' %e)

            
# Load Results Worker 만들기
class LoadWorker(QObject):               
    # Create signals
    finished = pyqtSignal()
    msg = pyqtSignal(str)
    result_data = pyqtSignal(object)
    def __init__(self, *args):
        super().__init__()
        
        # 변수 정리
        self.input_xlsx_path = args[0]
        self.result_xlsx_path = args[1]
        self.wall_design_xlsx_path = args[2]
        self.beam_design_xlsx_path = args[3]
        self.col_design_xlsx_path = args[4]
        self.get_base_SF = args[5]
        self.get_story_SF = args[6]
        self.get_IDR = args[7]
        self.get_BR = args[8]
        self.get_BSF = args[9]
        self.get_E_BSF = args[10]
        self.get_CR = args[11]
        self.get_CSF = args[12]
        self.get_E_CSF = args[13]
        self.get_WAS = args[14]
        self.get_WR = args[15]
        self.get_WSF = args[16]
        self.story_gap = args[17]
        self.max_shear = args[18]
        self.BR_scale_factor = args[19]
        self.time_start = args[20]
    
    # Properties 변환 function
    def load_result_fn(self):   
        try:
            # 함수 실행
            result = pbd.PostProc(self.input_xlsx_path, self.result_xlsx_path
                                  , self.get_base_SF, self.get_story_SF
                                  , self.get_IDR, self.get_BR, self.get_BSF
                                  , self.get_E_BSF, self.get_CR, self.get_CSF
                                  , self.get_E_CSF, self.get_WAS, self.get_WR
                                  , self.get_WSF, self.BR_scale_factor)
            
            # 결과 데이터를 pickle로 출력&저장
            result_dict = {}
            if self.get_base_SF == True:
                result.base_SF(self.max_shear)  
                # pickle 파일 읽기
                with open('pkl/base_SF.pkl', 'rb') as f:
                    result_dict['base_SF'] = pickle.load(f)
            if self.get_story_SF == True:
                result.story_SF(self.story_gap, self.max_shear)
                # pickle 파일 읽기
                with open('pkl/story_SF.pkl', 'rb') as f:
                    result_dict['story_SF'] = pickle.load(f)
            if self.get_IDR == True:
                result.IDR(yticks=self.story_gap)
                # pickle 파일 읽기
                with open('pkl/IDR.pkl', 'rb') as f:
                    result_dict['IDR'] = pickle.load(f)
            if self.get_BR == True:
                result.BR(self.input_xlsx_path, self.beam_design_xlsx_path
                          , yticks=self.story_gap, scale_factor=self.BR_scale_factor)
                # pickle 파일 읽기
                with open('pkl/BR.pkl', 'rb') as f:
                    result_dict['BR'] = pickle.load(f)
            if self.get_BSF == True:
                result.BSF(self.input_xlsx_path, self.beam_design_xlsx_path, yticks=self.story_gap)
                # pickle 파일 읽기
                with open('pkl/BSF.pkl', 'rb') as f:
                    result_dict['BSF'] = pickle.load(f)
            if self.get_WAS == True:
                result.WAS(self.wall_design_xlsx_path, yticks=self.story_gap)
                # pickle 파일 읽기
                with open('pkl/WAS.pkl', 'rb') as f:
                    result_dict['WAS'] = pickle.load(f)                    
            if self.get_WR == True:
                result.WR(self.input_xlsx_path, self.wall_design_xlsx_path, yticks=self.story_gap)
                # pickle 파일 읽기
                with open('pkl/WR.pkl', 'rb') as f:
                    result_dict['WR'] = pickle.load(f)
            if self.get_WSF == True:
                result.WSF(self.input_xlsx_path, self.wall_design_xlsx_path, yticks=self.story_gap)
                # pickle 파일 읽기
                with open('pkl/WSF.pkl', 'rb') as f:
                    result_dict['WSF'] = pickle.load(f)
            if self.get_E_CSF == True:
                result.E_CSF(self.input_xlsx_path, self.col_design_xlsx_path, yticks=self.story_gap)
                # pickle 파일 읽기
                # with open('pkl/E_CSF.pkl', 'rb') as f:
                #     result_dict['E_CSF'] = pickle.load(f)            
            
            # Result pickle과 time_start를 묶어서 결과로 내보냄
            result_dict_and_time = [result_dict, self.time_start]
            
            # 데이터 emit
            self.result_data.emit(result_dict_and_time)
            # 종료여부 emit
            self.finished.emit()
            
        except Exception as e:
            self.finished.emit()
            self.msg.emit('Error : %s' %e)
            
# Design Wall Worker 만들기
class RedesignWorker(QObject):               
    # Create signals
    finished = pyqtSignal()
    msg = pyqtSignal(str)
    def __init__(self, *args):
        super().__init__()
        
        # 변수 정리
        self.wall_design_xlsx_path = args[0]
        self.time_start = args[1]
    
    # 벽체 수평배근 function
    def redesign_wall_fn(self):   
        try:
            # 함수 실행
            pbd.WSF_redesign(self.wall_design_xlsx_path, rebar_limit=[None,None])      
            # 실행 시간 계산
            time_end = time.time()
            time_run = (time_end-self.time_start)/60            
            # Emit
            self.finished.emit()
            self.msg.emit('Completed!' + '  (total time = %0.3f min)' %(time_run))
            
        except Exception as e:
            self.finished.emit()
            self.msg.emit('Error : %s' %e)
            
# Print pdf Worker 만들기
class PdfWorker(QObject):               
    # Create signals
    finished = pyqtSignal()
    msg = pyqtSignal(str)
    def __init__(self, *args):
        super().__init__()
        
        # 변수 정리
        self.beam_design_xlsx_path = args[0]
        self.col_design_xlsx_path = args[1]
        self.wall_design_xlsx_path = args[2]
        self.get_cbeam = args[3]
        self.get_ecol = args[4]
        self.get_wall = args[5]
        self.project_name = args[6]
        self.bldg_name = args[7]
        self.time_start = args[8]
    
    # 벽체 수평배근 function
    def print_pdf_fn(self):   
        try:
            # 함수 실행
            pbd.print_pdf(self.beam_design_xlsx_path, self.col_design_xlsx_path
                          , self.wall_design_xlsx_path, self.get_cbeam, self.get_ecol
                          , self.get_wall, self.project_name, self.bldg_name)
            # 실행 시간 계산
            time_end = time.time()
            time_run = (time_end-self.time_start)/60            
            # Emit
            self.finished.emit()
            self.msg.emit('Completed!' + '  (total time = %0.3f min)' %(time_run))
            
        except Exception as e:
            self.finished.emit()
            self.msg.emit('Error : %s' %e)
            
# Print docx Worker 만들기
class DocxWorker(QObject):               
    # Create signals
    finished = pyqtSignal()
    msg = pyqtSignal(str)
    def __init__(self, *args):
        super().__init__()
        
        # 변수 정리
        self.result_xlsx_path = args[0]
        self.story_gap = args[1]
        self.max_shear = args[2]
        self.get_base_SF = args[3]
        self.get_story_SF = args[4]
        self.get_IDR = args[5]
        self.get_BR = args[6]
        self.get_BSF = args[7]
        self.get_E_BSF = args[8]
        self.get_CR = args[9]
        self.get_CSF = args[10]
        self.get_E_CSF = args[11]
        self.get_WAS = args[12]
        self.get_WR = args[13]
        self.get_WSF = args[14]
        self.project_name = args[15]
        self.bldg_name = args[16]
        self.story_gap = args[17]
        self.max_shear = args[18]
        self.time_start = args[19]
        
    # 벽체 수평배근 function
    def print_docx_fn(self):   
        try:
            # 함수 실행
            pbd.print_docx(self.result_xlsx_path, self.get_base_SF
                           , self.get_story_SF, self.get_IDR, self.get_BR
                           , self.get_BSF, self.get_E_BSF, self.get_CR
                           , self.get_CSF, self.get_E_CSF, self.get_WAS
                           , self.get_WR, self.get_WSF, self.project_name
                           , self.bldg_name, self.story_gap, self.max_shear)      
            # 실행 시간 계산
            time_end = time.time()
            time_run = (time_end-self.time_start)/60            
            # Emit
            self.finished.emit()
            self.msg.emit('Completed!' + '  (total time = %0.3f min)' %(time_run))
            
        except Exception as e:
            self.finished.emit()
            self.msg.emit('Error : %s' %e)