import time
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
            self.msg.emit('%s' %e)
        
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
            self.msg.emit('%s' %e)
        
# Convert Properties Worker 만들기   
class ConvertWorker(QObject):   
    # Create signals
    finished = pyqtSignal()
    msg = pyqtSignal(str)   
    def __init__(self, *args):
        super().__init__()

        self.input_xlsx_path = args[0]
        self.get_beam = args[1]
        self.get_column = args[2]
        self.get_wall = args[3]
        self.time_start = args[4]

    # Properties 변환 function
    def convert_property_fn(self):   
        try:
            # 함수 실행
            pbd.convert_property(self.input_xlsx_path, get_beam=self.get_beam
                                 , get_column=self.get_column, get_wall=self.get_wall)      
            # 실행 시간 계산
            time_end = time.time()
            time_run = (time_end-self.time_start)/60            
            # Emit
            self.finished.emit()
            self.msg.emit('Completed!' + '  (total time = %0.3f min)' %(time_run))
            
        except Exception as e:
            self.finished.emit()
            self.msg.emit('%s' %e)


            
# Print Results Worker 만들기
class ExportWorker(QObject):               
    # Create signals
    finished = pyqtSignal()
    msg = pyqtSignal(str)    
    def __init__(self, *args):
        super().__init__()
        
        # 변수 정리
        self.time_start = args[16]
    
    # Properties 변환 function
    def initialize_fn(self):   
        try:
            # 시작 메세지
            time_start = time.time()
            self.status_browser.append('Running.....')

            # Disable the Button
            self.show_result_btn.setEnabled(False)
            self.print_result_btn.setEnabled(False)    
            
            # Emit
            self.finished.emit()
            
        except Exception as e:
            self.finished.emit()
