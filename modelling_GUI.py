
import sys
import os
import time
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QSettings, QCoreApplication, QThread, QObject, pyqtSignal, Qt
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5 import uic # ui 파일을 사용하기 위한 모듈

import PBD_p3d as pbd

#%% Worker object
class worker1(QObject):               
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
        self.import_mass = args[10]
        self.import_nodal_load = args[11]
        self.time_start = args[12]
    
    # Import (MIDAS Gen -> Perform-3D) function    
    def import_midas_fn(self):  
        try:
            # 함수 실행
            pbd.import_midas(self.input_xlsx_path, DL_name=self.DL, LL_name=self.LL
                              , import_node=self.import_node, import_beam=self.import_beam
                              , import_column=self.import_col, import_wall=self.import_wall
                              , import_plate=self.import_plate, import_WR_gage=self.import_WR_gage
                              , import_WAS_gage=self.import_WAS_gage, import_mass=self.import_mass
                              , import_DL=self.import_nodal_load, import_LL=self.import_nodal_load)
            
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
class worker2(QObject):   
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
class worker3(QObject):   
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

#%% UI
ui_class = uic.loadUiType('PBD_p3d.ui')[0]

class main_window(QMainWindow, ui_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        
        # QPixmap 객체 생성 (image 불러오기)
        self.qPixmapVar = QPixmap()
        self.qPixmapVar.load('./images/p3d_logo.png')
        self.label_16.setPixmap(self.qPixmapVar.scaled(self.label_16.size()
                                , transformMode=Qt.SmoothTransformation))
        self.label_17.setPixmap(self.qPixmapVar.scaled(self.label_17.size()
                                , transformMode=Qt.SmoothTransformation))
        
        # setting에 저장된 value를 불러와서 입력
        QCoreApplication.setOrganizationName('CNP_Dongyang')
        QCoreApplication.setApplicationName('PBD_with_PERFORM-3D')
        self.setting = QSettings()

        # setting에 저장된 value를 불러와서 입력
        self.setting.beginGroup('file_path')
        self.data_conv_path_editbox.setText(self.setting.value('data_conversion_file_path', 'C:\\'))
        self.setting.endGroup()
        
        self.setting.beginGroup('setting_tab1')
        self.import_node_checkbox.setChecked(self.setting.value('import_node', True, type=bool))
        self.import_beam_checkbox.setChecked(self.setting.value('import_beam', True, type=bool))
        self.import_col_checkbox.setChecked(self.setting.value('import_column', False, type=bool))
        self.import_wall_checkbox.setChecked(self.setting.value('import_wall', True, type=bool))
        self.import_plate_checkbox.setChecked(self.setting.value('import_plate', False, type=bool))
        self.import_WR_gage_checkbox.setChecked(self.setting.value('import_wall_rotation_gage', True, type=bool))
        self.import_WAS_gage_checkbox.setChecked(self.setting.value('import_wall_axial_strain_gage', True, type=bool))
        self.import_mass_checkbox.setChecked(self.setting.value('import_mass', True, type=bool))
        self.import_nodal_load_checkbox.setChecked(self.setting.value('import_nodal_load', True, type=bool))
        self.DL_name_editbox.setText(self.setting.value('DL_name', 'DL'))
        self.LL_name_editbox.setText(self.setting.value('LL_name', 'LL'))
        self.drift_pos_editbox.setText(self.setting.value('drift_positions', '2,5,7,11'))
        self.convert_wall_checkbox.setChecked(self.setting.value('convert_wall', True, type=bool))
        self.convert_cbeam_checkbox.setChecked(self.setting.value('convert_beam', True, type=bool))
        self.convert_gcol_checkbox.setChecked(self.setting.value('convert_column', True, type=bool))
        self.setting.endGroup()

        # 버튼 누르기
        self.find_file_btn.clicked.connect(self.find_input_xlsx)
        self.import_midas_btn.clicked.connect(self.run_worker1)
        self.print_name_btn.clicked.connect(self.run_worker2)
        self.convert_prop_btn.clicked.connect(self.run_worker3)

         
        # 기타
        self.setWindowIcon(QIcon('./images/icon_earthquake.ico')) # icon 설정        

    # 파일 선택 function (QFileDialog)
    def find_input_xlsx(self):
        # global input_xlsx_path
        input_xlsx_path = QFileDialog.getOpenFileName(parent=self, caption='Open File'
                                    , directory=os.getcwd(), filter='Excel File (*.xlsx *.xls)')[0]
        
        self.data_conv_path_editbox.setText(input_xlsx_path)

    def run_worker1(self):
        # 시작 메세지
        time_start = time.time()
        self.status_browser.append('Running.....')
        
        # 변수 정리
        input_xlsx_path = self.data_conv_path_editbox.text()
        DL = self.DL_name_editbox.text()
        LL = self.LL_name_editbox.text()
        import_node = self.import_node_checkbox.isChecked()
        import_beam = self.import_beam_checkbox.isChecked()
        import_col = self.import_col_checkbox.isChecked()
        import_wall = self.import_wall_checkbox.isChecked()
        import_plate = self.import_plate_checkbox.isChecked()
        import_WR_gage = self.import_WR_gage_checkbox.isChecked()
        import_WAS_gage = self.import_WAS_gage_checkbox.isChecked()
        import_mass = self.import_mass_checkbox.isChecked()
        import_nodal_load = self.import_nodal_load_checkbox.isChecked()
        
        # self.thread.clear()
        self.thread = QThread(parent=self) # Create a QThread object
        self.worker = worker1(input_xlsx_path, DL, LL, import_node, import_beam
                              , import_col, import_wall, import_plate
                              , import_WR_gage, import_WAS_gage, import_mass
                              , import_nodal_load, time_start) # Create a worker object
        self.worker.moveToThread(self.thread) # Move worker to the thread
        
        # Connect signals and slots
        self.thread.started.connect(self.worker.import_midas_fn)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        
        # Start the thread
        self.thread.start()
        
        # Enable/Disable the Button
        self.import_midas_btn.setEnabled(False)
        self.print_name_btn.setEnabled(False)
        self.convert_prop_btn.setEnabled(False)
        self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
        self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
        self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
        
        # 완료 메세지 print
        self.worker.msg.connect(self.msg_fn)

    def run_worker2(self):
        # 시작 메세지
        time_start = time.time()
        self.status_browser.append('Running.....')
        
        # 변수 정리
        input_xlsx_path = self.data_conv_path_editbox.text()
        drift_pos_raw = self.drift_pos_editbox.text()
        drift_position = []
        for i in range(drift_pos_raw.count(',')+1):
            drift_pos_elem = drift_pos_raw.split(',')[i].strip()
            drift_position.append(drift_pos_elem)
        
        # self.thread.clear()
        self.thread = QThread(parent=self) # Create a QThread object
        self.worker = worker2(input_xlsx_path, drift_position, time_start) # Create a worker object
        self.worker.moveToThread(self.thread) # Move worker to the thread
        
        # Connect signals and slots
        self.thread.started.connect(self.worker.naming_fn)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        
        # Start the thread
        self.thread.start()
        
        # Enable/Disable the Button
        self.import_midas_btn.setEnabled(False)
        self.print_name_btn.setEnabled(False)
        self.convert_prop_btn.setEnabled(False)
        self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
        self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
        self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
        
        # 완료 메세지 print
        self.worker.msg.connect(self.msg_fn)
        
    def run_worker3(self):
        # 시작 메세지
        time_start = time.time()
        self.status_browser.append('Running.....')
        
        # 변수 정리
        input_xlsx_path = self.data_conv_path_editbox.text()
        get_beam = self.convert_cbeam_checkbox.isChecked()
        get_column = self.convert_gcol_checkbox.isChecked()
        get_wall = self.convert_wall_checkbox.isChecked()
        
        # self.thread.clear()
        self.thread = QThread(parent=self) # Create a QThread object
        self.worker = worker3(input_xlsx_path, get_beam, get_column, get_wall, time_start) # Create a worker object
        self.worker.moveToThread(self.thread) # Move worker to the thread
        
        # Connect signals and slots
        self.thread.started.connect(self.worker.convert_property_fn)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        
        # Start the thread
        self.thread.start()
        
        # Enable/Disable the Button
        self.import_midas_btn.setEnabled(False)
        self.print_name_btn.setEnabled(False)
        self.convert_prop_btn.setEnabled(False)
        self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
        self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
        self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
        
        # 완료 메세지 print
        self.worker.msg.connect(self.msg_fn)

    # 로그(실행 시간, 오류) print function
    def msg_fn(self, msg):
        self.status_browser.append(msg)

    # 실행해야할 명령 전달 when Qt receives a window close request        
    def closeEvent(self, event):            
        if self.rmb_setting_checkbox.isChecked():
            # setValue(key, value) : key와 value를 setting에 저장
            # .text() : 해당 box의 값 / .currentText() : 해당 combobox의 값
            self.setting.beginGroup('file_path')
            self.setting.setValue('data_conversion_file_path', self.data_conv_path_editbox.text())
            self.setting.endGroup()
            
            self.setting.beginGroup('setting_tab1')
            self.setting.setValue('import_node', self.import_node_checkbox.isChecked())
            self.setting.setValue('import_beam', self.import_beam_checkbox.isChecked())
            self.setting.setValue('import_column', self.import_col_checkbox.isChecked())
            self.setting.setValue('import_wall', self.import_wall_checkbox.isChecked())
            self.setting.setValue('import_plate', self.import_plate_checkbox.isChecked())
            self.setting.setValue('import_wall_rotation_gage', self.import_WR_gage_checkbox.isChecked())
            self.setting.setValue('import_wall_axial_strain_gage', self.import_WAS_gage_checkbox.isChecked())
            self.setting.setValue('import_mass', self.import_mass_checkbox.isChecked())
            self.setting.setValue('import_nodal_load', self.import_nodal_load_checkbox.isChecked())
            self.setting.setValue('DL_name', self.DL_name_editbox.text())
            self.setting.setValue('LL_name', self.LL_name_editbox.text())
            self.setting.setValue('drift_positions', self.drift_pos_editbox.text())
            self.setting.setValue('convert_wall', self.convert_wall_checkbox.isChecked())
            self.setting.setValue('convert_beam', self.convert_cbeam_checkbox.isChecked())
            self.setting.setValue('convert_column', self.convert_gcol_checkbox.isChecked())
            self.setting.endGroup()
        else: self.setting.clear()

# =============================================================================
#         # Properties Assign 매크로 Option
#         drag_duration = 0.15 # drag 하는 속도(너무 빨리하면 팅길 수 있으므로 적당한 속도 권장)
#         offset = 2 # 픽셀 오차 방지용 여유치, 단위 : pixel
#         start_material_index = 761 # wall_material_data_repeat 에서 시작하고자 하는 material name의 index 입력, 처음부터일때는 0 입력
# =============================================================================

#%% Properties Assign 매크로

# pbd.property_assign_macro(drag_duration=0.15, offset=2, start_material_index=761)

#%% 실행

if __name__ == '__main__':
    app = QApplication(sys.argv) # QApplication : 프로그램을 실행시켜주는 class
    mywindow = main_window() # WindowClass의 인스턴스 생성   
    mywindow.show() # 프로그램 보여주기
    app.exec_() # 프로그램을 작동시키는 코드
