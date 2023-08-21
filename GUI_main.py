import sys
import os
import shutil
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QSettings, QCoreApplication, QThread, QObject, Qt, pyqtSlot
from PyQt5.QtGui import QIcon, QPixmap

from PyQt5 import uic # ui 파일을 사용하기 위한 모듈
import multiprocess as mp

#%% UI
ui_class = uic.loadUiType('PBD_p3d.ui')[0]

class MainWindow(QMainWindow, ui_class):

    # Import external classes/functions
    from GUI_workers import ImportWorker, NameWorker, ConvertWorker, InsertWorker, LoadWorker, RedesignWorker, PdfWorker
    from GUI_runs import run_worker1, run_worker2, run_worker3, run_worker4, run_worker5, run_worker7, run_worker8
    from GUI_plots import plot_display

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        
        ##### QPixmap 객체 생성 (image 불러오기)
        # P3D 로고
        self.qPixmapVar = QPixmap()
        self.qPixmapVar.load('./images/p3d_logo.png')
        self.P3D_img.setPixmap(self.qPixmapVar.scaled(self.P3D_img.size()
                                , transformMode=Qt.SmoothTransformation))
        # CNP 동양 로고
        self.qPixmapVar2 = QPixmap()
        self.qPixmapVar2.load('./images/CNP_logo.png')
        self.CNP_img.setPixmap(self.qPixmapVar2.scaled(self.CNP_img.size()
                                , transformMode=Qt.SmoothTransformation))
        
        ##### setting에 저장된 value를 불러와서 입력
        # QSettings 클래스 생성
        QCoreApplication.setOrganizationName('CNP_Dongyang')
        QCoreApplication.setApplicationName('PBD_with_PERFORM-3D')
        self.setting = QSettings()

        # setting에 저장된 value를 불러와서 입력
        self.setting.beginGroup('file_path')
        self.data_conv_path_editbox.setText(self.setting.value('data_conversion_file_path', 'C:\\'))
        self.result_path_editbox.setText(self.setting.value('result_file_path', 'C:\\'))
        result_xlsx_path = self.result_path_editbox.text().split('"')
        result_xlsx_path = [i for i in result_xlsx_path if len(i) > 4]
        self.display_selected_result_path.setText('%i files selected' %len(result_xlsx_path))
        self.wall_design_path_editbox.setText(self.setting.value('wall_design_file_path', 'C:\\'))
        self.beam_design_path_editbox.setText(self.setting.value('beam_design_file_path', 'C:\\'))
        self.col_design_path_editbox.setText(self.setting.value('column_design_file_path', 'C:\\'))
        self.setting.endGroup()
        
        # Tab 1
        self.setting.beginGroup('setting_tab1')
        self.import_node_checkbox.setChecked(self.setting.value('import_node', True, type=bool))
        self.import_beam_checkbox.setChecked(self.setting.value('import_beam', True, type=bool))
        self.import_col_checkbox.setChecked(self.setting.value('import_column', False, type=bool))
        self.import_wall_checkbox.setChecked(self.setting.value('import_wall', True, type=bool))
        self.import_plate_checkbox.setChecked(self.setting.value('import_plate', False, type=bool))
        self.import_WR_gage_checkbox.setChecked(self.setting.value('import_wall_rotation_gage', True, type=bool))
        self.import_WAS_gage_checkbox.setChecked(self.setting.value('import_wall_axial_strain_gage', True, type=bool))
        self.import_I_beam_checkbox.setChecked(self.setting.value('import_I_beam', True, type=bool))
        self.import_mass_checkbox.setChecked(self.setting.value('import_mass', True, type=bool))
        self.import_nodal_load_checkbox.setChecked(self.setting.value('import_nodal_load', True, type=bool))
        self.DL_name_editbox.setText(self.setting.value('DL_name', 'DL'))
        self.LL_name_editbox.setText(self.setting.value('LL_name', 'LL'))
        self.drift_pos_editbox.setText(self.setting.value('drift_positions', '2,5,7,11'))
        self.convert_wall_checkbox.setChecked(self.setting.value('convert_wall', True, type=bool))
        self.convert_cbeam_checkbox.setChecked(self.setting.value('convert_cbeam', True, type=bool))
        self.convert_gbeam_checkbox.setChecked(self.setting.value('convert_gbeam', False, type=bool))
        self.convert_ebeam_checkbox.setChecked(self.setting.value('convert_ebeam', False, type=bool))
        self.convert_gcol_checkbox.setChecked(self.setting.value('convert_gcolumn', False, type=bool))
        self.convert_ecol_checkbox.setChecked(self.setting.value('convert_ecolumn', False, type=bool))
        self.setting.endGroup()        
        # Tab 2
        self.setting.beginGroup('setting_tab2')
        self.insert_gbeam_checkbox.setChecked(self.setting.value('insert_gbeam', False, type=bool))
        self.insert_gcol_checkbox.setChecked(self.setting.value('insert_gcolumn', False, type=bool))
        self.insert_ecol_checkbox.setChecked(self.setting.value('insert_ecolumn', False, type=bool))
        self.setting.endGroup()
        # self.drag_duration_slider.setValue(2)        
        # Tab 3
        self.setting.beginGroup('setting_tab3')
        self.base_SF_checkbox.setChecked(self.setting.value('base_SF', True, type=bool))
        self.story_SF_checkbox.setChecked(self.setting.value('story_SF', False, type=bool))
        self.IDR_checkbox.setChecked(self.setting.value('IDR', True, type=bool))
        self.BR_checkbox.setChecked(self.setting.value('BR', True, type=bool))
        self.BSF_checkbox.setChecked(self.setting.value('BSF', True, type=bool))
        self.E_BSF_checkbox.setChecked(self.setting.value('E_BSF', False, type=bool))
        self.CR_checkbox.setChecked(self.setting.value('CR', False, type=bool))
        self.CSF_checkbox.setChecked(self.setting.value('CSF', False, type=bool))
        self.E_CSF_checkbox.setChecked(self.setting.value('E_CSF', False, type=bool))
        self.WAS_checkbox.setChecked(self.setting.value('WAS', True, type=bool))
        self.WR_checkbox.setChecked(self.setting.value('WR', True, type=bool))
        self.WSF_checkbox.setChecked(self.setting.value('WSF', True, type=bool))
        self.cbeam_pdf_checkbox.setChecked(self.setting.value('cbeam_to_pdf', True, type=bool))
        self.ecol_pdf_checkbox.setChecked(self.setting.value('ecolumn_to_pdf', True, type=bool))
        self.wall_pdf_checkbox.setChecked(self.setting.value('wall_to_pdf', True, type=bool))
        self.bldg_name_editbox.setText(self.setting.value('bldg_name', '101동'))
        self.story_gap_editbox.setText(self.setting.value('story_gap', '2'))
        self.max_shear_editbox.setText(self.setting.value('max_shear', '60000'))
        self.output_docx_editbox.setText(self.setting.value('output_docx', '해석 결과'))
        self.setting.endGroup()

        ##### Connect Button
        # Load File
        self.find_file_btn.clicked.connect(self.find_input_xlsx)
        self.find_file_btn_2.clicked.connect(self.find_result_xlsx)
        self.find_file_btn_3.clicked.connect(self.find_wall_design_xlsx)
        self.find_file_btn_4.clicked.connect(self.find_beam_design_xlsx)
        self.find_file_btn_5.clicked.connect(self.find_col_design_xlsx)
        
        # Tab 1
        self.import_midas_btn.clicked.connect(self.run_worker1)
        self.print_name_btn.clicked.connect(self.run_worker2)
        self.convert_prop_btn.clicked.connect(self.run_worker3)    
        # Tab 2
        self.insert_force_btn.clicked.connect(self.run_worker4)
        # Tab 3
        self.load_result_btn.clicked.connect(self.run_worker5)
        # self.load_result_btn.clicked.connect(self.plot_display)
        # self.print_result_btn.clicked.connect(self.run_worker6)
        self.design_wall_btn.clicked.connect(self.run_worker7)
        self.print_pdf_btn.clicked.connect(self.run_worker8)
         
        # Icon 설정
        self.setWindowIcon(QIcon('./images/icon_earthquake.ico'))
        
        # 마우스 좌표 real-time tracking
        self.setMouseTracking(True)

    # 파일 선택 function (QFileDialog)
    def find_input_xlsx(self): # Data Conversion Sheets
        # global input_xlsx_path
        input_xlsx_path = QFileDialog.getOpenFileName(parent=self, caption='Open File'
                                    , directory=os.getcwd(), filter='Excel File (*.xlsx *.xls)')[0]
        
        self.data_conv_path_editbox.setText(input_xlsx_path)
        
    def find_result_xlsx(self): # Analysis Resutls
        # global input_xlsx_path
        result_xlsx_path = QFileDialog.getOpenFileNames(parent=self, caption='Open Folder'
                                    , directory=os.getcwd(), filter='Excel File (*.xlsx *.xls)')[0]
        
        all_result_xlsx_path = ['"%s"' %file_name for file_name in result_xlsx_path]
        joined_result_xlsx_path = ','.join(all_result_xlsx_path)
        self.result_path_editbox.setText(joined_result_xlsx_path)
        self.display_selected_result_path.setText('%i files selected' %len(result_xlsx_path))
        
    def find_wall_design_xlsx(self): # Wall Results Sheets
        # global input_xlsx_path
        wall_design_xlsx_path = QFileDialog.getOpenFileName(parent=self, caption='Open File'
                                    , directory=os.getcwd(), filter='Excel File (*.xlsx *.xls)')[0]
        
        self.wall_design_path_editbox.setText(wall_design_xlsx_path)
        
    def find_beam_design_xlsx(self): # Beam Results Sheets
        # global input_xlsx_path
        beam_design_xlsx_path = QFileDialog.getOpenFileName(parent=self, caption='Open File'
                                    , directory=os.getcwd(), filter='Excel File (*.xlsx *.xls)')[0]
        
        self.beam_design_path_editbox.setText(beam_design_xlsx_path)
    
    def find_col_design_xlsx(self): # Column Results Sheets
        # global input_xlsx_path
        col_design_xlsx_path = QFileDialog.getOpenFileName(parent=self, caption='Open File'
                                    , directory=os.getcwd(), filter='Excel File (*.xlsx *.xls)')[0]
        
        self.col_design_path_editbox.setText(col_design_xlsx_path)
    
    
    
    # Macro 사용을 위한 function
    def mouseMoveEvent(self, event):
        self.up_left_coord_editbox.setText('( %d : %d )' %(event.globalX(), event.globalY()))
        
    def threshold1(self, value):
        self.drag_duration_slider = value
        self.test.setText("Value : " + str(value))

    def threshold2(self, value):
        self.offset_slider = value

    # 로그(실행 시간, 오류) print function
    def msg_fn(self, msg):
        if 'Completed' in msg: # pyqt signal에 'Completed' 들어있는 경우, 파란색으로 표시
            msg_colored = '<span style=\" color: #0000ff;\">%s</span>' % msg
        else: # 에러 생길 경우, 빨간색으로 표시
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
        self.status_browser.append(msg_colored)

    # 실행해야할 명령 전달 when Qt receives a window close request        
    def closeEvent(self, event):            
        if self.rmb_setting_checkbox.isChecked():
            # setValue(key, value) : key와 value를 setting에 저장
            # .text() : 해당 box의 값 / .currentText() : 해당 combobox의 값
            self.setting.beginGroup('file_path')
            self.setting.setValue('data_conversion_file_path', self.data_conv_path_editbox.text())
            self.setting.setValue('result_file_path', self.result_path_editbox.text())
            self.setting.setValue('wall_design_file_path', self.wall_design_path_editbox.text())
            self.setting.setValue('beam_design_file_path', self.beam_design_path_editbox.text())
            self.setting.setValue('column_design_file_path', self.col_design_path_editbox.text())
            self.setting.endGroup()
            
            self.setting.beginGroup('setting_tab1')
            self.setting.setValue('import_node', self.import_node_checkbox.isChecked())
            self.setting.setValue('import_beam', self.import_beam_checkbox.isChecked())
            self.setting.setValue('import_column', self.import_col_checkbox.isChecked())
            self.setting.setValue('import_wall', self.import_wall_checkbox.isChecked())
            self.setting.setValue('import_plate', self.import_plate_checkbox.isChecked())
            self.setting.setValue('import_wall_rotation_gage', self.import_WR_gage_checkbox.isChecked())
            self.setting.setValue('import_wall_axial_strain_gage', self.import_WAS_gage_checkbox.isChecked())
            self.setting.setValue('import_I_beam', self.import_I_beam_checkbox.isChecked())
            self.setting.setValue('import_mass', self.import_mass_checkbox.isChecked())
            self.setting.setValue('import_nodal_load', self.import_nodal_load_checkbox.isChecked())
            self.setting.setValue('DL_name', self.DL_name_editbox.text())
            self.setting.setValue('LL_name', self.LL_name_editbox.text())
            self.setting.setValue('drift_positions', self.drift_pos_editbox.text())
            self.setting.setValue('convert_wall', self.convert_wall_checkbox.isChecked())
            self.setting.setValue('convert_cbeam', self.convert_cbeam_checkbox.isChecked())
            self.setting.setValue('convert_gbeam', self.convert_gbeam_checkbox.isChecked())
            self.setting.setValue('convert_ebeam', self.convert_ebeam_checkbox.isChecked())
            self.setting.setValue('convert_gcolumn', self.convert_gcol_checkbox.isChecked())
            self.setting.setValue('convert_ecolumn', self.convert_ecol_checkbox.isChecked())
            self.setting.endGroup()

            self.setting.beginGroup('setting_tab2')
            self.setting.setValue('insert_gbeam', self.insert_gbeam_checkbox.isChecked())
            self.setting.setValue('insert_gcolumn', self.insert_gcol_checkbox.isChecked())
            self.setting.setValue('insert_ecolumn', self.insert_ecol_checkbox.isChecked())
            self.setting.endGroup()

            self.setting.beginGroup('setting_tab3')
            self.setting.setValue('base_SF', self.base_SF_checkbox.isChecked())
            self.setting.setValue('story_SF', self.story_SF_checkbox.isChecked())
            self.setting.setValue('IDR', self.IDR_checkbox.isChecked())
            self.setting.setValue('BR', self.BR_checkbox.isChecked())
            self.setting.setValue('BSF', self.BSF_checkbox.isChecked())
            self.setting.setValue('E_BSF', self.E_BSF_checkbox.isChecked())
            self.setting.setValue('CR', self.CR_checkbox.isChecked())
            self.setting.setValue('CSF', self.CSF_checkbox.isChecked())
            self.setting.setValue('E_CSF', self.E_CSF_checkbox.isChecked())
            self.setting.setValue('WAS', self.WAS_checkbox.isChecked())
            self.setting.setValue('WR', self.WR_checkbox.isChecked())
            self.setting.setValue('WSF', self.WSF_checkbox.isChecked())
            self.setting.setValue('cbeam_to_pdf', self.cbeam_pdf_checkbox.isChecked())
            self.setting.setValue('ecolumn_to_pdf', self.ecol_pdf_checkbox.isChecked())
            self.setting.setValue('wall_to_pdf', self.wall_pdf_checkbox.isChecked())
            self.setting.setValue('bldg_name', self.bldg_name_editbox.text())
            self.setting.setValue('story_gap', self.story_gap_editbox.text())
            self.setting.setValue('max_shear', self.max_shear_editbox.text())
            self.setting.setValue('output_docx', self.output_docx_editbox.text())
            self.setting.endGroup()

        else: self.setting.clear()
        
        # pkl 폴더 삭제
        def delete_dir(directory):
            try:
                if os.path.exists(directory):
                    shutil.rmtree(directory)
            except OSError:
                print("Error: Failed to delete the directory.")
        delete_dir('pkl')

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
    mp.freeze_support() # multiprocessing fix
    app = QApplication(sys.argv) # QApplication : 프로그램을 실행시켜주는 class
    mywindow = MainWindow() # WindowClass의 인스턴스 생성   
    mywindow.show() # 프로그램 보여주기
    app.exec_() # 프로그램을 작동시키는 코드