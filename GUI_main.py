import sys
import os
import shutil
import multiprocess as mp
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QSettings, QCoreApplication, Qt
from PyQt5.QtGui import QIcon, QPixmap, QFontDatabase, QFont, QKeySequence
from PyQt5 import uic # ui 파일을 사용하기 위한 모듈

import pyautogui as pag
import numpy as np

from GUI_second import *

#%% Main 윈도우 UI
# .ui 파일(designer에서 만든 파일) 불러오기
ui_class = uic.loadUiType('PBD_p3d.ui')[0]

class MainWindow(QMainWindow, ui_class):
    # Import external classes/functions 
    # 여기서 import해야 / 사용되는 함수를 일일이 import해야 오류가 안남 (이유는 모름)
    from GUI_runs import run_worker1, run_worker2, run_worker3, run_worker4, run_worker5, run_worker6, run_worker7, run_worker8, run_worker9, run_worker10, run_worker11
    from GUI_plots import plot_display
    from GUI_workers import ImportWorker, NameWorker, ConvertWorker, InsertWorker, LoadWorker, RedesignWorker, PdfWorker, DocxWorker, HwpWorker
    
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        
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
        self.up_left_coord_editbox.setText(self.setting.value('upleft_coord', ' 0 , 0'))
        self.up_right_coord_editbox.setText(self.setting.value('upright_coord', ' 0 , 0'))
        self.low_left_coord_editbox.setText(self.setting.value('lowleft_coord', ' 0 , 0'))
        self.low_right_coord_editbox.setText(self.setting.value('lowright_coord', ' 0 , 0'))
        self.p3d_bar_editbox.setText(self.setting.value('p3d_bar_coord', ' 0 , 0'))
        self.add_cuts_editbox.setText(self.setting.value('add_cuts_coord', ' 0 , 0'))
        self.delete_cuts_editbox.setText(self.setting.value('delete_cuts_coord', ' 0 , 0'))
        self.ok_editbox.setText(self.setting.value('ok_coord', ' 0 , 0'))
        self.next_section_editbox.setText(self.setting.value('next_section_coord', ' 0 , 0'))
        self.next_frame_editbox.setText(self.setting.value('next_frame_coord', ' 0 , 0'))
        self.ok_delete_editbox.setText(self.setting.value('ok_delete_coord', ' 0 , 0'))
        self.missing_data_editbox.setText(self.setting.value('missing_data_coord', ' 0 , 0'))
        self.assign_comp_editbox.setText(self.setting.value('assign_comp_coord', ' 0 , 0'))
        self.clear_elem_editbox.setText(self.setting.value('clear_element_coord', ' 0 , 0'))
        self.elem_name_editbox.setText(self.setting.value('element_name', 'W1_1_B2'))
        self.setting.endGroup()
        # self.drag_duration_slider.setValue(2)        
        # Tab 3
        self.setting.beginGroup('setting_tab3')
        self.load_cbeam_checkbox.setChecked(self.setting.value('load_cbeam', True, type=bool))
        self.load_wall_checkbox.setChecked(self.setting.value('load_wall', True, type=bool))
        self.load_ecol_checkbox.setChecked(self.setting.value('load_ecolumn', False, type=bool))
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
        self.ecol_pdf_checkbox.setChecked(self.setting.value('ecolumn_to_pdf', False, type=bool))
        self.wall_pdf_checkbox.setChecked(self.setting.value('wall_to_pdf', True, type=bool))
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
        self.coord_shortcut = QShortcut(QKeySequence('F10'), self) # Shortcuts 설정
        self.coord_shortcut.activated.connect(self.get_mouse_coords)
        self.start_shortcut = QShortcut(QKeySequence('F2'), self)
        self.start_shortcut.activated.connect(self.run_worker10) # 1: start, 0: end
        # self.end_shortcut = QShortcut(QKeySequence('Escape'), self)
        # self.end_shortcut.activated.connect(self.run_worker10(0))
        # Tab 3
        self.load_result_btn.clicked.connect(self.run_worker5)
        self.print_docx_btn.clicked.connect(self.run_worker6)
        self.design_wall_btn.clicked.connect(self.run_worker7)
        self.print_pdf_btn.clicked.connect(self.run_worker8)
        self.preview_result_btn.clicked.connect(self.run_worker9)
        self.print_hwp_btn.clicked.connect(self.run_worker11)
        
        # Child Windows
        self.BR_setting_btn.clicked.connect(self.open_BRSettingWindow)
        self.print_setting_btn.clicked.connect(self.open_PrintSettingWindow)
        
        # Menu Bar
        self.action_about.triggered.connect(self.open_about)
        self.action_release_note.triggered.connect(self.open_release_note)
        self.action_docs.triggered.connect(self.open_docs)
        self.action_sheets.triggered.connect(self.open_sheets_folder)
        
        ### Child Windows
        # BR Setting 창
        self.win_BR = QMainWindow()
        self.win_BR = BRSettingWindow(self.status_browser) # child window의 에러를 status_browser에 입력하기 위해
        
        # Print Setting 창
        self.win_print = QMainWindow()
        self.win_print = PrintSettingWindow(self.status_browser) # child window의 에러를 status_browser에 입력하기 위해

        # About Setting 창
        self.win_about = QMainWindow()
        self.win_about = AboutWindow()
        
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
        
        ### etc
        # Icon 설정
        self.setWindowIcon(QIcon('./images/icon_earthquake.ico'))
        
        # 마우스 좌표 real-time tracking
        self.setMouseTracking(True)

    # 파일 선택 function (using QFileDialog)
    def find_input_xlsx(self): # Data Conversion Sheets
        # global input_xlsx_path
        input_xlsx_path = QFileDialog.getOpenFileName(parent=self, caption='Choose Data Conversion Sheets'
                                    , directory=os.getcwd(), filter='Excel File (*.xlsx *.xls)')[0]
        if input_xlsx_path != '':
            self.data_conv_path_editbox.setText(input_xlsx_path)
        
    def find_result_xlsx(self): # Analysis Resutls
        # global input_xlsx_path
        result_xlsx_path = QFileDialog.getOpenFileNames(parent=self, caption='Choose Analysis Result Sheets (Multiple Files Available)'
                                    , directory=os.getcwd(), filter='Excel File (*.xlsx *.xls)')[0]
        if len(result_xlsx_path) != 0:
            all_result_xlsx_path = ['"%s"' %file_name for file_name in result_xlsx_path]
            joined_result_xlsx_path = ','.join(all_result_xlsx_path)
            self.result_path_editbox.setText(joined_result_xlsx_path)
            self.display_selected_result_path.setText('%i files selected' %len(result_xlsx_path))
        
    def find_wall_design_xlsx(self): # Wall Results Sheets
        # global input_xlsx_path
        wall_design_xlsx_path = QFileDialog.getOpenFileName(parent=self, caption='Choose Seismic Design Sheets (Shear Wall)'
                                    , directory=os.getcwd(), filter='Excel File (*.xlsx *.xls)')[0]        
        if wall_design_xlsx_path != '':
            self.wall_design_path_editbox.setText(wall_design_xlsx_path)
        
    def find_beam_design_xlsx(self): # Beam Results Sheets
        # global input_xlsx_path
        beam_design_xlsx_path = QFileDialog.getOpenFileName(parent=self, caption='Choose Seismic Design Sheets (Coupling Beam)'
                                    , directory=os.getcwd(), filter='Excel File (*.xlsx *.xls)')[0]        
        if beam_design_xlsx_path != '':
            self.beam_design_path_editbox.setText(beam_design_xlsx_path)
    
    def find_col_design_xlsx(self): # Column Results Sheets
        # global input_xlsx_path
        col_design_xlsx_path = QFileDialog.getOpenFileName(parent=self, caption='Choose Seismic Design Sheets (Elastic Column)'
                                    , directory=os.getcwd(), filter='Excel File (*.xlsx *.xls)')[0]        
        if col_design_xlsx_path != '':
            self.col_design_path_editbox.setText(col_design_xlsx_path)
    
    # Macro 사용을 위한 function
    def get_mouse_coords(self):
        # 현재 focus된 widget
        focused_widget = QApplication.focusWidget()                
        # 마우스 좌표를 입력할 LineEdit widget list(in order)
        coord_widget_list = [self.up_left_coord_editbox, self.up_right_coord_editbox
                             , self.low_left_coord_editbox, self.low_right_coord_editbox
                             , self.p3d_bar_editbox, self.add_cuts_editbox, self.delete_cuts_editbox
                             , self.ok_editbox, self.next_section_editbox, self.next_frame_editbox
                             , self.ok_delete_editbox, self.missing_data_editbox
                             , self.assign_comp_editbox, self.clear_elem_editbox]
        # 마우스 좌표 입력 LineEdit이 focus되어 있는 경우, 입력값 받음
        if focused_widget in coord_widget_list:
            focused_widget.setText(' %d , %d' %(pag.position().x, pag.position().y))
            # 다음 LineEdit으로 focus 변경
            focused_idx = coord_widget_list.index(focused_widget)
            next_widget = coord_widget_list[np.clip(focused_idx + 1, 0, 13)] 
            # np.clip: Clip (limit) the values, ex) if last editbox is focused, prevent to go to last + 1 editbox 
            next_widget.setFocus()
    
    def threshold1(self, value):
        self.drag_duration_slider = value
        self.test.setText("Value : " + str(value))

    def threshold2(self, value):
        self.offset_slider = value
    
    # Child Window 또는 Menu Bar 띄우기 function
    def open_BRSettingWindow(self):
        self.win_BR.show()
        
    def open_PrintSettingWindow(self):
        self.win_print.show()
        
    def open_about(self):
        self.win_about.show()
        
    def open_release_note(self):
        note = os.path.join(os.getcwd(), 'docs/PBSD/build/html/release_note.html')
        
        if sys.platform == 'linux2':
            subprocess.call(["xdg-open", note])
        else:
            os.startfile(note)
            
    def open_docs(self):
        docs = os.path.join(os.getcwd(), 'docs/PBSD/build/html/pbd_p3d_manual.html')
        
        if sys.platform == 'linux2':
            subprocess.call(["xdg-open", note])
        else:
            os.startfile(docs)

    def open_sheets_folder(self):
       folder = r'K:\2104-박재성\성능기반 내진설계\Excel Sheet'
       
       os.startfile(folder)

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
            self.setting.setValue('upleft_coord', self.up_left_coord_editbox.text())
            self.setting.setValue('upright_coord', self.up_right_coord_editbox.text())
            self.setting.setValue('lowleft_coord', self.low_left_coord_editbox.text())
            self.setting.setValue('lowright_coord', self.low_right_coord_editbox.text())
            self.setting.setValue('p3d_bar_coord', self.p3d_bar_editbox.text())
            self.setting.setValue('add_cuts_coord', self.add_cuts_editbox.text())
            self.setting.setValue('delete_cuts_coord', self.delete_cuts_editbox.text())
            self.setting.setValue('ok_coord', self.ok_editbox.text())
            self.setting.setValue('next_section_coord', self.next_section_editbox.text())
            self.setting.setValue('next_frame_coord', self.next_frame_editbox.text())
            self.setting.setValue('ok_delete_coord', self.ok_delete_editbox.text())
            self.setting.setValue('missing_data_coord', self.missing_data_editbox.text())
            self.setting.setValue('assign_comp_coord', self.assign_comp_editbox.text())
            self.setting.setValue('clear_element_coord', self.clear_elem_editbox.text())
            self.setting.setValue('element_name', self.elem_name_editbox.text())
            self.setting.endGroup()

            self.setting.beginGroup('setting_tab3')
            self.setting.setValue('load_cbeam', self.load_cbeam_checkbox.isChecked())
            self.setting.setValue('load_wall', self.load_wall_checkbox.isChecked())
            self.setting.setValue('load_ecolumn', self.load_ecol_checkbox.isChecked())
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
