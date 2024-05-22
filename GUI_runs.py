import sys
import os
import time
import pandas as pd
import pickle
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QSettings, QCoreApplication, QThread, Qt

import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
from matplotlib.backends.backend_qt5agg import FigureCanvas as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure

import PBD_p3d as pbd
from GUI_workers import *

#%% Tab1 - Import
def run_worker1(self):    
    # 시작 메세지
    time_start = time.time()
    self.status_browser.append('Running.....')
    
    # 변수 정리
    input_xlsx_path = self.data_conv_path_editbox.text().strip()
    DL = self.DL_name_editbox.text().strip()
    LL = self.LL_name_editbox.text().strip()
    import_node = self.import_node_checkbox.isChecked()
    import_beam = self.import_beam_checkbox.isChecked()
    import_col = self.import_col_checkbox.isChecked()
    import_wall = self.import_wall_checkbox.isChecked()
    import_plate = self.import_plate_checkbox.isChecked()
    import_WR_gage = self.import_WR_gage_checkbox.isChecked()
    import_WAS_gage = self.import_WAS_gage_checkbox.isChecked()
    import_I_beam = self.import_I_beam_checkbox.isChecked()
    import_mass = self.import_mass_checkbox.isChecked()
    import_nodal_load = self.import_nodal_load_checkbox.isChecked()
    
    # checkbox or editbox 비어있는 경우 break
    while True:
        if (import_node == False) & (import_beam == False) & (import_col == False)\
            & (import_wall == False) & (import_plate == False) & (import_WR_gage == False)\
            & (import_WAS_gage == False) & (import_I_beam == False) & (import_mass == False)\
            & (import_nodal_load == False):
            msg = 'Nothing Checked!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return # return nothing & break the function
        elif (input_xlsx_path == '') | (DL == False) | (LL == False):
            msg = 'Nothing Entered!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        else: break
    
    
    # QThread 오브젝트 생성
    self.thread = QThread(parent=self)
    # ImportWorker 오브젝트 생성
    self.worker = pbd.ImportWorker(input_xlsx_path, DL, LL, import_node, import_beam
                                , import_col, import_wall, import_plate
                                , import_WR_gage, import_WAS_gage, import_I_beam
                                , import_mass, import_nodal_load, time_start) # Create a worker object
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
    self.insert_force_btn.setEnabled(False)
    self.load_result_btn.setEnabled(False)
    self.design_wall_btn.setEnabled(False)
    self.print_pdf_btn.setEnabled(False)
    self.print_docx_btn.setEnabled(False)
    self.print_hwp_btn.setEnabled(False)
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_hwp_btn.setEnabled(True))
    
    # 완료 메세지
    self.worker.msg.connect(self.msg_fn)

#%% Tab1 - Name
def run_worker2(self):
    # 시작 메세지
    time_start = time.time()
    self.status_browser.append('Running.....')
    
    # 변수 정리
    input_xlsx_path = self.data_conv_path_editbox.text().strip()
    drift_pos_raw = self.drift_pos_editbox.text().strip()
    drift_position = []
    for i in range(drift_pos_raw.count(',')+1):
        drift_pos_elem = drift_pos_raw.split(',')[i].strip()
        drift_position.append(drift_pos_elem)
        
    # checkbox or editbox 비어있는 경우 break
    while True:
        if (input_xlsx_path == '') | (drift_pos_raw == False):
            msg = 'Nothing Entered!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        else: break
    
    # QThread 오브젝트 생성
    self.thread = QThread(parent=self) # Create a QThread object
    # NameWorker 오브젝트 생성
    self.worker = pbd.NameWorker(input_xlsx_path, drift_position, time_start) # Create a worker object
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
    self.insert_force_btn.setEnabled(False)
    self.load_result_btn.setEnabled(False)
    self.design_wall_btn.setEnabled(False)
    self.print_pdf_btn.setEnabled(False)
    self.print_docx_btn.setEnabled(False)
    self.print_hwp_btn.setEnabled(False)
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_hwp_btn.setEnabled(True))
    
    # 완료 메세지 print
    self.worker.msg.connect(self.msg_fn)
    
#%% Tab1 - Convert    
def run_worker3(self):
    # 시작 메세지
    time_start = time.time()
    self.status_browser.append('Running.....')
    
    # 변수 정리
    input_xlsx_path = self.data_conv_path_editbox.text().strip()
    get_cbeam = self.convert_cbeam_checkbox.isChecked()
    get_gbeam = self.convert_gbeam_checkbox.isChecked()
    get_ebeam = self.convert_ebeam_checkbox.isChecked()
    get_gcol = self.convert_gcol_checkbox.isChecked()
    get_ecol = self.convert_ecol_checkbox.isChecked()
    get_wall = self.convert_wall_checkbox.isChecked()
    
    # checkbox or editbox 비어있는 경우 break
    while True:
        if (get_cbeam == False) & (get_gbeam == False) & (get_ebeam == False)\
            & (get_gcol == False) & (get_ecol == False) & (get_wall == False):
            msg = 'Nothing Checked!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        elif (input_xlsx_path == ''):
            msg = 'Nothing Entered!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        else: break
    
    # QThread 오브젝트 생성
    self.thread = QThread(parent=self) # Create a QThread object
    # ConvertWorker 오브젝트 생성
    self.worker = pbd.ConvertWorker(input_xlsx_path, get_cbeam, get_gbeam
                                    , get_ebeam, get_gcol, get_ecol, get_wall
                                    , time_start) # Create a worker object
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
    self.insert_force_btn.setEnabled(False)
    self.load_result_btn.setEnabled(False)
    self.design_wall_btn.setEnabled(False)
    self.print_pdf_btn.setEnabled(False)
    self.print_docx_btn.setEnabled(False)
    self.print_hwp_btn.setEnabled(False)
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_hwp_btn.setEnabled(True))
    
    # 완료 메세지 print
    self.worker.msg.connect(self.msg_fn)

#%% Tab2 - Insert    
def run_worker4(self):
    # 시작 메세지
    time_start = time.time()
    self.status_browser.append('Running.....')
    
    # 변수 정리
    result_xlsx_path = self.result_path_editbox.text()
    result_xlsx_path = result_xlsx_path.split('"')
    result_xlsx_path = [i.strip() for i in result_xlsx_path if len(i) > 2]
    input_xlsx_path = self.data_conv_path_editbox.text().strip()
    get_gbeam = self.insert_gbeam_checkbox.isChecked()
    get_gcol = self.insert_gcol_checkbox.isChecked()
    get_ecol = self.insert_ecol_checkbox.isChecked()
    
    # checkbox or editbox 비어있는 경우 break
    while True:
        if (get_gbeam == False) & (get_gcol == False) & (get_ecol == False):
            msg = 'Nothing Checked!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        elif (input_xlsx_path == '') | (len(result_xlsx_path) == 0):
            msg = 'Nothing Entered!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        else: break
    
    # QThread 오브젝트 생성
    self.thread = QThread(parent=self) # Create a QThread object
    # InsertWorker 오브젝트 생성
    self.worker = pbd.InsertWorker(input_xlsx_path, result_xlsx_path
                                   , get_gbeam, get_gcol, get_ecol
                                   , time_start) # Create a worker object
    self.worker.moveToThread(self.thread) # Move worker to the thread
    
    # Connect signals and slots
    self.thread.started.connect(self.worker.insert_force_fn)
    self.worker.finished.connect(self.thread.quit)
    self.worker.finished.connect(self.worker.deleteLater)
    self.thread.finished.connect(self.thread.deleteLater)
    
    # Start the thread
    self.thread.start()
    
    # Enable/Disable the Button
    self.import_midas_btn.setEnabled(False)
    self.print_name_btn.setEnabled(False)
    self.convert_prop_btn.setEnabled(False)
    self.insert_force_btn.setEnabled(False)
    self.load_result_btn.setEnabled(False)
    self.design_wall_btn.setEnabled(False)
    self.print_pdf_btn.setEnabled(False)
    self.print_docx_btn.setEnabled(False)
    self.print_hwp_btn.setEnabled(False)
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_hwp_btn.setEnabled(True))
    
    # 완료 메세지 print
    self.worker.msg.connect(self.msg_fn)

#%% Tab3 - Load
def run_worker5(self):
    # 시작 메세지
    time_start = time.time()
    self.status_browser.append('Running.....')
    
    # 변수 정리
    result_xlsx_path = self.result_path_editbox.text()
    result_xlsx_path = result_xlsx_path.split('"')
    result_xlsx_path = [i.strip() for i in result_xlsx_path if len(i) > 2]
    input_xlsx_path = self.data_conv_path_editbox.text().strip()
    wall_design_xlsx_path = self.wall_design_path_editbox.text().strip()
    beam_design_xlsx_path = self.beam_design_path_editbox.text().strip()
    col_design_xlsx_path = self.col_design_path_editbox.text().strip()
    get_cbeam = self.load_cbeam_checkbox.isChecked()
    get_wall = self.load_wall_checkbox.isChecked()
    get_ecol = self.load_ecol_checkbox.isChecked()
    
    BR_scale_factor = self.win_BR.BR_scale_factor_editbox.text().strip()
    BR_scale_factor = float(BR_scale_factor)
    
    # Worker 클래스에 입력할 Keyword Arguments Dict 생성
    kwargs = {'result_xlsx_path': result_xlsx_path,
              'input_xlsx_path': input_xlsx_path,
              'wall_design_xlsx_path': wall_design_xlsx_path,
              'beam_design_xlsx_path': beam_design_xlsx_path,
              'col_design_xlsx_path': col_design_xlsx_path,
              'get_cbeam': get_cbeam,
              'get_wall': get_wall,
              'get_ecol': get_ecol,
              'BR_scale_factor': BR_scale_factor,
              'time_start': time_start
              }
    
    # checkbox or editbox 비어있는 경우 break
    while True:
        if (get_cbeam == False) & (get_wall == False) & (get_ecol == False):
            msg = 'Nothing Checked!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        elif (input_xlsx_path == '') | (len(result_xlsx_path) == 0)\
            | (BR_scale_factor == ''):
            msg = 'Nothing Entered!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        elif (wall_design_xlsx_path == '') & (beam_design_xlsx_path == '')\
            & (col_design_xlsx_path == ''):
            msg = 'Nothing Entered!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        else: break

    # QThread 오브젝트 생성
    self.thread = QThread(parent=self) # Create a QThread object
    # InsertWorker 오브젝트 생성
    self.worker = LoadWorker(**kwargs) # Create a worker object
    self.worker.moveToThread(self.thread) # Move worker to the thread
    
    # Connect signals and slots
    self.thread.started.connect(self.worker.load_result_fn)
    self.worker.finished.connect(self.thread.quit)
    self.worker.finished.connect(self.worker.deleteLater)
    self.thread.finished.connect(self.thread.deleteLater)
    
    # Start the thread
    self.thread.start()
    
    # Enable/Disable the Button
    self.import_midas_btn.setEnabled(False)
    self.print_name_btn.setEnabled(False)
    self.convert_prop_btn.setEnabled(False)
    self.insert_force_btn.setEnabled(False)
    self.load_result_btn.setEnabled(False)
    self.design_wall_btn.setEnabled(False)
    self.print_pdf_btn.setEnabled(False)
    self.print_docx_btn.setEnabled(False)
    self.preview_result_btn.setEnabled(False)
    self.print_hwp_btn.setEnabled(False)
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.preview_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_hwp_btn.setEnabled(True))
    
    # 실패할 경우, 그냥 msg_fn 이용해서 오류 메세지 내보내기
    # (단, 여기서 오류가 발생하는 경우, loadworker에서 데이터 처리하는 과정에서의 오류밖에 캐치 안됨)
    self.worker.msg.connect(self.msg_fn)

#%% Tab3 - Print to docx
def run_worker6(self):
    # 시작 메세지
    time_start = time.time()
    self.status_browser.append('Running.....')
    
    # 변수 정리
    result_xlsx_path = self.result_path_editbox.text()
    result_xlsx_path = result_xlsx_path.split('"')
    result_xlsx_path = [i.strip() for i in result_xlsx_path if len(i) > 2]
    get_base_SF = self.base_SF_checkbox.isChecked()
    get_story_SF = self.story_SF_checkbox.isChecked()
    get_IDR = self.IDR_checkbox.isChecked()
    get_BR = self.BR_checkbox.isChecked()
    get_BSF = self.BSF_checkbox.isChecked()
    get_E_BSF = self.E_BSF_checkbox.isChecked()
    get_CR = self.CR_checkbox.isChecked()
    get_CSF = self.CSF_checkbox.isChecked()
    get_E_CSF = self.E_CSF_checkbox.isChecked()
    get_WAS = self.WAS_checkbox.isChecked()
    get_WR = self.WR_checkbox.isChecked()
    get_WSF = self.WSF_checkbox.isChecked()

    project_name = self.win_print.project_name_editbox.text().strip()
    bldg_name = self.win_print.bldg_name_editbox.text().strip()
    story_gap = self.win_print.story_gap_editbox.text().strip()
    max_shear = self.win_print.max_shear_editbox.text().strip()

    story_gap = int(story_gap)
    max_shear = int(max_shear)
    
    # Worker 클래스에 입력할 Keyword Arguments Dict 생성
    kwargs = {'result_xlsx_path': result_xlsx_path,
              'get_base_SF': get_base_SF,
              'get_story_SF': get_story_SF,
              'get_IDR': get_IDR,
              'get_BR': get_BR,
              'get_BSF': get_BSF,
              'get_E_BSF': get_E_BSF,
              'get_CR': get_CR,
              'get_CSF': get_CSF,
              'get_E_CSF': get_E_CSF,
              'get_WAS': get_WAS,
              'get_WR': get_WR,
              'get_WSF': get_WSF,
              'project_name': project_name,
              'bldg_name': bldg_name,
              'story_gap': story_gap,
              'max_shear': max_shear,
              'time_start': time_start}

    # checkbox or editbox 비어있는 경우 break
    while True:
        if (get_base_SF == False) & (get_story_SF == False) & (get_IDR == False)\
            & (get_BR == False) & (get_BSF == False) & (get_E_BSF == False)\
            & (get_CR == False) & (get_CSF == False) & (get_E_CSF == False)\
            & (get_WAS == False) & (get_WR == False) & (get_WSF == False):
            msg = 'Nothing Checked!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        elif (len(result_xlsx_path) == 0)\
            | (project_name == '') | (bldg_name == '')\
            | (story_gap == '') | (max_shear == ''):
            msg = 'Nothing Entered!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        else: break

    # QThread 오브젝트 생성
    self.thread = QThread(parent=self) # Create a QThread object
    # PdfWorker 오브젝트 생성
    self.worker = DocxWorker(**kwargs) # Create a worker object
    self.worker.moveToThread(self.thread) # Move worker to the thread
    
    # Connect signals and slots
    self.thread.started.connect(self.worker.print_docx_fn)
    self.worker.finished.connect(self.thread.quit)
    self.worker.finished.connect(self.worker.deleteLater)
    self.thread.finished.connect(self.thread.deleteLater)

    # Start the thread
    self.thread.start()
    
    # Enable/Disable the Button
    self.import_midas_btn.setEnabled(False)
    self.print_name_btn.setEnabled(False)
    self.convert_prop_btn.setEnabled(False)
    self.insert_force_btn.setEnabled(False)
    self.load_result_btn.setEnabled(False)
    self.design_wall_btn.setEnabled(False)
    self.print_pdf_btn.setEnabled(False)
    self.print_docx_btn.setEnabled(False)
    self.preview_result_btn.setEnabled(False)
    self.print_hwp_btn.setEnabled(False)
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.preview_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_hwp_btn.setEnabled(True))
    
    # 완료 메세지 print
    self.worker.msg.connect(self.msg_fn)

#%% Tab3 - Design
def run_worker7(self):
    # 시작 메세지
    time_start = time.time()
    self.status_browser.append('Running.....')
    
    # 변수 정리
    wall_design_xlsx_path = self.wall_design_path_editbox.text().strip()
    
    # checkbox or editbox 비어있는 경우 break
    while True:
        if (wall_design_xlsx_path == ''):
            msg = 'Nothing Entered!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        else: break

    # QThread 오브젝트 생성
    self.thread = QThread(parent=self) # Create a QThread object
    # InsertWorker 오브젝트 생성
    self.worker = RedesignWorker(wall_design_xlsx_path, time_start) # Create a worker object
    self.worker.moveToThread(self.thread) # Move worker to the thread
    
    # Connect signals and slots
    self.thread.started.connect(self.worker.redesign_wall_fn)
    self.worker.finished.connect(self.thread.quit)
    self.worker.finished.connect(self.worker.deleteLater)
    self.thread.finished.connect(self.thread.deleteLater)
    
    # Start the thread
    self.thread.start()
    
    # Enable/Disable the Button
    self.import_midas_btn.setEnabled(False)
    self.print_name_btn.setEnabled(False)
    self.convert_prop_btn.setEnabled(False)
    self.insert_force_btn.setEnabled(False)
    self.load_result_btn.setEnabled(False)
    self.design_wall_btn.setEnabled(False)
    self.print_pdf_btn.setEnabled(False)
    self.print_docx_btn.setEnabled(False)
    self.preview_result_btn.setEnabled(False)
    self.print_hwp_btn.setEnabled(False)
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.preview_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_hwp_btn.setEnabled(True))
    
    # 완료 메세지 print
    self.worker.msg.connect(self.msg_fn)
    
#%% Tab3 - Print to pdf
def run_worker8(self):
    # 시작 메세지
    time_start = time.time()
    self.status_browser.append('Running.....')
    
    # 변수 정리
    wall_design_xlsx_path = self.wall_design_path_editbox.text().strip()
    beam_design_xlsx_path = self.beam_design_path_editbox.text().strip()
    col_design_xlsx_path = self.col_design_path_editbox.text().strip()
    get_cbeam = self.cbeam_pdf_checkbox.isChecked()
    get_ecol = self.ecol_pdf_checkbox.isChecked()
    get_wall = self.wall_pdf_checkbox.isChecked()
    
    project_name = self.win_print.project_name_editbox.text().strip()
    bldg_name = self.win_print.bldg_name_editbox.text().strip()
    
    # Worker 클래스에 입력할 Keyword Arguments Dict 생성
    kwargs = {'wall_design_xlsx_path':wall_design_xlsx_path, 
              'beam_design_xlsx_path':beam_design_xlsx_path,
              'col_design_xlsx_path':col_design_xlsx_path,
              'get_cbeam':get_cbeam,
              'get_ecol':get_ecol,
              'get_wall':get_wall,
              'project_name':project_name,
              'bldg_name':bldg_name,
              'time_start':time_start}
    
    # checkbox or editbox 비어있는 경우 break
    while True:
        if (get_cbeam == False) & (get_ecol == False) & (get_wall == False):
            msg = 'Nothing Checked!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        elif (project_name == '') | (bldg_name == ''):
            msg = 'Nothing Entered! (Project Name, Building Name)'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        elif (wall_design_xlsx_path == '') & (beam_design_xlsx_path == '')\
            & (col_design_xlsx_path == ''):
            msg = 'Nothing Entered!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        else: break

    # QThread 오브젝트 생성
    self.thread = QThread(parent=self) # Create a QThread object
    # PdfWorker 오브젝트 생성
    self.worker = PdfWorker(**kwargs) # Create a worker object
    self.worker.moveToThread(self.thread) # Move worker to the thread
    
    # Connect signals and slots
    self.thread.started.connect(self.worker.print_pdf_fn)
    self.worker.finished.connect(self.thread.quit)
    self.worker.finished.connect(self.worker.deleteLater)
    self.thread.finished.connect(self.thread.deleteLater)
    
    # Start the thread
    self.thread.start()
    
    # Enable/Disable the Button
    self.import_midas_btn.setEnabled(False)
    self.print_name_btn.setEnabled(False)
    self.convert_prop_btn.setEnabled(False)
    self.insert_force_btn.setEnabled(False)
    self.load_result_btn.setEnabled(False)
    self.design_wall_btn.setEnabled(False)
    self.print_pdf_btn.setEnabled(False)
    self.print_docx_btn.setEnabled(False)
    self.preview_result_btn.setEnabled(False)
    self.print_hwp_btn.setEnabled(False)
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.preview_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_hwp_btn.setEnabled(True))
    
    # 완료 메세지 print
    self.worker.msg.connect(self.msg_fn)
    
    
#%% Tab3 - Preview
def run_worker9(self):
    # 시작 메세지
    time_start = time.time()
    self.status_browser.append('Running.....')
    
    # 변수 정리 (dictionary)  
    # story_gap = int(story_gap)
    # max_shear = int(max_shear)
    story_gap = 2
    max_shear = 60000
    
    result_xlsx_path = self.result_path_editbox.text()
    result_xlsx_path = result_xlsx_path.split('"')
    result_xlsx_path = [i.strip() for i in result_xlsx_path if len(i) > 2]    
    input_xlsx_path = self.data_conv_path_editbox.text().strip()
    wall_design_xlsx_path = self.wall_design_path_editbox.text().strip()
    beam_design_xlsx_path = self.beam_design_path_editbox.text().strip()
    col_design_xlsx_path = self.col_design_path_editbox.text().strip()
    get_base_SF = self.base_SF_checkbox.isChecked()
    get_story_SF = self.story_SF_checkbox.isChecked()
    get_IDR = self.IDR_checkbox.isChecked()
    get_BR = self.BR_checkbox.isChecked()
    get_BSF = self.BSF_checkbox.isChecked()
    get_E_BSF = self.E_BSF_checkbox.isChecked()
    get_CR = self.CR_checkbox.isChecked()
    get_CSF = self.CSF_checkbox.isChecked()
    get_E_CSF = self.E_CSF_checkbox.isChecked()
    get_WAS = self.WAS_checkbox.isChecked()
    get_WR = self.WR_checkbox.isChecked()
    get_WSF = self.WSF_checkbox.isChecked()
    
    # Worker 클래스에 입력할 Keyword Arguments Dict 생성
    kwargs = {'result_xlsx_path':result_xlsx_path, 
              'input_xlsx_path':input_xlsx_path, 
              'wall_design_xlsx_path':wall_design_xlsx_path, 
              'beam_design_xlsx_path':beam_design_xlsx_path,
              'col_design_xlsx_path':col_design_xlsx_path,
              'get_base_SF':get_base_SF,
              'get_story_SF':get_story_SF,
              'get_IDR':get_IDR,
              'get_BR':get_BR,
              'get_BSF':get_BSF,
              'get_E_BSF':get_E_BSF,
              'get_CR':get_CR,
              'get_CSF':get_CSF,
              'get_E_CSF':get_E_CSF,
              'get_WAS':get_WAS,
              'get_WR':get_WR,
              'get_WSF':get_WSF,
              'story_gap':story_gap,
              'max_shear':max_shear,
              'time_start':time_start}
    
    # checkbox or editbox 비어있는 경우 함수종료
    while True:
        if (get_base_SF == False) & (get_story_SF == False) & (get_IDR == False)\
            & (get_BR == False) & (get_BSF == False) & (get_E_BSF == False)\
            & (get_CR == False) & (get_CSF == False) & (get_E_CSF == False)\
            & (get_WAS == False) & (get_WR == False) & (get_WSF == False):
            msg = 'Nothing Checked!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return # 함수 종료
        elif (input_xlsx_path == '') | (len(result_xlsx_path) == 0):
            msg = 'Nothing Entered!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        elif (wall_design_xlsx_path == '') & (beam_design_xlsx_path == '')\
            & (col_design_xlsx_path == ''):
                if (get_base_SF == True) | (get_story_SF == True) | (get_IDR == True):
                    break
                else:
                    msg = 'Nothing Entered!'
                    msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
                    self.status_browser.append(msg_colored)
                    return
        elif ((get_BR == True) | (get_BSF == True)) & (beam_design_xlsx_path == ''):
            msg = 'Nothing Entered! (Seismic Design Coupling Beam)'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        elif ((get_WAS == True) | (get_WR == True) | (get_WSF == True)) & (wall_design_xlsx_path == ''):
            msg = 'Nothing Entered! (Seismic Design Shear Wall)'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        else: break

    # QThread 오브젝트 생성
    self.thread = QThread(parent=self) # Create a QThread object
    # InsertWorker 오브젝트 생성
    self.worker = PreviewWorker(**kwargs) # Create a worker object
    self.worker.moveToThread(self.thread) # Move worker to the thread
    
    # Connect signals and slots
    self.thread.started.connect(self.worker.preview_result_fn)
    self.worker.finished.connect(self.thread.quit)
    self.worker.finished.connect(self.worker.deleteLater)
    self.thread.finished.connect(self.thread.deleteLater)
    
    # Start the thread
    self.thread.start()
    
    # Send Signals(incl. result data) to plot_display/load_time_count functions
    self.worker.result_data.connect(self.plot_display)
    # self.worker.msg.connect(self.load_time_count)
    
    # Enable/Disable the Button
    self.import_midas_btn.setEnabled(False)
    self.print_name_btn.setEnabled(False)
    self.convert_prop_btn.setEnabled(False)
    self.insert_force_btn.setEnabled(False)
    self.load_result_btn.setEnabled(False)
    self.design_wall_btn.setEnabled(False)
    self.print_pdf_btn.setEnabled(False)
    self.print_docx_btn.setEnabled(False)
    self.preview_result_btn.setEnabled(False)
    self.print_hwp_btn.setEnabled(False)
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.preview_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_hwp_btn.setEnabled(True))
    
    # 실패할 경우, 그냥 msg_fn 이용해서 오류 메세지 내보내기
    # (단, 여기서 오류가 발생하는 경우, loadworker에서 데이터 처리하는 과정에서의 오류밖에 캐치 안됨)
    self.worker.msg.connect(self.msg_fn)
    
#%% Tab3 - Print to hwp
def run_worker11(self):
    # 시작 메세지
    time_start = time.time()
    self.status_browser.append('Running.....')
    
    # 변수 정리
    result_xlsx_path = self.result_path_editbox.text()
    result_xlsx_path = result_xlsx_path.split('"')
    result_xlsx_path = [i.strip() for i in result_xlsx_path if len(i) > 2]
    get_base_SF = self.base_SF_checkbox.isChecked()
    get_story_SF = self.story_SF_checkbox.isChecked()
    get_IDR = self.IDR_checkbox.isChecked()
    get_BR = self.BR_checkbox.isChecked()
    get_BSF = self.BSF_checkbox.isChecked()
    get_E_BSF = self.E_BSF_checkbox.isChecked()
    get_CR = self.CR_checkbox.isChecked()
    get_CSF = self.CSF_checkbox.isChecked()
    get_E_CSF = self.E_CSF_checkbox.isChecked()
    get_WAS = self.WAS_checkbox.isChecked()
    get_WR = self.WR_checkbox.isChecked()
    get_WSF = self.WSF_checkbox.isChecked()

    project_name = self.win_print.project_name_editbox.text().strip()
    bldg_name = self.win_print.bldg_name_editbox.text().strip()
    story_gap = self.win_print.story_gap_editbox.text().strip()
    max_shear = self.win_print.max_shear_editbox.text().strip()

    story_gap = int(story_gap)
    max_shear = int(max_shear)
    
    # Worker 클래스에 입력할 Keyword Arguments Dict 생성
    kwargs = {'result_xlsx_path': result_xlsx_path,
              'get_base_SF': get_base_SF,
              'get_story_SF': get_story_SF,
              'get_IDR': get_IDR,
              'get_BR': get_BR,
              'get_BSF': get_BSF,
              'get_E_BSF': get_E_BSF,
              'get_CR': get_CR,
              'get_CSF': get_CSF,
              'get_E_CSF': get_E_CSF,
              'get_WAS': get_WAS,
              'get_WR': get_WR,
              'get_WSF': get_WSF,
              'project_name': project_name,
              'bldg_name': bldg_name,
              'story_gap': story_gap,
              'max_shear': max_shear,
              'time_start': time_start}

    # checkbox or editbox 비어있는 경우 break
    while True:
        if (get_base_SF == False) & (get_story_SF == False) & (get_IDR == False)\
            & (get_BR == False) & (get_BSF == False) & (get_E_BSF == False)\
            & (get_CR == False) & (get_CSF == False) & (get_E_CSF == False)\
            & (get_WAS == False) & (get_WR == False) & (get_WSF == False):
            msg = 'Nothing Checked!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        elif (len(result_xlsx_path) == 0)\
            | (project_name == '') | (bldg_name == '')\
            | (story_gap == '') | (max_shear == ''):
            msg = 'Nothing Entered!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        else: break

    # QThread 오브젝트 생성
    self.thread = QThread(parent=self) # Create a QThread object
    # PdfWorker 오브젝트 생성
    self.worker = HwpWorker(**kwargs) # Create a worker object
    self.worker.moveToThread(self.thread) # Move worker to the thread
    
    # Connect signals and slots
    self.thread.started.connect(self.worker.print_hwp_fn)
    self.worker.finished.connect(self.thread.quit)
    self.worker.finished.connect(self.worker.deleteLater)
    self.thread.finished.connect(self.thread.deleteLater)

    # Start the thread
    self.thread.start()
    
    # Enable/Disable the Button
    self.import_midas_btn.setEnabled(False)
    self.print_name_btn.setEnabled(False)
    self.convert_prop_btn.setEnabled(False)
    self.insert_force_btn.setEnabled(False)
    self.load_result_btn.setEnabled(False)
    self.design_wall_btn.setEnabled(False)
    self.print_pdf_btn.setEnabled(False)
    self.print_docx_btn.setEnabled(False)
    self.preview_result_btn.setEnabled(False)
    self.print_hwp_btn.setEnabled(False)
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.preview_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_hwp_btn.setEnabled(True))
    
    # 완료 메세지 print
    self.worker.msg.connect(self.msg_fn)

#%% Tab2 - Macro    
def run_worker10(self):
    # 시작 메세지
    time_start = time.time()
    self.status_browser.append('Running.....')
    
    # 변수 정리
    input_xlsx_path = self.data_conv_path_editbox.text().strip()
    macro_mode = self.macro_comboBox.currentText()
    start_or_end = 'None'
    
    pos_lefttop = self.up_left_coord_editbox.text()
    pos_righttop = self.up_right_coord_editbox.text()
    pos_leftbot = self.low_left_coord_editbox.text()
    pos_rightbot = self.low_right_coord_editbox.text()
    pos_p3dbar = self.p3d_bar_editbox.text()
    pos_addcuts = self.add_cuts_editbox.text()
    pos_deletecuts = self.delete_cuts_editbox.text()
    pos_ok = self.ok_editbox.text()
    pos_nextsection = self.next_section_editbox.text()
    pos_nextframe = self.next_frame_editbox.text()
    pos_ok_delete = self.ok_delete_editbox.text()
    pos_missingdata = self.missing_data_editbox.text()
    pos_assigncom = self.assign_comp_editbox.text()
    pos_clearelem = self.clear_elem_editbox.text()
    
    drag_duration = self.drag_duration_slider.value() / 100
    offset = self.offset_slider.value()
    wall_name = self.elem_name_editbox.text()
    
    # checkbox or editbox 비어있는 경우 break
    # while True:
    #     if (get_gbeam == False) & (get_gcol == False) & (get_ecol == False):
    #         msg = 'Nothing Checked!'
    #         msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
    #         self.status_browser.append(msg_colored)
    #         return
    #     elif (input_xlsx_path == '') | (len(result_xlsx_path) == 0):
    #         msg = 'Nothing Entered!'
    #         msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
    #         self.status_browser.append(msg_colored)
    #         return
    #     else: break
    
    # QThread 오브젝트 생성
    self.thread = QThread(parent=self) # Create a QThread object
    # InsertWorker 오브젝트 생성
    self.worker = MacroWorker(input_xlsx_path, start_or_end, macro_mode
                                  , pos_lefttop, pos_righttop, pos_leftbot
                                  , pos_rightbot, pos_p3dbar, pos_addcuts
                                  , pos_deletecuts, pos_ok, pos_nextsection
                                  , pos_nextframe, pos_ok_delete
                                  , pos_missingdata, pos_assigncom
                                  , pos_clearelem, drag_duration, offset
                                  , wall_name, time_start) # Create a worker object
    self.worker.moveToThread(self.thread) # Move worker to the thread
    
    # Connect signals and slots
    self.thread.started.connect(self.worker.macro_fn)
    self.worker.finished.connect(self.thread.quit)
    self.worker.finished.connect(self.worker.deleteLater)
    self.thread.finished.connect(self.thread.deleteLater)
    
    # Start the thread
    self.thread.start()
    
    # Enable/Disable the Button
    self.import_midas_btn.setEnabled(False)
    self.print_name_btn.setEnabled(False)
    self.convert_prop_btn.setEnabled(False)
    self.insert_force_btn.setEnabled(False)
    self.load_result_btn.setEnabled(False)
    self.design_wall_btn.setEnabled(False)
    self.print_pdf_btn.setEnabled(False)
    self.print_docx_btn.setEnabled(False)
    self.print_hwp_btn.setEnabled(False)
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_hwp_btn.setEnabled(True))
    
    # 완료 메세지 print
    self.worker.msg.connect(self.msg_fn)