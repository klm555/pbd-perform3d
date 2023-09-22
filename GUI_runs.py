import sys
import os
import time
import pandas as pd
import pickle
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QSettings, QCoreApplication, QThread, QObject, Qt
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5 import uic # ui 파일을 사용하기 위한 모듈

import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
from matplotlib.backends.backend_qt5agg import FigureCanvas as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure

from GUI_workers import *
import PBD_p3d as pbd
from GUI_second import BRSettingWindow

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
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    
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
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    
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
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    
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
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    
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
    
    BR_scale_factor = self.win_BR.BR_scale_factor_editbox.text().strip()
    
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

    # story_gap = int(story_gap)
    # max_shear = int(max_shear)
    story_gap = 2
    max_shear = 60000
    
    BR_scale_factor = float(BR_scale_factor)

    # QThread 오브젝트 생성
    self.thread = QThread(parent=self) # Create a QThread object
    # InsertWorker 오브젝트 생성
    self.worker = LoadWorker(input_xlsx_path, result_xlsx_path, wall_design_xlsx_path
                             , beam_design_xlsx_path, col_design_xlsx_path
                             , get_base_SF
                             , get_story_SF, get_IDR, get_BR, get_BSF
                             , get_E_BSF, get_CR, get_CSF, get_E_CSF
                             , get_WAS, get_WR, get_WSF, story_gap
                             , max_shear, BR_scale_factor, time_start) # Create a worker object
    self.worker.moveToThread(self.thread) # Move worker to the thread
    
    # Connect signals and slots
    self.thread.started.connect(self.worker.load_result_fn)
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
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    
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

# '''
#     if E_BSF == True:
#         # E_BSF 결과값 가져오기
#         node_map_z, node_map_list, element_map_list = result.E_BSF()

#         # V, M 크기에 따른 Color 지정
#         cmap_V = plt.get_cmap('Reds')
#         cmap_M = plt.get_cmap('YlOrBr')
        
#         # 층별 Loop
#         for i in node_map_z:   
#             # 해당 층에 해당하는 Nodes와 Elements만 Extract
#             node_map_list_extracted = node_map_list[node_map_list['V'] == i]
#             element_map_list_extracted = element_map_list[element_map_list['i-V'] == i]
#             element_map_list_extracted.reset_index(inplace=True, drop=True)
            
#             # Colorbar, 그래프 Coloring을 위한 설정
#             norm_V = plt.Normalize(vmin = element_map_list_extracted['DCR(V)'].min()\
#                                 , vmax = element_map_list_extracted['DCR(V)'].max())
#             cmap_V_elem = cmap_V(norm_V(element_map_list_extracted['DCR(V)']))
#             scalar_map_V = mpl.cm.ScalarMappable(norm_V, cmap_V)
            
#             norm_M = plt.Normalize(vmin = element_map_list_extracted['DCR(M)'].min()\
#                                 , vmax = element_map_list_extracted['DCR(M)'].max())
#             cmap_M_elem = cmap_M(norm_M(element_map_list_extracted['DCR(M)']))
#             scalar_map_M = mpl.cm.ScalarMappable(norm_M, cmap_M)
            
#             # E.Beam Contour 그래프 그리기
#             # V(전단)
#             # MplCanvas 생성
#             sc15 = pbd.ShowResult(self, width=6, height=3)
            
#             # Contour plot             
#             sc15.axes.scatter(node_map_list_extracted['H1'], node_map_list_extracted['H2'], color='k', s=1)
            
#             for idx, row in element_map_list_extracted.iterrows():
                
#                 element_map_x = [row['i-H1'], row['j-H1']]
#                 element_map_y = [row['i-H2'], row['j-H2']]
                
#                 sc15.axes.plot(element_map_x, element_map_y, c = cmap_V_elem[idx])
            
#             # Colorbar 만들기
#             sc15.fig.colorbar(scalar_map_V, shrink=0.7, label='DCR (V)')
        
#             # 기타
#             sc15.axes.set_axis_off()
#             sc15.axes.set_aspect('equal') # aspect 알아서 맞춤
#             sc15.axes.set_title(result.story_info['Story Name'][result.story_info['Height(mm)'] == i].iloc[0])

#             # toolbar 생성
#             toolbar15 = NavigationToolbar(sc15, self)

#             # layout에 toolbar, canvas 추가
#             layout.addWidget(toolbar15, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
#             layout.addWidget(sc15, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
#             row_count += 2
            
#             ## M(모멘트)     
#             # Graph    
#             sc16 = pbd.ShowResult(self, width=6, height=3)
            
#             # Contour plot
#             sc16.axes.scatter(node_map_list_extracted['H1'], node_map_list_extracted['H2'], color='k', s=1)
            
#             for idx, row in element_map_list_extracted.iterrows():
                
#                 element_map_x = [row['i-H1'], row['j-H1']]
#                 element_map_y = [row['i-H2'], row['j-H2']]
                
#                 sc16.axes.plot(element_map_x, element_map_y, c = cmap_M_elem[idx])
            
#             # Colorbar 만들기
#             sc16.fig.colorbar(scalar_map_M, shrink=0.7, label='DCR (M)')
        
#             # 기타
#             sc16.axes.set_axis_off()
#             sc16.axes.set_aspect('equal') # aspect 알아서 맞춤
#             sc16.axes.set_title(result.story_info['Story Name'][result.story_info['Height(mm)'] == i].iloc[0])

#             # toolbar 생성
#             toolbar16 = NavigationToolbar(sc16, self)

#             # layout에 toolbar, canvas 추가
#             layout.addWidget(toolbar16, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
#             layout.addWidget(sc16, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
#             row_count += 2
# '''

    # QThread 오브젝트 생성
    self.thread = QThread(parent=self) # Create a QThread object
    # PdfWorker 오브젝트 생성
    self.worker = DocxWorker(result_xlsx_path, story_gap, max_shear, get_base_SF
                             , get_story_SF, get_IDR, get_BR, get_BSF, get_E_BSF
                             , get_CR, get_CSF, get_E_CSF, get_WAS, get_WR
                             , get_WSF, project_name, bldg_name, story_gap
                             , max_shear, time_start) # Create a worker object
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
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    
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
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    
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
    
    # checkbox or editbox 비어있는 경우 break
    while True:
        if (get_cbeam == False) & (get_ecol == False) & (get_wall == False):
            msg = 'Nothing Checked!'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            return
        elif (project_name == '') | (bldg_name == ''):
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
    # PdfWorker 오브젝트 생성
    self.worker = PdfWorker(beam_design_xlsx_path, col_design_xlsx_path, wall_design_xlsx_path
                            , get_cbeam, get_ecol, get_wall, project_name, bldg_name, time_start) # Create a worker object
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
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.insert_force_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.load_result_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.design_wall_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_pdf_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_docx_btn.setEnabled(True))
    
    # 완료 메세지 print
    self.worker.msg.connect(self.msg_fn)