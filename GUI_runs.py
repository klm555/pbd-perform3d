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

#%% Tab1 - Import
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
    import_I_beam = self.import_I_beam_checkbox.isChecked()
    import_mass = self.import_mass_checkbox.isChecked()
    import_nodal_load = self.import_nodal_load_checkbox.isChecked()
    
    # 아무것도 check 안되어있는 경우 break
    if (base_SF == False) & (story_SF == False) & (IDR == False) & (BR == False)\
        & (E_BSF == False) & (E_CSF == False) & (WAS == False) & (WR == False)\
        & (WSF == False) & (WSF_each == False):
        self.status_browser.append('Nothing Checked')
        return
    
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
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    
    # 완료 메세지
    self.worker.msg.connect(self.msg_fn)

#%% Tab1 - Name
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
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    
    # 완료 메세지 print
    self.worker.msg.connect(self.msg_fn)
    
#%% Tab1 - Convert    
def run_worker3(self):
    # 시작 메세지
    time_start = time.time()
    self.status_browser.append('Running.....')
    
    # 변수 정리
    input_xlsx_path = self.data_conv_path_editbox.text()
    get_cbeam = self.convert_cbeam_checkbox.isChecked()
    get_gbeam = self.convert_gbeam_checkbox.isChecked()
    get_ebeam = self.convert_ebeam_checkbox.isChecked()
    get_gcol = self.convert_gcol_checkbox.isChecked()
    get_ecol = self.convert_ecol_checkbox.isChecked()
    get_wall = self.convert_wall_checkbox.isChecked()
    
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
    self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    
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
    result_xlsx_path = [i for i in result_xlsx_path if len(i) > 2]
    input_xlsx_path = self.data_conv_path_editbox.text()
    get_gbeam = self.insert_gbeam_checkbox.isChecked()
    get_gcol = self.insert_gcol_checkbox.isChecked()
    get_ecol = self.insert_ecol_checkbox.isChecked()
    
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
    # self.import_midas_btn.setEnabled(False)
    # self.print_name_btn.setEnabled(False)
    # self.convert_prop_btn.setEnabled(False)
    # self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    # self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    # self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    
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
    result_xlsx_path = [i for i in result_xlsx_path if len(i) > 2]
    input_xlsx_path = self.data_conv_path_editbox.text()
    # output_docx = self.output_docx_editbox.text()
    # bldg_name = self.bldg_name_editbox.text()
    story_gap = self.story_gap_editbox.text()
    max_shear = self.max_shear_editbox.text()
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
    # get_WSF_each = self.WSF_each_checkbox.isChecked()
    
    # 아무것도 check 안되어있는 경우 break
    # if (base_SF == False) & (story_SF == False) & (IDR == False) & (BR == False)\
    #     & (E_BSF == False) & (E_CSF == False) & (WAS == False) & (WR == False)\
    #     & (WSF == False) & (WSF_each == False):
    #     self.status_browser.append('Nothing Checked')
    #     return

    story_gap = int(story_gap)
    max_shear = int(max_shear)

    # QThread 오브젝트 생성
    self.thread = QThread(parent=self) # Create a QThread object
    # InsertWorker 오브젝트 생성
    self.worker = LoadWorker(input_xlsx_path, result_xlsx_path, get_base_SF
                                 , get_story_SF, get_IDR, get_BR, get_BSF
                                 , get_E_BSF, get_CR, get_CSF, get_E_CSF
                                 , get_WAS, get_WR, get_WSF, story_gap
                                 , max_shear, time_start) # Create a worker object
    self.worker.moveToThread(self.thread) # Move worker to the thread
    
    # Connect signals and slots
    self.thread.started.connect(self.worker.load_result_fn)
    self.worker.finished.connect(self.thread.quit)
    self.worker.finished.connect(self.worker.deleteLater)
    self.thread.finished.connect(self.thread.deleteLater)
    
    # Start the thread
    self.thread.start()
    
    # Send Signals(incl. result data) to plot_display function
    self.worker.result_data.connect(self.plot_display)
    
    # Enable/Disable the Button
    # self.import_midas_btn.setEnabled(False)
    # self.print_name_btn.setEnabled(False)
    # self.convert_prop_btn.setEnabled(False)
    # self.thread.finished.connect(lambda: self.import_midas_btn.setEnabled(True))
    # self.thread.finished.connect(lambda: self.print_name_btn.setEnabled(True))
    # self.thread.finished.connect(lambda: self.convert_prop_btn.setEnabled(True))
    
    # 실행 시간 계산
    time_end = time.time()
    time_run = (time_end-time_start)/60
    
    # 완료 메세지 print
    msg = 'Completed!' + '  (total time = %0.3f min)' %(time_run)
    msg_colored = '<span style=\" color: #0000ff;\">%s</span>' % msg
    self.status_browser.append(msg_colored)    
    
    # 실패할 경우, 그냥 msg_fn 이용해서 오류 메세지 내보내기
    # (단, 여기서 오류가 발생하는 경우, loadworker에서 데이터 처리하는 과정에서의 오류밖에 캐치 안됨)
    self.worker.msg.connect(self.msg_fn)

#%% Tab3 - Load
def run_worker6(self):
    # 시작 메세지
    time_start = time.time()
    self.status_browser.append('Running.....')
    
    # 변수 정리
    result_xlsx_path = self.result_path_editbox.text()
    result_xlsx_path = result_xlsx_path.split('"')
    result_xlsx_path = [i for i in result_xlsx_path if len(i) > 2]
    input_xlsx_path = self.data_conv_path_editbox.text()
    output_docx = self.output_docx_editbox.text()
    bldg_name = self.bldg_name_editbox.text()
    story_gap = self.story_gap_editbox.text()
    max_shear = self.max_shear_editbox.text()
    base_SF = self.base_SF_checkbox.isChecked()
    story_SF = self.story_SF_checkbox.isChecked()
    IDR = self.IDR_checkbox.isChecked()
    BR = self.BR_checkbox.isChecked()
    BSF = self.BSF_checkbox.isChecked()
    E_BSF = self.E_BSF_checkbox.isChecked()
    CR = self.CR_checkbox.isChecked()
    CSF = self.CSF_checkbox.isChecked()
    E_CSF = self.E_CSF_checkbox.isChecked()
    WAS = self.WAS_checkbox.isChecked()
    WR = self.WR_checkbox.isChecked()
    WSF = self.WSF_checkbox.isChecked()
    WSF_each = self.WSF_each_checkbox.isChecked()
    
    # 아무것도 check 안되어있는 경우 break
    # if (base_SF == False) & (story_SF == False) & (IDR == False) & (BR == False)\
    #     & (E_BSF == False) & (E_CSF == False) & (WAS == False) & (WR == False)\
    #     & (WSF == False) & (WSF_each == False):
    #     self.status_browser.append('Nothing Checked')
    #     return

    story_gap = int(story_gap)
    max_shear = int(max_shear)

    # QScrollAreadp container QWidget 생성
    self.plot_display_area.setWidgetResizable(True)
    container = QWidget()
    self.plot_display_area.setWidget(container)
    # grid layout
    layout = QGridLayout(container)
    # 후처리 결과 object 생성
    result = pbd.PrintResult(input_xlsx_path, result_xlsx_path, bldg_name, story_gap, max_shear)
    row_count = 0

    # try:

    if base_SF == True:
        # base_SF 결과값 가져오기
        base_SF_result = result.base_SF()

        # 결과값 classify & assign
        base_shear_H1 = base_SF_result[0]
        base_shear_H2 = base_SF_result[1]
        DE_load_name_list = base_SF_result[2]
        MCE_load_name_list = base_SF_result[3]
        base_SF_markers = base_SF_result[4]

        # base_SF 그래프 그리기
        # DE Plot
        if len(DE_load_name_list) != 0:
            # H1_DE
            # MplCanvas 생성
            sc1 = pbd.ShowResult(self, width=6, height=4)

            # MplCanvas에 그래프 그리기
            sc1.axes.bar(range(len(DE_load_name_list)), base_shear_H1.iloc[0, 0:len(DE_load_name_list)]\
                        , color='darkblue', edgecolor='k', label = 'Max. Base Shear')
            sc1.axes.axhline(y= base_shear_H1.iloc[0, 0:len(DE_load_name_list)].mean(), color='r', linestyle='-', label='Average')
            
            sc1.axes.set_ylim(0, max_shear)
            sc1.axes.set_xticks(range(14), range(1,15))

            sc1.axes.set_xlabel('Ground Motion No.')
            sc1.axes.set_ylabel('Base Shear(kN)')
            sc1.axes.legend(loc = 2)
            sc1.axes.set_title('X DE')
            
            # toolbar 생성
            toolbar1 = NavigationToolbar(sc1, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar1, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc1, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # H2_DE
            sc2 = pbd.ShowResult(self, width=6, height=4)

            sc2.axes.bar(range(len(DE_load_name_list)), base_shear_H2.iloc[0, 0:len(DE_load_name_list)]\
                        , color='darkblue', edgecolor='k', label = 'Max. Base Shear')
            sc2.axes.axhline(y= base_shear_H2.iloc[0, 0:len(DE_load_name_list)].mean(), color='r', linestyle='-', label='Average')
            
            sc2.axes.set_ylim(0, max_shear)
            sc2.axes.set_xticks(range(14), range(1,15))

            sc2.axes.set_xlabel('Ground Motion No.')
            sc2.axes.set_ylabel('Base Shear(kN)')
            sc2.axes.legend(loc = 2)
            sc2.axes.set_title('Y DE')

            toolbar2 = NavigationToolbar(sc2, self)

            layout.addWidget(toolbar2, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc2, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

        # MCE Plot
        if len(MCE_load_name_list) != 0:

            # H1_MCE
            sc3 = pbd.ShowResult(self, width=6, height=4)
            
            sc3.axes.bar(range(len(MCE_load_name_list)), base_shear_H1\
                    .iloc[0, len(DE_load_name_list):len(DE_load_name_list)+len(MCE_load_name_list)]\
                    , color='darkblue', edgecolor='k', label = 'Max. Base Shear')
            
            sc3.axes.axhline(y= base_shear_H1.iloc[0, len(DE_load_name_list):len(DE_load_name_list)+len(MCE_load_name_list)]\
                        .mean(), color='r', linestyle='-', label='Average')
            
            sc3.axes.set_ylim(0, max_shear)
            sc3.axes.set_xticks(range(14), range(1,15))
            
            sc3.axes.set_xlabel('Ground Motion No.')
            sc3.axes.set_ylabel('Base Shear(kN)')
            sc3.axes.legend(loc = 2)
            sc3.axes.set_title('X MCE')
            
            toolbar3 = NavigationToolbar(sc3, self)

            layout.addWidget(toolbar3, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc3, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # H2_MCE
            sc4 = pbd.ShowResult(self, width=6, height=4)
            
            sc4.axes.bar(range(len(MCE_load_name_list)), base_shear_H2\
                    .iloc[0, len(DE_load_name_list):len(DE_load_name_list)+len(MCE_load_name_list)]\
                    , color='darkblue', edgecolor='k', label = 'Max. Base Shear')
            
            sc4.axes.axhline(y= base_shear_H2.iloc[0, len(DE_load_name_list):len(DE_load_name_list)+len(MCE_load_name_list)]\
                        .mean(), color='r', linestyle='-', label='Average')
            
            sc4.axes.set_ylim(0, max_shear)
            sc4.axes.set_xticks(range(14), range(1,15))
            
            sc4.axes.set_xlabel('Ground Motion No.')
            sc4.axes.set_ylabel('Base Shear(kN)')
            sc4.axes.legend(loc = 2)
            sc4.axes.set_title('Y MCE')

            toolbar4 = NavigationToolbar(sc4, self)

            layout.addWidget(toolbar4, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc4, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

    if story_SF == True:
        # base_SF 결과값 가져오기
        story_SF_result = result.story_SF()

        # 결과값 classify & assign
        shear_force_H1_max = story_SF_result[0]
        shear_force_H2_max = story_SF_result[1]
        DE_load_name_list = story_SF_result[2]
        MCE_load_name_list = story_SF_result[3]
        base_SF_markers = story_SF_result[4]

        # base_SF 그래프 그리기
        # DE Plot
        if len(DE_load_name_list) != 0:
            # H1_DE
            # MplCanvas 생성
            sc5 = pbd.ShowResult(self, width=6, height=5)

            # MplCanvas에 그래프 그리기                    
            # 지진파별 plot
            for i in range(len(DE_load_name_list)):
                sc5.axes.plot(shear_force_H1_max.iloc[:,i], range(shear_force_H1_max.shape[0]), label=DE_load_name_list[i], linewidth=0.7)
                
            # 평균 plot
            sc5.axes.plot(shear_force_H1_max.iloc[:,0:len(DE_load_name_list)]\
                    .mean(axis=1), range(shear_force_H1_max.shape[0]), color='k', label='Average', linewidth=2)
            
            sc5.axes.set_xlim(0, max_shear)
            sc5.axes.set_yticks(range(shear_force_H1_max.shape[0])[::story_gap], shear_force_H1_max.index[::story_gap], fontsize=8.5)
            
            # 기타
            sc5.axes.grid(linestyle='-.')
            sc5.axes.set_xlabel('Story Shear(kN)')
            sc5.axes.set_ylabel('Story')
            sc5.axes.legend(loc=1, fontsize=8)
            sc5.axes.set_title('X DE')
            
            # toolbar 생성
            toolbar5 = NavigationToolbar(sc5, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar5, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc5, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # H2_DE
            sc6 = pbd.ShowResult(self, width=6, height=5)

            for i in range(len(DE_load_name_list)):
                sc5.axes.plot(shear_force_H2_max.iloc[:,i], range(shear_force_H2_max.shape[0]), label=DE_load_name_list[i], linewidth=0.7)
            
            sc6.axes.plot(shear_force_H2_max.iloc[:,0:len(DE_load_name_list)]\
                    .mean(axis=1), range(shear_force_H2_max.shape[0]), color='k', label='Average', linewidth=2)
            
            sc6.axes.set_xlim(0, max_shear)
            sc6.axes.set_yticks(range(shear_force_H2_max.shape[0])[::story_gap], shear_force_H2_max.index[::story_gap], fontsize=8.5)
        
            sc6.axes.grid(linestyle='-.')
            sc6.axes.set_xlabel('Story Shear(kN)')
            sc6.axes.set_ylabel('Story')
            sc6.axes.legend(loc=1, fontsize=8)
            sc6.axes.set_title('Y DE')

            toolbar6 = NavigationToolbar(sc6, self)

            layout.addWidget(toolbar6, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc6, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

        # MCE Plot
        if len(MCE_load_name_list) != 0:

            # H1_MCE
            sc7 = pbd.ShowResult(self, width=6, height=5)

            for i in range(len(MCE_load_name_list)):
                sc7.axes.plot(shear_force_H1_max.iloc[:,i+len(DE_load_name_list)], range(shear_force_H1_max.shape[0]), label=MCE_load_name_list[i], linewidth=0.7)
            sc7.axes.plot(shear_force_H1_max.iloc[:,len(DE_load_name_list)\
                                                    :len(DE_load_name_list)+len(MCE_load_name_list)]\
                            .mean(axis=1), range(shear_force_H1_max.shape[0]), color='k', label='Average', linewidth=2)
            
            sc7.axes.set_xlim(0, max_shear)
            sc7.axes.set_yticks(range(shear_force_H1_max.shape[0])[::story_gap], shear_force_H1_max.index[::story_gap], fontsize=8.5)
        
            sc7.axes.grid(linestyle='-.')
            sc7.axes.set_xlabel('Story Shear(kN)')
            sc7.axes.set_ylabel('Story')
            sc7.axes.legend(loc=1, fontsize=8)
            sc7.axes.set_title('X MCE')

            toolbar7 = NavigationToolbar(sc7, self)

            layout.addWidget(toolbar7, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc7, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # H2_MCE
            sc8 = pbd.ShowResult(self, width=6, height=5)

            for i in range(len(MCE_load_name_list)):
                sc8.axes.plot(shear_force_H2_max.iloc[:,i+len(DE_load_name_list)], range(shear_force_H2_max.shape[0]), label=MCE_load_name_list[i], linewidth=0.7)
            sc8.axes.plot(shear_force_H2_max.iloc[:,len(DE_load_name_list)\
                                                    :len(DE_load_name_list)+len(MCE_load_name_list)]\
                            .mean(axis=1), range(shear_force_H2_max.shape[0]), color='k', label='Average', linewidth=2)
            
            sc8.axes.set_xlim(0, max_shear)
            sc8.axes.set_yticks(range(shear_force_H2_max.shape[0])[::story_gap], shear_force_H2_max.index[::story_gap], fontsize=8.5)
        
            sc8.axes.grid(linestyle='-.')
            sc8.axes.set_xlabel('Story Shear(kN)')
            sc8.axes.set_ylabel('Story')
            sc8.axes.legend(loc=1, fontsize=8)
            sc8.axes.set_title('Y MCE')

            toolbar8 = NavigationToolbar(sc8, self)

            layout.addWidget(toolbar8, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc8, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2
            
    if IDR == True:
        # base_SF 결과값 가져오기
        IDR_result = result.IDR()

        # 결과값 classify & assign
        IDR_result_each = IDR_result[0]
        IDR_result_avg = IDR_result[1]
        story_name_window_reordered = IDR_result[2]

        cri_DE = 0.015
        cri_MCE = 0.02

            # IDR 그래프 그리기
        count_x = 0
        count_y = 2
        count_avg = 0

        # DE Plot
        if len(result.DE_load_name_list) != 0:
            # H1_DE
            # MplCanvas 생성
            sc9 = pbd.ShowResult(self, width=5, height=7)

            # MplCanvas에 그래프 그리기                    
            # 지진파별 plot
            for load_name in result.DE_load_name_list:
                sc9.axes.plot(IDR_result_each[count_x].iloc[:,-1], IDR_result_each[count_x].iloc[:,0]
                                , label='{}'.format(load_name), linewidth=0.7)
                sc9.axes.plot(IDR_result_each[count_x+1].iloc[:,-1], IDR_result_each[count_x].iloc[:,0]
                                , linewidth=0.7)
                count_x += 2
                
            # 평균 plot
            sc9.axes.plot(IDR_result_avg[count_avg].iloc[:,0], story_name_window_reordered, color='k', label='Average', linewidth=2)
            sc9.axes.plot(IDR_result_avg[count_avg].iloc[:,1], story_name_window_reordered, color='k', linewidth=2)
            
            # reference line 그려서 허용치 나타내기
            sc9.axes.axvline(x=-cri_DE, color='r', linestyle='--', label='LS')
            sc9.axes.axvline(x=cri_DE, color='r', linestyle='--')

            sc9.axes.set_xlim(-0.025, 0.025)
            sc9.axes.set_yticks(story_name_window_reordered[::story_gap], story_name_window_reordered[::story_gap])
            
            # 기타
            sc9.axes.grid(linestyle='-.')
            sc9.axes.set_xlabel('Interstory Drift Ratios(m/m)')
            sc9.axes.set_ylabel('Story')
            sc9.axes.legend(loc=4, fontsize=8)
            sc9.axes.set_title('X DE')
            
            # toolbar 생성
            toolbar9 = NavigationToolbar(sc9, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar9, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc9, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # H2_DE
            # MplCanvas 생성
            sc10 = pbd.ShowResult(self, width=5, height=7)

            # MplCanvas에 그래프 그리기                    
            # 지진파별 plot
            for load_name in result.DE_load_name_list:
                sc10.axes.plot(IDR_result_each[count_y].iloc[:,-1], IDR_result_each[count_y].iloc[:,0]
                                , label='{}'.format(load_name), linewidth=0.7)
                sc10.axes.plot(IDR_result_each[count_y+1].iloc[:,-1], IDR_result_each[count_y].iloc[:,0]
                                , linewidth=0.7)
                count_y += 2
                
            # 평균 plot
            sc10.axes.plot(IDR_result_avg[count_avg].iloc[:,2], story_name_window_reordered, color='k', label='Average', linewidth=2)
            sc10.axes.plot(IDR_result_avg[count_avg].iloc[:,3], story_name_window_reordered, color='k', linewidth=2)
            count_avg += 1
            
            # reference line 그려서 허용치 나타내기
            sc10.axes.axvline(x=-cri_DE, color='r', linestyle='--', label='LS')
            sc10.axes.axvline(x=cri_DE, color='r', linestyle='--')

            sc10.axes.set_xlim(-0.025, 0.025)
            sc10.axes.set_yticks(story_name_window_reordered[::story_gap], story_name_window_reordered[::story_gap])
            
            # 기타
            sc10.axes.grid(linestyle='-.')
            sc10.axes.set_xlabel('Interstory Drift Ratios(m/m)')
            sc10.axes.set_ylabel('Story')
            sc10.axes.legend(loc=4, fontsize=8)
            sc10.axes.set_title('Y DE')
            
            # toolbar 생성
            toolbar10 = NavigationToolbar(sc10, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar10, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc10, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

        # DE Plot
        if len(result.MCE_load_name_list) != 0:
            # H1_MCE
            # MplCanvas 생성
            sc11 = pbd.ShowResult(self, width=5, height=7)

            # MplCanvas에 그래프 그리기                    
            # 지진파별 plot
            for load_name in result.MCE_load_name_list:
                sc11.axes.plot(IDR_result_each[count_x].iloc[:,-1], IDR_result_each[count_x].iloc[:,0]
                                , label='{}'.format(load_name), linewidth=0.7)
                sc11.axes.plot(IDR_result_each[count_x+1].iloc[:,-1], IDR_result_each[count_x].iloc[:,0]
                                , linewidth=0.7)
                count_x += 2
                
            # 평균 plot
            sc11.axes.plot(IDR_result_avg[count_avg].iloc[:,0], story_name_window_reordered, color='k', label='Average', linewidth=2)
            sc11.axes.plot(IDR_result_avg[count_avg].iloc[:,1], story_name_window_reordered, color='k', linewidth=2)
            
            # reference line 그려서 허용치 나타내기
            sc11.axes.axvline(x=-cri_MCE, color='r', linestyle='--', label='LS')
            sc11.axes.axvline(x=cri_MCE, color='r', linestyle='--')

            sc11.axes.set_xlim(-0.025, 0.025)
            sc11.axes.set_yticks(story_name_window_reordered[::story_gap], story_name_window_reordered[::story_gap])
            
            # 기타
            sc11.axes.grid(linestyle='-.')
            sc11.axes.set_xlabel('Interstory Drift Ratios(m/m)')
            sc11.axes.set_ylabel('Story')
            sc11.axes.legend(loc=4, fontsize=8)
            sc11.axes.set_title('X MCE')
            
            # toolbar 생성
            toolbar11 = NavigationToolbar(sc11, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar11, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc11, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # H2_MCE
            # MplCanvas 생성
            sc12 = pbd.ShowResult(self, width=5, height=7)

            # MplCanvas에 그래프 그리기                    
            # 지진파별 plot
            for load_name in result.MCE_load_name_list:
                sc12.axes.plot(IDR_result_each[count_y].iloc[:,-1], IDR_result_each[count_y].iloc[:,0]
                                , label='{}'.format(load_name), linewidth=0.7)
                sc12.axes.plot(IDR_result_each[count_y+1].iloc[:,-1], IDR_result_each[count_y].iloc[:,0]
                                , linewidth=0.7)
                count_y += 2
                
            # 평균 plot
            sc12.axes.plot(IDR_result_avg[count_avg].iloc[:,2], story_name_window_reordered, color='k', label='Average', linewidth=2)
            sc12.axes.plot(IDR_result_avg[count_avg].iloc[:,3], story_name_window_reordered, color='k', linewidth=2)
            count_avg += 1
            
            # reference line 그려서 허용치 나타내기
            sc12.axes.axvline(x=-cri_MCE, color='r', linestyle='--', label='LS')
            sc12.axes.axvline(x=cri_MCE, color='r', linestyle='--')

            sc12.axes.set_xlim(-0.025, 0.025)
            sc12.axes.set_yticks(story_name_window_reordered[::story_gap], story_name_window_reordered[::story_gap])
            
            # 기타
            sc12.axes.grid(linestyle='-.')
            sc12.axes.set_xlabel('Interstory Drift Ratios(m/m)')
            sc12.axes.set_ylabel('Story')
            sc12.axes.legend(loc=4, fontsize=8)
            sc12.axes.set_title('Y MCE')
            
            # toolbar 생성
            toolbar12 = NavigationToolbar(sc12, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar12, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc12, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2
                
    if BR == True:
        # BR 결과값 가져오기
        BR_result = result.BR(c_beam_group = 'C.Beam')

        # 결과값 classify & assign
        beam_rot_data = BR_result[0]
        story_info = BR_result[1]

        DCR_criteria = 1
        xlim = 3

        # BR 그래프 그리기
        count = 0

        # DE Plot
        if len(result.DE_load_name_list) != 0:
            # MplCanvas 생성
            sc13 = pbd.ShowResult(self, width=5, height=6)

            # DCR plot                
            sc13.axes.scatter(beam_rot_data[count]['DE Max avg'], beam_rot_data[count].loc[:,'V'], color='k', s=1)
            sc13.axes.scatter(beam_rot_data[count]['DE Min avg'], beam_rot_data[count].loc[:,'V'], color='k', s=1)

            # 허용치(DCR) 기준선
            sc13.axes.axvline(x=DCR_criteria, color='r', linestyle='--')

            sc13.axes.set_xlim(0, xlim)
            sc13.axes.set_yticks(story_info['Height(mm)'][::-story_gap], story_info['Story Name'][::-story_gap])

            # 기타
            sc13.axes.grid(linestyle='-.')
            sc13.axes.set_xlabel('D/C Ratios')
            sc13.axes.set_ylabel('Story')
            sc13.axes.set_title('Beam Rotation (DE)')

            # 기준 넘는 점 확인
            error_beam_DE = beam_rot_data[count][['Element Name', 'Property Name', 'Story Name', 'DE Max avg', 'DE Min avg']]\
                            [(beam_rot_data[count]['DE Max avg'] >= DCR_criteria) | (beam_rot_data[count]['DE Min avg'] >= DCR_criteria)]
            
            self.status_browser.append(str(error_beam_DE['Property Name']))
            
            count += 1

            # toolbar 생성
            toolbar13 = NavigationToolbar(sc13, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar13, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc13, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2
            
        if len(result.MCE_load_name_list) != 0:
            # MplCanvas 생성
            sc14 = pbd.ShowResult(self, width=5, height=6)

            # DCR plot                
            sc14.axes.scatter(beam_rot_data[count]['MCE Max avg'], beam_rot_data[count].loc[:,'V'], color='k', s=1)
            sc14.axes.scatter(beam_rot_data[count]['MCE Min avg'], beam_rot_data[count].loc[:,'V'], color='k', s=1)

            # 허용치(DCR) 기준선
            sc14.axes.axvline(x=DCR_criteria, color='r', linestyle='--')

            sc14.axes.set_xlim(0, xlim)
            sc14.axes.set_yticks(story_info['Height(mm)'][::-story_gap], story_info['Story Name'][::-story_gap])

            # 기타
            sc14.axes.grid(linestyle='-.')
            sc14.axes.set_xlabel('D/C Ratios')
            sc14.axes.set_ylabel('Story')
            sc14.axes.set_title('Beam Rotation (MCE)')

            # 기준 넘는 점 확인
            error_beam_MCE = beam_rot_data[count][['Element Name', 'Property Name', 'Story Name', 'MCE Max avg', 'MCE Min avg']]\
                            [(beam_rot_data[count]['MCE Max avg'] >= DCR_criteria) | (beam_rot_data[count]['MCE Min avg'] >= DCR_criteria)]
            
            self.status_browser.append(str(error_beam_MCE['Property Name']))
            
            # toolbar 생성
            toolbar14 = NavigationToolbar(sc14, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar14, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc14, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

    if E_BSF == True:
        # E_BSF 결과값 가져오기
        node_map_z, node_map_list, element_map_list = result.E_BSF()

        # V, M 크기에 따른 Color 지정
        cmap_V = plt.get_cmap('Reds')
        cmap_M = plt.get_cmap('YlOrBr')
        
        # 층별 Loop
        for i in node_map_z:   
            # 해당 층에 해당하는 Nodes와 Elements만 Extract
            node_map_list_extracted = node_map_list[node_map_list['V'] == i]
            element_map_list_extracted = element_map_list[element_map_list['i-V'] == i]
            element_map_list_extracted.reset_index(inplace=True, drop=True)
            
            # Colorbar, 그래프 Coloring을 위한 설정
            norm_V = plt.Normalize(vmin = element_map_list_extracted['DCR(V)'].min()\
                                , vmax = element_map_list_extracted['DCR(V)'].max())
            cmap_V_elem = cmap_V(norm_V(element_map_list_extracted['DCR(V)']))
            scalar_map_V = mpl.cm.ScalarMappable(norm_V, cmap_V)
            
            norm_M = plt.Normalize(vmin = element_map_list_extracted['DCR(M)'].min()\
                                , vmax = element_map_list_extracted['DCR(M)'].max())
            cmap_M_elem = cmap_M(norm_M(element_map_list_extracted['DCR(M)']))
            scalar_map_M = mpl.cm.ScalarMappable(norm_M, cmap_M)
            
            # E.Beam Contour 그래프 그리기
            # V(전단)
            # MplCanvas 생성
            sc15 = pbd.ShowResult(self, width=6, height=3)
            
            # Contour plot             
            sc15.axes.scatter(node_map_list_extracted['H1'], node_map_list_extracted['H2'], color='k', s=1)
            
            for idx, row in element_map_list_extracted.iterrows():
                
                element_map_x = [row['i-H1'], row['j-H1']]
                element_map_y = [row['i-H2'], row['j-H2']]
                
                sc15.axes.plot(element_map_x, element_map_y, c = cmap_V_elem[idx])
            
            # Colorbar 만들기
            sc15.fig.colorbar(scalar_map_V, shrink=0.7, label='DCR (V)')
        
            # 기타
            sc15.axes.set_axis_off()
            sc15.axes.set_aspect('equal') # aspect 알아서 맞춤
            sc15.axes.set_title(result.story_info['Story Name'][result.story_info['Height(mm)'] == i].iloc[0])

            # toolbar 생성
            toolbar15 = NavigationToolbar(sc15, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar15, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc15, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2
            
            ## M(모멘트)     
            # Graph    
            sc16 = pbd.ShowResult(self, width=6, height=3)
            
            # Contour plot
            sc16.axes.scatter(node_map_list_extracted['H1'], node_map_list_extracted['H2'], color='k', s=1)
            
            for idx, row in element_map_list_extracted.iterrows():
                
                element_map_x = [row['i-H1'], row['j-H1']]
                element_map_y = [row['i-H2'], row['j-H2']]
                
                sc16.axes.plot(element_map_x, element_map_y, c = cmap_M_elem[idx])
            
            # Colorbar 만들기
            sc16.fig.colorbar(scalar_map_M, shrink=0.7, label='DCR (M)')
        
            # 기타
            sc16.axes.set_axis_off()
            sc16.axes.set_aspect('equal') # aspect 알아서 맞춤
            sc16.axes.set_title(result.story_info['Story Name'][result.story_info['Height(mm)'] == i].iloc[0])

            # toolbar 생성
            toolbar16 = NavigationToolbar(sc16, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar16, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc16, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

    if E_CSF == True:
        # E_CSF 결과값 가져오기
        result.E_CSF(bldg_name=bldg_name)

    if WAS == True:
        # WAS 결과값 가져오기
        AS_output, story_info = result.WAS()

        max_criteria = 0.04
        min_criteria = -0.002

        # WAS 그래프 그리기
        # DE Plot
        if len(result.DE_load_name_list) != 0:
            # DE_1
            # MplCanvas 생성
            sc17 = pbd.ShowResult(self, width=5, height=4)

            # WAS plot
            sc17.axes.scatter(AS_output['DE_min_avg'], AS_output['Z(mm)'], color='r', s=5)
            sc17.axes.scatter(AS_output['DE_max_avg'], AS_output['Z(mm)'], color='k', s=5)

            # 허용치 기준선
            sc17.axes.axvline(x=min_criteria, color='r', linestyle='--')
            sc17.axes.axvline(x=max_criteria, color='r', linestyle='--')

            sc17.axes.set_xlim(-0.003, 0)
            sc17.axes.set_yticks(story_info['Height(mm)'][::-story_gap], story_info['Story Name'][::-story_gap])

            # 기타
            sc17.axes.grid(linestyle='-.')
            sc17.axes.set_xlabel('Axial Strain (m/m)')
            sc17.axes.set_ylabel('Story')
            sc17.axes.set_title('DE (Compressive)')

            # toolbar 생성
            toolbar17 = NavigationToolbar(sc17, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar17, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc17, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # DE_2
            # MplCanvas 생성
            sc18 = pbd.ShowResult(self, width=5, height=4)

            # WAS plot
            sc18.axes.scatter(AS_output['DE_min_avg'], AS_output['Z(mm)'], color='r', s=5)
            sc18.axes.scatter(AS_output['DE_max_avg'], AS_output['Z(mm)'], color='k', s=5)

            # 허용치 기준선
            sc18.axes.axvline(x=min_criteria, color='r', linestyle='--')
            sc18.axes.axvline(x=max_criteria, color='r', linestyle='--')

            sc18.axes.set_xlim(0, 0.013)
            sc18.axes.set_yticks(story_info['Height(mm)'][::-story_gap], story_info['Story Name'][::-story_gap])

            # 기타
            sc18.axes.grid(linestyle='-.')
            sc18.axes.set_xlabel('Axial Strain (m/m)')
            sc18.axes.set_ylabel('Story')
            sc18.axes.set_title('DE (Tensile)')

            # toolbar 생성
            toolbar18 = NavigationToolbar(sc18, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar18, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc18, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2
            
            # 기준 넘는 점 확인
            error_coord_DE = AS_output[(AS_output['DE_max_avg'] >= max_criteria)
                                        | (AS_output['DE_min_avg'] <= min_criteria)]
            
        # MCE Plot
        if len(result.MCE_load_name_list) != 0:
            # MCE_1
            # MplCanvas 생성
            sc19 = pbd.ShowResult(self, width=5, height=4)

            # WAS plot
            sc19.axes.scatter(AS_output['MCE_min_avg'], AS_output['Z(mm)'], color='r', s=5)
            sc19.axes.scatter(AS_output['MCE_max_avg'], AS_output['Z(mm)'], color='k', s=5)

            # 허용치 기준선
            sc19.axes.axvline(x=min_criteria, color='r', linestyle='--')
            sc19.axes.axvline(x=max_criteria, color='r', linestyle='--')

            sc19.axes.set_xlim(-0.003, 0)
            sc19.axes.set_yticks(story_info['Height(mm)'][::-story_gap], story_info['Story Name'][::-story_gap])

            # 기타
            sc19.axes.grid(linestyle='-.')
            sc19.axes.set_xlabel('Axial Strain (m/m)')
            sc19.axes.set_ylabel('Story')
            sc19.axes.set_title('MCE (Compressive)')

            # toolbar 생성
            toolbar19 = NavigationToolbar(sc19, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar19, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc19, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # MCE_2
            # MplCanvas 생성
            sc20 = pbd.ShowResult(self, width=5, height=4)

            # WAS plot
            sc20.axes.scatter(AS_output['MCE_min_avg'], AS_output['Z(mm)'], color='r', s=5)
            sc20.axes.scatter(AS_output['MCE_max_avg'], AS_output['Z(mm)'], color='k', s=5)

            # 허용치 기준선
            sc20.axes.axvline(x=min_criteria, color='r', linestyle='--')
            sc20.axes.axvline(x=max_criteria, color='r', linestyle='--')

            sc20.axes.set_xlim(0, 0.013)
            sc20.axes.set_yticks(story_info['Height(mm)'][::-story_gap], story_info['Story Name'][::-story_gap])

            # 기타
            sc20.axes.grid(linestyle='-.')
            sc20.axes.set_xlabel('Axial Strain (m/m)')
            sc20.axes.set_ylabel('Story')
            sc20.axes.set_title('MCE (Tensile)')

            # toolbar 생성
            toolbar20 = NavigationToolbar(sc20, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar20, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc20, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # 기준 넘는 점 확인
            error_coord_MCE = AS_output[(AS_output['MCE_max_avg'] >= max_criteria)
                                        | (AS_output['MCE_min_avg'] <= min_criteria)]

    if WR == True:
        # WR 결과값 가져오기
        SWR_avg_total = result.WR()

        DCR_criteria = 1
        xlim = 3

        # WR 그래프 그리기
        # DE Plot
        if len(result.DE_load_name_list) != 0:
            # DE
            # MplCanvas 생성
            sc21 = pbd.ShowResult(self, width=5, height=6)

            # WR plot
            sc21.axes.scatter(SWR_avg_total['DCR_DE_min'], SWR_avg_total['Height'], color='k', s=1)
            sc21.axes.scatter(SWR_avg_total['DCR_DE_max'], SWR_avg_total['Height'], color='k', s=1)

            # 허용치 기준선
            sc21.axes.axvline(x = DCR_criteria, color='r', linestyle='--')

            sc21.axes.set_xlim(0, xlim)
            sc21.axes.set_yticks(result.story_info['Height(mm)'][::-story_gap], result.story_info['Story Name'][::-story_gap])
            
            # 기타
            sc21.axes.grid(linestyle='-.')
            sc21.axes.set_xlabel('D/C Ratios')
            sc21.axes.set_ylabel('Story')
            sc21.axes.set_title('Wall Rotation (DE)')

            # toolbar 생성
            toolbar21 = NavigationToolbar(sc21, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar21, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc21, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # 기준 넘는 벽체 확인
            error_wall_DE = SWR_avg_total[['gage_name', 'DCR_DE_min', 'DCR_DE_max']]
            [(SWR_avg_total['DCR_DE_min']>= DCR_criteria) | (SWR_avg_total['DCR_DE_max']>= DCR_criteria)]

        if len(result.MCE_load_name_list) != 0:
            # MCE
            # MplCanvas 생성
            sc22 = pbd.ShowResult(self, width=5, height=6)

            # WR plot
            sc22.axes.scatter(SWR_avg_total['DCR_MCE_min'], SWR_avg_total['Height'], color='k', s=1)
            sc22.axes.scatter(SWR_avg_total['DCR_MCE_max'], SWR_avg_total['Height'], color='k', s=1)

            # 허용치 기준선
            sc22.axes.axvline(x = DCR_criteria, color='r', linestyle='--')

            sc22.axes.set_xlim(0, xlim)
            sc22.axes.set_yticks(result.story_info['Height(mm)'][::-story_gap], result.story_info['Story Name'][::-story_gap])

            # 기타
            sc22.axes.grid(linestyle='-.')
            sc22.axes.set_xlabel('D/C Ratios')
            sc22.axes.set_ylabel('Story')
            sc22.axes.set_title('Wall Rotation (MCE)')

            # toolbar 생성
            toolbar22 = NavigationToolbar(sc22, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar22, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc22, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # 기준 넘는 벽체 확인
            error_wall_MCE = SWR_avg_total[['gage_name', 'DCR_MCE_min', 'DCR_MCE_max']]
            [(SWR_avg_total['DCR_MCE_min']>= DCR_criteria) | (SWR_avg_total['DCR_MCE_max']>= DCR_criteria)]

    if WSF == True:
        # WSF 결과값 가져오기
        wall_result_output = result.WSF()

        DCR_criteria = 1
        xlim = 3

        # WSF 그래프 그리기
        # DE Plot
        if len(result.DE_load_name_list) != 0:
            # DE H1
            # MplCanvas 생성
            sc23 = pbd.ShowResult(self, width=5, height=6)

            # WSF plot
            sc23.axes.scatter(wall_result_output['DE_H1'], wall_result_output['Height(mm)'], color = 'k', s=1)

            # 허용치 기준선
            sc23.axes.axvline(x = DCR_criteria, color='r', linestyle='--')

            sc23.axes.set_xlim(0, xlim)
            sc23.axes.set_yticks(result.story_info['Height(mm)'][::-story_gap], result.story_info['Story Name'][::-story_gap])

            # 기타
            sc23.axes.grid(linestyle='-.')
            sc23.axes.set_xlabel('D/C Ratios')
            sc23.axes.set_ylabel('Story')
            sc23.axes.set_title('Shear Strength (X DE)')

            # toolbar 생성
            toolbar23 = NavigationToolbar(sc23, self)
            
            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar23, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc23, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # DE H2
            # MplCanvas 생성
            sc24 = pbd.ShowResult(self, width=5, height=6)

            # WSF plot
            sc24.axes.scatter(wall_result_output['DE_H2'], wall_result_output['Height(mm)'], color = 'k', s=1)

            # 허용치 기준선
            sc24.axes.axvline(x = DCR_criteria, color='r', linestyle='--')

            sc24.axes.set_xlim(0, xlim)
            sc24.axes.set_yticks(result.story_info['Height(mm)'][::-story_gap], result.story_info['Story Name'][::-story_gap])

            # 기타
            sc24.axes.grid(linestyle='-.')
            sc24.axes.set_xlabel('D/C Ratios')
            sc24.axes.set_ylabel('Story')
            sc24.axes.set_title('Shear Strength (Y DE)')

            # toolbar 생성
            toolbar24 = NavigationToolbar(sc24, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar24, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc24, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

        # MCE Plot
        if len(result.MCE_load_name_list) != 0:
            # MCE H1
            # MplCanvas 생성
            sc25 = pbd.ShowResult(self, width=5, height=6)

            # WSF plot
            sc25.axes.scatter(wall_result_output['MCE_H1'], wall_result_output['Height(mm)'], color = 'k', s=1)

            # 허용치 기준선
            sc25.axes.axvline(x = DCR_criteria, color='r', linestyle='--')

            sc25.axes.set_xlim(0, xlim)
            sc25.axes.set_yticks(result.story_info['Height(mm)'][::-story_gap], result.story_info['Story Name'][::-story_gap])

            # 기타
            sc25.axes.grid(linestyle='-.')
            sc25.axes.set_xlabel('D/C Ratios')
            sc25.axes.set_ylabel('Story')
            sc25.axes.set_title('Shear Strength (X MCE)')

            # toolbar 생성
            toolbar25 = NavigationToolbar(sc25, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar25, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc25, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # MCE H2
            # MplCanvas 생성
            sc26 = pbd.ShowResult(self, width=5, height=6)

            # WSF plot
            sc26.axes.scatter(wall_result_output['MCE_H2'], wall_result_output['Height(mm)'], color = 'k', s=1)

            # 허용치 기준선
            sc26.axes.axvline(x = DCR_criteria, color='r', linestyle='--')

            sc26.axes.set_xlim(0, xlim)
            sc26.axes.set_yticks(result.story_info['Height(mm)'][::-story_gap], result.story_info['Story Name'][::-story_gap])

            # 기타
            sc26.axes.grid(linestyle='-.')
            sc26.axes.set_xlabel('D/C Ratios')
            sc26.axes.set_ylabel('Story')
            sc26.axes.set_title('Shear Strength (Y MCE)')

            # toolbar 생성
            toolbar26 = NavigationToolbar(sc26, self)

            # layout에 toolbar, canvas
            layout.addWidget(toolbar26, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc26, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2



    # except Exception as e:
    #     self.status_browser.append('%s' %e)

    
    # spacer = QWidget()
    # spacer.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
    # layout.addWidget(spacer, layout.rowCount(), 0)

    # self.plot_display_area.setLayout(layout)
    # self.show()

    # 실행 시간 계산
    time_end = time.time()
    time_run = (time_end-time_start)/60
    self.status_browser.append('Completed!' + '  (total time = %0.5f min)' %(time_run))

    self.show_result_btn.setEnabled(True)
    self.print_result_btn.setEnabled(True)

    # Enable/Disable the Button
    # self.show_result_btn.setEnabled(False)
    # self.print_result_btn.setEnabled(False)
    # self.thread.finished.connect(lambda: self.show_result_btn.setEnabled(True))
    # self.thread.finished.connect(lambda: self.print_result_btn.setEnabled(True))