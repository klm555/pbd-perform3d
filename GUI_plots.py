import time
import pickle
import PBD_p3d as pbd
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt

from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT, FigureCanvasQTAgg
from matplotlib.figure import Figure

#%% Matplotlib Canvas
class MplCanvas(FigureCanvasQTAgg):
    
    def __init__(self, parent=None, width=5, height=4):
        self.fig = Figure(figsize=(width, height), tight_layout=True)
        self.axes = self.fig.add_subplot(111)
        FigureCanvasQTAgg.__init__(self, self.fig)
        FigureCanvasQTAgg.setMinimumSize(self, self.size())
        # super(ShowResult, self).__init__(self.fig)

# def load_time_count(self, time_start):
#     time_end = time.time()
#     time_run = (time_end-float(time_start))/60

#     # 완료 메세지 print
#     msg = 'Completed!' + '  (total time = %0.3f min)' %(time_run)
#     msg_colored = '<span style=\" color: #0000ff;\">%s</span>' % msg
#     self.status_browser.append(msg_colored)    

    # 실패할 경우, 그냥 msg_fn 이용해서 오류 메세지 내보내기
    # (단, 여기서 오류가 발생하는 경우, loadworker에서 데이터 처리하는 과정에서의 오류밖에 캐치 안됨)
    # self.worker.msg.connect(self.msg_fn)

#%% Tab3 - Show Plots on Display
def plot_display(self, result_dict):
    # 시작 메세지
    time_start = time.time()
    
    # 변수 정리
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

    # intger 변수
    story_gap = int(story_gap)
    max_shear = int(max_shear)
    # 기타 변수들 (향 후, UI에서 조작할 수 있게끔)
    cri_DE=0.015 # IDR
    cri_MCE=0.02 # IDR
    max_criteria=0.04 # WAS
    min_criteria=-0.002 # WAS
    DCR_criteria=1
    xlim = 3 # BR
    WAS_gage_group='AS' # WAS
    WSF_graph = True # WSF
    
    # QScrollAreadp container QWidget 생성
    self.plot_display_area.setWidgetResizable(True)
    container = QWidget()
    self.plot_display_area.setWidget(container)
    # grid layout
    layout = QGridLayout(container)
    row_count = 0
    
    #%% Base Shear Force 그래프
    if get_base_SF == True:
        # base_SF 결과값 가져오기    
        base_SF_result = result_dict['base_SF']
        # 결과값 classify & assign
        base_shear_H1 = base_SF_result[0]
        base_shear_H2 = base_SF_result[1]
        DE_load_name_list = base_SF_result[2]
        MCE_load_name_list = base_SF_result[3]
    
        # base_SF 그래프 그리기
        # DE Plot
        if len(DE_load_name_list) != 0:
            # H1_DE                
            sc1 = MplCanvas(self, width=6, height=4)
    
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
            toolbar1 = NavigationToolbar2QT(sc1, self)
    
            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar1, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc1, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2
    
            # H2_DE
            sc2 = MplCanvas(self, width=6, height=4)
    
            sc2.axes.bar(range(len(DE_load_name_list)), base_shear_H2.iloc[0, 0:len(DE_load_name_list)]\
                        , color='darkblue', edgecolor='k', label = 'Max. Base Shear')
            sc2.axes.axhline(y= base_shear_H2.iloc[0, 0:len(DE_load_name_list)].mean(), color='r', linestyle='-', label='Average')
            
            sc2.axes.set_ylim(0, max_shear)
            sc2.axes.set_xticks(range(14), range(1,15))
    
            sc2.axes.set_xlabel('Ground Motion No.')
            sc2.axes.set_ylabel('Base Shear(kN)')
            sc2.axes.legend(loc = 2)
            sc2.axes.set_title('Y DE')
    
            toolbar2 = NavigationToolbar2QT(sc2, self)
    
            layout.addWidget(toolbar2, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc2, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2
    
        # MCE Plot
        if len(MCE_load_name_list) != 0:
    
            # H1_MCE
            sc3 = MplCanvas(self, width=6, height=4)
            
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
            
            toolbar3 = NavigationToolbar2QT(sc3, self)
    
            layout.addWidget(toolbar3, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc3, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2
    
            # H2_MCE
            sc4 = MplCanvas(self, width=6, height=4)
            
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
    
            toolbar4 = NavigationToolbar2QT(sc4, self)
    
            layout.addWidget(toolbar4, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc4, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2
            
    #%% Story Shear Force 그래프
    if get_story_SF == True:
        # story_SF 결과값 가져오기    
        story_SF_result = result_dict['story_SF']
        # 결과값 classify & assign
        shear_force_H1_max = story_SF_result[0]
        shear_force_H2_max = story_SF_result[1]
        DE_load_name_list = story_SF_result[2]
        MCE_load_name_list = story_SF_result[3]

        # base_SF 그래프 그리기
        # DE Plot
        if len(DE_load_name_list) != 0:
            # H1_DE
            # MplCanvas 생성
            sc5 = MplCanvas(self, width=6, height=5)

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
            toolbar5 = NavigationToolbar2QT(sc5, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar5, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc5, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # H2_DE
            sc6 = MplCanvas(self, width=6, height=5)

            for i in range(len(DE_load_name_list)):
                sc6.axes.plot(shear_force_H2_max.iloc[:,i], range(shear_force_H2_max.shape[0]), label=DE_load_name_list[i], linewidth=0.7)
            
            sc6.axes.plot(shear_force_H2_max.iloc[:,0:len(DE_load_name_list)]\
                    .mean(axis=1), range(shear_force_H2_max.shape[0]), color='k', label='Average', linewidth=2)
            
            sc6.axes.set_xlim(0, max_shear)
            sc6.axes.set_yticks(range(shear_force_H2_max.shape[0])[::story_gap], shear_force_H2_max.index[::story_gap], fontsize=8.5)
        
            sc6.axes.grid(linestyle='-.')
            sc6.axes.set_xlabel('Story Shear(kN)')
            sc6.axes.set_ylabel('Story')
            sc6.axes.legend(loc=1, fontsize=8)
            sc6.axes.set_title('Y DE')

            toolbar6 = NavigationToolbar2QT(sc6, self)

            layout.addWidget(toolbar6, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc6, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

        # MCE Plot
        if len(MCE_load_name_list) != 0:

            # H1_MCE
            sc7 = MplCanvas(self, width=6, height=5)

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

            toolbar7 = NavigationToolbar2QT(sc7, self)

            layout.addWidget(toolbar7, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc7, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # H2_MCE
            sc8 = MplCanvas(self, width=6, height=5)

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

            toolbar8 = NavigationToolbar2QT(sc8, self)

            layout.addWidget(toolbar8, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc8, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2
    
    #%% Inter-Story Drift 그래프
    if get_IDR == True:
        # IDR 결과값 가져오기    
        IDR_result = result_dict['IDR']
        # 결과값 classify & assign
        IDR_result_each = IDR_result[0]
        IDR_result_avg = IDR_result[1]
        DE_load_name_list = IDR_result[2]
        MCE_load_name_list = IDR_result[3]
        story_name_window_reordered = IDR_result[4]

        # IDR 그래프 그리기
        count_x = 0
        count_y = 2
        count_avg = 0

        # DE Plot
        if len(DE_load_name_list) != 0:
            # H1_DE
            # MplCanvas 생성
            sc9 = MplCanvas(self, width=5, height=7)

            # MplCanvas에 그래프 그리기                    
            # 지진파별 plot
            for load_name in DE_load_name_list:
                sc9.axes.plot(IDR_result_each[count_x].iloc[:,-1], IDR_result_each[count_x].iloc[:,0]
                                , label='{}'.format(load_name), linewidth=0.7)
                sc9.axes.plot(IDR_result_each[count_x+1].iloc[:,-1], IDR_result_each[count_x].iloc[:,0]
                                , linewidth=0.7)
                count_x += 4
                
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
            toolbar9 = NavigationToolbar2QT(sc9, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar9, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc9, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # H2_DE
            # MplCanvas 생성
            sc10 = MplCanvas(self, width=5, height=7)

            # MplCanvas에 그래프 그리기                    
            # 지진파별 plot
            for load_name in DE_load_name_list:
                sc10.axes.plot(IDR_result_each[count_y].iloc[:,-1], IDR_result_each[count_y].iloc[:,0]
                                , label='{}'.format(load_name), linewidth=0.7)
                sc10.axes.plot(IDR_result_each[count_y+1].iloc[:,-1], IDR_result_each[count_y].iloc[:,0]
                                , linewidth=0.7)
                count_y += 4
                
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
            toolbar10 = NavigationToolbar2QT(sc10, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar10, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc10, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

        # DE Plot
        if len(MCE_load_name_list) != 0:
            # H1_MCE
            # MplCanvas 생성
            sc11 = MplCanvas(self, width=5, height=7)
            
            # MplCanvas에 그래프 그리기                    
            # 지진파별 plot
            for load_name in MCE_load_name_list:
                sc11.axes.plot(IDR_result_each[count_x].iloc[:,-1], IDR_result_each[count_x].iloc[:,0]
                                , label='{}'.format(load_name), linewidth=0.7)
                sc11.axes.plot(IDR_result_each[count_x+1].iloc[:,-1], IDR_result_each[count_x].iloc[:,0]
                                , linewidth=0.7)
                count_x += 4
                
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
            toolbar11 = NavigationToolbar2QT(sc11, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar11, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc11, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # H2_MCE
            # MplCanvas 생성
            sc12 = MplCanvas(self, width=5, height=7)

            # MplCanvas에 그래프 그리기                    
            # 지진파별 plot
            for load_name in MCE_load_name_list:
                sc12.axes.plot(IDR_result_each[count_y].iloc[:,-1], IDR_result_each[count_y].iloc[:,0]
                                , label='{}'.format(load_name), linewidth=0.7)
                sc12.axes.plot(IDR_result_each[count_y+1].iloc[:,-1], IDR_result_each[count_y].iloc[:,0]
                                , linewidth=0.7)
                count_y += 4
                
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
            toolbar12 = NavigationToolbar2QT(sc12, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar12, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc12, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2
            
    #%% C.Beam Rotation 그래프
    if get_BR == True:
        # BR 결과값 가져오기    
        BR_result = result_dict['BR']
        # 결과값 classify & assign
        BR_plot = BR_result[0]
        story_info = BR_result[1]
        DE_load_name_list = BR_result[2]
        MCE_load_name_list = BR_result[3]

        # BR 그래프 그리기
        # DE Plot
        if len(DE_load_name_list) != 0:
            # MplCanvas 생성
            sc13 = pbd.ShowResult(self, width=5, height=6)

            # DCR plot                
            sc13.axes.scatter(BR_plot['DCR(DE_pos)'], BR_plot['Height(mm)'], color='k', s=1)
            sc13.axes.scatter(BR_plot['DCR(DE_neg)'], BR_plot['Height(mm)'], color='k', s=1)

            # 허용치(DCR) 기준선
            sc13.axes.axvline(x = DCR_criteria, color='r', linestyle='--')
            sc13.axes.axvline(x = -DCR_criteria, color='r', linestyle='--')

            sc13.axes.set_xlim(-xlim, xlim)
            sc13.axes.set_yticks(story_info['Height(mm)'][::-story_gap], story_info['Story Name'][::-story_gap])

            # 기타
            sc13.axes.grid(linestyle='-.')
            sc13.axes.set_xlabel('D/C Ratios')
            sc13.axes.set_ylabel('Story')
            sc13.axes.set_title('Beam Rotation (DE)')

            # toolbar 생성
            toolbar13 = NavigationToolbar2QT(sc13, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar13, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc13, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2
            
        if len(MCE_load_name_list) != 0:
            # MplCanvas 생성
            sc14 = pbd.ShowResult(self, width=5, height=6)

            # DCR plot                
            sc14.axes.scatter(BR_plot['DCR(MCE_pos)'], BR_plot['Height(mm)'], color='k', s=1)
            sc14.axes.scatter(BR_plot['DCR(MCE_neg)'], BR_plot['Height(mm)'], color='k', s=1)

            # 허용치(DCR) 기준선
            sc14.axes.axvline(x = DCR_criteria, color='r', linestyle='--')
            sc14.axes.axvline(x = -DCR_criteria, color='r', linestyle='--')

            sc14.axes.set_xlim(-xlim, xlim)
            sc14.axes.set_yticks(story_info['Height(mm)'][::-story_gap], story_info['Story Name'][::-story_gap])

            # 기타
            sc14.axes.grid(linestyle='-.')
            sc14.axes.set_xlabel('D/C Ratios')
            sc14.axes.set_ylabel('Story')
            sc14.axes.set_title('Beam Rotation (MCE)')

            # toolbar 생성
            toolbar14 = NavigationToolbar2QT(sc14, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar14, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc14, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

    #%% C.Beam Shear Force 그래프
    if get_BSF == True:
        # BR 결과값 가져오기    
        BSF_result = result_dict['BSF']
        # 결과값 classify & assign
        BSF_plot = BSF_result[0]
        story_info = BSF_result[1]
        DE_load_name_list = BSF_result[2]
        MCE_load_name_list = BSF_result[3]
    
        # BSF 그래프 그리기
        # DE Plot
        if len(DE_load_name_list) != 0:
            # MplCanvas 생성
            sc15 = pbd.ShowResult(self, width=5, height=6)
    
            # DCR plot                
            sc15.axes.scatter(BSF_plot['DE'], BSF_plot['Height(mm)'], color='k', s=1)
    
            # 허용치(DCR) 기준선
            sc15.axes.axvline(x = DCR_criteria, color='r', linestyle='--')
    
            sc15.axes.set_xlim(0, xlim)
            sc15.axes.set_yticks(story_info['Height(mm)'][::-story_gap], story_info['Story Name'][::-story_gap])
    
            # 기타
            sc15.axes.grid(linestyle='-.')
            sc15.axes.set_xlabel('D/C Ratios')
            sc15.axes.set_ylabel('Story')
            sc15.axes.set_title('Shear Strength (DE)')
    
            # toolbar 생성
            toolbar15 = NavigationToolbar2QT(sc15, self)
    
            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar15, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc15, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2
            
        if len(MCE_load_name_list) != 0:
            # MplCanvas 생성
            sc16 = pbd.ShowResult(self, width=5, height=6)
    
            # DCR plot                
            sc16.axes.scatter(BSF_plot['MCE'], BSF_plot['Height(mm)'], color='k', s=1)
    
            # 허용치(DCR) 기준선
            sc16.axes.axvline(x = DCR_criteria, color='r', linestyle='--')
    
            sc16.axes.set_xlim(0, xlim)
            sc16.axes.set_yticks(story_info['Height(mm)'][::-story_gap], story_info['Story Name'][::-story_gap])
    
            # 기타
            sc16.axes.grid(linestyle='-.')
            sc16.axes.set_xlabel('D/C Ratios')
            sc16.axes.set_ylabel('Story')
            sc16.axes.set_title('Shear Strength (MCE)')
    
            # toolbar 생성
            toolbar16 = NavigationToolbar2QT(sc16, self)
    
            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar16, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc16, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

    #%% Wall Axial Strain 그래프
    if get_WAS == True:
        # WAS 결과값 가져오기    
        WAS_result = result_dict['WAS']
        # 결과값 classify & assign
        WAS_output = WAS_result[0]
        story_info = WAS_result[1]
        DE_load_name_list = WAS_result[2]
        MCE_load_name_list = WAS_result[3]
        
        # WAS 그래프 그리기
        # DE Plot
        if len(DE_load_name_list) != 0:
            # DE_1
            # MplCanvas 생성
            sc17 = pbd.ShowResult(self, width=5, height=4)

            # WAS plot
            sc17.axes.scatter(WAS_output['DE_min_avg'], WAS_output['Z(mm)'], color='r', s=5)
            sc17.axes.scatter(WAS_output['DE_max_avg'], WAS_output['Z(mm)'], color='k', s=5)

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
            toolbar17 = NavigationToolbar2QT(sc17, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar17, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc17, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # DE_2
            # MplCanvas 생성
            sc18 = pbd.ShowResult(self, width=5, height=4)

            # WAS plot
            sc18.axes.scatter(WAS_output['DE_min_avg'], WAS_output['Z(mm)'], color='r', s=5)
            sc18.axes.scatter(WAS_output['DE_max_avg'], WAS_output['Z(mm)'], color='k', s=5)

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
            toolbar18 = NavigationToolbar2QT(sc18, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar18, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc18, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2
            
            # 기준 넘는 점 확인
            error_coord_DE = WAS_output[(WAS_output['DE_max_avg'] >= max_criteria)
                                        | (WAS_output['DE_min_avg'] <= min_criteria)]
            
        # MCE Plot
        if len(MCE_load_name_list) != 0:
            # MCE_1
            # MplCanvas 생성
            sc19 = pbd.ShowResult(self, width=5, height=4)

            # WAS plot
            sc19.axes.scatter(WAS_output['MCE_min_avg'], WAS_output['Z(mm)'], color='r', s=5)
            sc19.axes.scatter(WAS_output['MCE_max_avg'], WAS_output['Z(mm)'], color='k', s=5)

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
            toolbar19 = NavigationToolbar2QT(sc19, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar19, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc19, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # MCE_2
            # MplCanvas 생성
            sc20 = pbd.ShowResult(self, width=5, height=4)

            # WAS plot
            sc20.axes.scatter(WAS_output['MCE_min_avg'], WAS_output['Z(mm)'], color='r', s=5)
            sc20.axes.scatter(WAS_output['MCE_max_avg'], WAS_output['Z(mm)'], color='k', s=5)

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
            toolbar20 = NavigationToolbar2QT(sc20, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar20, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc20, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # 기준 넘는 점 확인
            error_coord_MCE = WAS_output[(WAS_output['MCE_max_avg'] >= max_criteria)
                                        | (WAS_output['MCE_min_avg'] <= min_criteria)]

    #%% Wall Rotation 그래프
    if get_WR == True:
        # WR 결과값 가져오기    
        WR_result = result_dict['WR']
        # 결과값 classify & assign
        WR_plot = WR_result[0]
        story_info = WR_result[1]
        DE_load_name_list = WR_result[2]
        MCE_load_name_list = WR_result[3]
        
        # WR 그래프 그리기
        # DE Plot
        if len(DE_load_name_list) != 0:
            # DE
            # MplCanvas 생성
            sc21 = pbd.ShowResult(self, width=5, height=6)

            # WR plot
            sc21.axes.scatter(WR_plot['DCR(DE_pos)'], WR_plot['Height(mm)'], color='k', s=1)
            sc21.axes.scatter(WR_plot['DCR(DE_neg)'], WR_plot['Height(mm)'], color='k', s=1)

            # 허용치 기준선
            sc21.axes.axvline(x = DCR_criteria, color='r', linestyle='--')
            sc21.axes.axvline(x = -DCR_criteria, color='r', linestyle='--')

            sc21.axes.set_xlim(-xlim, xlim)
            sc21.axes.set_yticks(story_info['Height(mm)'][::-story_gap], story_info['Story Name'][::-story_gap])
            
            # 기타
            sc21.axes.grid(linestyle='-.')
            sc21.axes.set_xlabel('D/C Ratios')
            sc21.axes.set_ylabel('Story')
            sc21.axes.set_title('Wall Rotation (DE)')

            # toolbar 생성
            toolbar21 = NavigationToolbar2QT(sc21, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar21, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc21, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # 기준 넘는 벽체 확인
            error_wall_DE = WR_plot[['Name', 'DCR(DE_pos)', 'DCR(DE_neg)']]
            [(WR_plot['DCR(DE_pos)'] >= DCR_criteria) | (WR_plot['DCR(DE_neg)'] <= -DCR_criteria)]

        if len(MCE_load_name_list) != 0:
            # MCE
            # MplCanvas 생성
            sc22 = pbd.ShowResult(self, width=5, height=6)

            # WR plot
            sc22.axes.scatter(WR_plot['DCR(MCE_pos)'], WR_plot['Height(mm)'], color='k', s=1)
            sc22.axes.scatter(WR_plot['DCR(MCE_neg)'], WR_plot['Height(mm)'], color='k', s=1)

            # 허용치 기준선
            sc22.axes.axvline(x = DCR_criteria, color='r', linestyle='--')
            sc22.axes.axvline(x = -DCR_criteria, color='r', linestyle='--')

            sc22.axes.set_xlim(-xlim, xlim)
            sc22.axes.set_yticks(story_info['Height(mm)'][::-story_gap], story_info['Story Name'][::-story_gap])

            # 기타
            sc22.axes.grid(linestyle='-.')
            sc22.axes.set_xlabel('D/C Ratios')
            sc22.axes.set_ylabel('Story')
            sc22.axes.set_title('Wall Rotation (MCE)')

            # toolbar 생성
            toolbar22 = NavigationToolbar2QT(sc22, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar22, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc22, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

            # 기준 넘는 벽체 확인
            error_wall_MCE = WR_plot[['Name', 'DCR(MCE_pos)', 'DCR(MCE_neg)']]
            [(WR_plot['DCR(MCE_pos)'] >= DCR_criteria) | (WR_plot['DCR(MCE_neg)'] <= -DCR_criteria)]
    
    #%% Wall Shear Force 그래프
    if get_WSF == True:
        # WSF 결과값 가져오기    
        WSF_result = result_dict['WSF']
        # 결과값 classify & assign
        wall_result = WSF_result[0]
        story_info = WSF_result[1]
        DE_load_name_list = WSF_result[2]
        MCE_load_name_list = WSF_result[3]

        # WSF 그래프 그리기
        # DE Plot
        if len(DE_load_name_list) != 0:
            # DE
            # MplCanvas 생성
            sc23 = pbd.ShowResult(self, width=5, height=6)

            # WSF plot
            sc23.axes.scatter(wall_result['DE'], wall_result['Height(mm)'], color = 'k', s=1)

            # 허용치 기준선
            sc23.axes.axvline(x = DCR_criteria, color='r', linestyle='--')

            sc23.axes.set_xlim(0, xlim)
            sc23.axes.set_yticks(story_info['Height(mm)'][::-story_gap], story_info['Story Name'][::-story_gap])

            # 기타
            sc23.axes.grid(linestyle='-.')
            sc23.axes.set_xlabel('D/C Ratios')
            sc23.axes.set_ylabel('Story')
            sc23.axes.set_title('Shear Strength (DE)')

            # toolbar 생성
            toolbar23 = NavigationToolbar2QT(sc23, self)
            
            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar23, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc23, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2

        # MCE Plot
        if len(MCE_load_name_list) != 0:
            # MCE
            # MplCanvas 생성
            sc24 = pbd.ShowResult(self, width=5, height=6)

            # WSF plot
            sc24.axes.scatter(wall_result['MCE'], wall_result['Height(mm)'], color = 'k', s=1)

            # 허용치 기준선
            sc24.axes.axvline(x = DCR_criteria, color='r', linestyle='--')

            sc24.axes.set_xlim(0, xlim)
            sc24.axes.set_yticks(story_info['Height(mm)'][::-story_gap], story_info['Story Name'][::-story_gap])

            # 기타
            sc24.axes.grid(linestyle='-.')
            sc24.axes.set_xlabel('D/C Ratios')
            sc24.axes.set_ylabel('Story')
            sc24.axes.set_title('Shear Strength (MCE)')

            # toolbar 생성
            toolbar24 = NavigationToolbar2QT(sc24, self)

            # layout에 toolbar, canvas 추가
            layout.addWidget(toolbar24, row_count, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            layout.addWidget(sc24, row_count+1, 0, Qt.AlignHCenter|Qt.AlignVCenter)
            row_count += 2