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
        self.fig = Figure(figsize=(width, height), layout='tight')
        self.axes = self.fig.add_subplot(111)
        FigureCanvasQTAgg.__init__(self, self.fig)
        FigureCanvasQTAgg.setMinimumSize(self, self.size())
        # super(ShowResult, self).__init__(self.fig)

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

    story_gap = int(story_gap)
    max_shear = int(max_shear)
    
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
            
    #%% Base Shear Force 그래프
    if get_story_SF == True:
        # base_SF 결과값 가져오기    
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