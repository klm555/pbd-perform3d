import sys
import os
import shutil
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QSettings, QCoreApplication, QThread, QObject, Qt, pyqtSlot
from PyQt5.QtGui import QIcon, QPixmap

from PyQt5 import uic # ui 파일을 사용하기 위한 모듈

#%% UI

BR_ui_class = uic.loadUiType('BR_setting.ui')[0]
print_ui_class = uic.loadUiType('print_setting.ui')[0]
about_ui_class = uic.loadUiType('about.ui')[0]

class BRSettingWindow(QMainWindow, BR_ui_class):

    def __init__(self, status_browser):
        super().__init__()
        self.setupUi(self)
        self.status_browser = status_browser
                
        ##### setting에 저장된 value를 불러와서 입력
        # QSettings 클래스 생성
        QCoreApplication.setOrganizationName('CNP_Dongyang')
        QCoreApplication.setApplicationName('PBD_with_PERFORM-3D')
        self.setting = QSettings()

        # setting에 저장된 value를 불러와서 입력
        self.setting.beginGroup('setting_tab3_BR')
        self.BR_scale_factor_editbox.setText(self.setting.value('BR_scale_factor', '1.0'))
        self.setting.endGroup()   

        ##### Connect Button
        # Load File
        self.ok_BR_setting_btn.clicked.connect(self.ok_BR_setting)
        self.cancel_BR_setting_btn.clicked.connect(self.cancel_BR_setting)
        self.reset_BR_setting_btn.clicked.connect(self.reset_BR_setting)

    # 각 버튼에 function 지정
    def ok_BR_setting(self):        
        scale_factor = self.BR_scale_factor_editbox.text()
        try:
            # 입력된 값이 float형으로 변환되는지(사용할 수 있는 값인지) 확인
            scale_factor = float(scale_factor)
            # ok 누르면 창이 꺼지면서 입력된 값이 setting에 저장
            self.setting.beginGroup('setting_tab3_BR')
            self.setting.setValue('BR_scale_factor', self.BR_scale_factor_editbox.text())
            self.setting.endGroup()
        except:
            msg = '설정창에 잘못된 값이 입력되었습니다.'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
        
        self.close()
    
    def cancel_BR_setting(self):
        scale_factor = self.BR_scale_factor_editbox.text()
        try:
            scale_factor = float(scale_factor)
        except:
            msg = '설정창에 잘못된 값이 입력되었습니다.'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
        
        self.close()

    # BR 설정 reset function
    def reset_BR_setting(self):
        self.BR_scale_factor_editbox.setText(str(1.0))
        
class PrintSettingWindow(QMainWindow, print_ui_class):

    def __init__(self, status_browser):
        super().__init__()
        self.setupUi(self)
        self.status_browser = status_browser
                
        ##### setting에 저장된 value를 불러와서 입력
        # QSettings 클래스 생성
        QCoreApplication.setOrganizationName('CNP_Dongyang')
        QCoreApplication.setApplicationName('PBD_with_PERFORM-3D')
        self.setting = QSettings()

        # setting에 저장된 value를 불러와서 입력
        self.setting.beginGroup('setting_tab3_print')
        self.project_name_editbox.setText(self.setting.value('project_name', '성능기반 내진설계'))
        self.bldg_name_editbox.setText(self.setting.value('bldg_name', '1동'))
        self.story_gap_editbox.setText(self.setting.value('story_gap', '2'))
        self.max_shear_editbox.setText(self.setting.value('max_shear', '60000'))
        self.setting.endGroup()   

        ##### Connect Button
        # Load File
        self.ok_print_setting_btn.clicked.connect(self.ok_print_setting)
        self.cancel_print_setting_btn.clicked.connect(self.cancel_print_setting)
        self.reset_print_setting_btn.clicked.connect(self.reset_print_setting)
        
    # 각 버튼에 function 지정
    def ok_print_setting(self):        
        story_gap = self.story_gap_editbox.text()
        max_shear = self.max_shear_editbox.text()
        try:
            # 입력된 값이 float형으로 변환되는지(사용할 수 있는 값인지) 확인
            story_gap = float(story_gap)
            max_shear = float(max_shear)
            # ok 누르면 창이 꺼지면서 입력된 값이 setting에 저장
            self.setting.beginGroup('setting_tab3_print')
            self.setting.setValue('project_name', self.project_name_editbox.text())
            self.setting.setValue('bldg_name', self.bldg_name_editbox.text())
            self.setting.setValue('story_gap', self.story_gap_editbox.text())
            self.setting.setValue('max_shear', self.max_shear_editbox.text())
            self.setting.endGroup()
        except:
            msg = '설정창에 잘못된 값이 입력되었습니다.'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
        
        self.close()
    
    def cancel_print_setting(self):
        story_gap = self.story_gap_editbox.text()
        max_shear = self.max_shear_editbox.text()
        try:
            # 입력된 값이 float형으로 변환되는지(사용할 수 있는 값인지) 확인
            story_gap = float(story_gap)
            max_shear = float(max_shear)
        except:
            msg = '설정창에 잘못된 값이 입력되었습니다.'
            msg_colored = '<span style=\" color: #ff0000;\">%s</span>' % msg
            self.status_browser.append(msg_colored)
            
        self.close()

    # Print 설정 reset function
    def reset_print_setting(self):
        self.project_name_editbox.setText('성능기반 내진설계')
        self.bldg_name_editbox.setText('1동')
        self.story_gap_editbox.setText('2')
        self.max_shear_editbox.setText('60000')
        
class AboutWindow(QMainWindow, about_ui_class):

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        
        # CNP 동양 로고
        self.qPixmapVar2 = QPixmap()
        self.qPixmapVar2.load('./images/CNP_logo.png')
        self.CNP_img_2.setPixmap(self.qPixmapVar2.scaled(self.CNP_img_2.size()
                                , transformMode=Qt.SmoothTransformation))