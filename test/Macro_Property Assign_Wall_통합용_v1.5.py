import pyautogui as pag
import pandas as pd
import os

import time
from threading import Thread
from pynput import keyboard

from PIL import ImageGrab
from functools import partial
ImageGrab.grab = partial(ImageGrab.grab, all_screens=True) # pag.locateOnScreen이 듀얼 모니터에서 안되는 문제 해결

#%% 

# def property_assign_macro(input_path, input_xlsx, drag_duration=0.15, offset=2, start_material_index=761): 

############################ 처음에 입력해야 할 부분 ###########################
### 초기 경로 설정
input_path = r'D:\이형우\성능기반 내진설계\21-GR-222 광명 4R구역 주택재개발사업 성능기반내진설계\101\101D_Data Conversion_Ver.1.2.xlsx' # Input Sheets 가 있는 폴더 경로

### 초기 좌표 획득(해당 위치에 마우스 올려놓고 ctrl+Enter 로 실행하면 순서대로 좌표 획득 가능)
# 반드시 H1 view 에서 좌표 획득하기
position_lefttop = pag.position() # 좌상단점
position_righttop = pag.position() # 우상단점
position_leftbot = pag.position() # 좌하단점
position_rightbot = pag.position() # 우하단점

################ 옵션 ###############
drag_duration = 0.15 # drag 하는 속도(너무 빨리하면 팅길 수 있으므로 적당한 속도 권장)
offset = 2 # 픽셀 오차 방지용 여유치, 단위 : pixel
wall_index = 761 # wall_material_info_repeat 에서 시작하고자 하는 material name의 index 입력, 처음부터일때는 0 입력
#####################################
# 매크로 중간에 중단하고 싶을 때는 (조금 없어보이지만) ctrl+alt+Delete 누르기
# 매크로 작업중 다른 작업 불가능하므로 점심시간이나 퇴근시간 이용 추천
# Property Assign의 경우 혹여나 잘못 입력해도 다시 그 부재부터 다시 입력하면 됨.
###############################################################################

# 자동 좌표 획득
position_p3dbar = pag.locateCenterOnScreen('images/p3d_status_bar.png'
                                           , confidence=0.6, grayscale=True) # 퍼폼 상태바
position_AssCom = pag.locateCenterOnScreen('images/assign_component.png'
                                           , confidence=0.6, grayscale=True) # Assign Component 버튼
position_CleSelEle = pag.locateCenterOnScreen('images/clear_selected_elements.png'
                                              , confidence=0.6, grayscale=True) # Clear Selected Elements 버튼
position_assign = pag.locateCenterOnScreen('images/assign.png'
                                           , confidence=0.6, grayscale=True) # Assign 버튼
position_cancel = pag.locateCenterOnScreen('images/cancel.png'
                                           , confidence=0.6, grayscale=True) # Cancel 버튼
position_missingdata = pag.locateCenterOnScreen('images/confirm_kr.png'
                                                , confidence=0.6, grayscale=True) 
# Missing Data 알람 확인(아무 부재도 선택하지 않고 Assign 누를 시 뜨는 팝업 창에서의 확인 버튼)
position_test = pag.locateCenterOnScreen('images/test.png'
                                                , confidence=0.8, grayscale=True) 


position_next = []
for i in pag.locateAllOnScreen('images/next.png', confidence=0.6, grayscale=True):
    position_next_x = i[0] + i[2]/2
    position_next_y = i[1] + i[3]/2
    
    position_next.append([position_next_x, position_next_y])
   
position_nextprop = position_next[2] # 다음 property로 넘어가기 화살표
position_nextframe = pag.locateCenterOnScreen('images/next_frame.png'
                                              , confidence=0.6, grayscale=True) # 다음 프레임 넘어가기 화살표
# test
pag.moveTo(position_test)
###############################################################################

# Story 정보 load
story_info = pd.read_excel(input_path, sheet_name='Story Data', keep_default_na=False
                           , index_col= None, skiprows=[0,2,3]).iloc[:,[1,2]]

# story 좌표 정의
story_info['mouse_coord'] = position_leftbot.y + (position_lefttop.y - position_leftbot.y) * (story_info['Level'] - story_info.iloc[len(story_info)-1,1])/(story_info.iloc[0, 1]-story_info.iloc[len(story_info)-1,1])
story_info = story_info.loc[::-1] # 편의를 위해 역순으로 배치
story_info = story_info.reset_index(drop=True, inplace=False)
story_info.reset_index(level=0, inplace=True) # level은 index를 몇번째 column으로 지정할 것인가

# Section 정보 불러오기
section_info = pd.read_excel(input_path, sheet_name = 'Input_S.Wall', skiprows=[0,2,3]).iloc[:,0]

section_info_splited_1 = section_info.apply(lambda x: x.split('_')[0])
section_info_splited_2 = section_info.apply(lambda x: x.split('_')[1])
section_info_splited_3 = section_info.apply(lambda x: x.split('_')[2])
section_info = pd.concat([section_info, section_info_splited_1, section_info_splited_3, section_info_splited_2], axis=1)
section_info.columns = ['Name', 'Wall Name', 'Story Name', 'Divide']

section_info = pd.merge(section_info, story_info[['Story Name', 'index']], how='left')

########################### Property Assign 매크로 ############################

wall_name_count = section_info.iloc[wall_index,0].split('_') # loop 돌면서 전 row와 비교하고, 부재 바뀔 때 프레임 넘기기
# wall_num_count = 
for i in range(wall_index, len(section_info)):
    story_idx = section_info.iloc[i,4]
    pag.click(position_p3dbar)
    
    # 이전, 현재 부재 비교 후, 다르면 프레임 넘기기
    wall_name_current = section_info.iloc[i,0].split('_')
    if (wall_name_current[0] != wall_name_count[0]) | (wall_name_current[1] != wall_name_count[1]):
        pag.click(position_nextframe)
        wall_name_count = wall_name_current
        
    pag.moveTo(position_lefttop.x-50, story_info.iloc[story_idx, 3]-offset)
    pag.dragTo(position_righttop.x+50, story_info.iloc[story_idx+1, 3]+offset, duration=drag_duration)    
    pag.click(position_assign)
    pag.click(position_missingdata)
    pag.click(position_cancel)
    pag.click(position_cancel)
    pag.click(position_AssCom)
    pag.click(position_nextprop)
        
#%%

def exit_program():
    def on_press(key):
        if str(key) == 'Key.esc':
            main.status = 'pause'
            user_input = input('Program paused, would you like to continue? (y/n) ')

            while user_input != 'y' and user_input != 'n':
                user_input = input('Incorrect input, try either "y" or "n" ')

            if user_input == 'y':
                main.status = 'run'
            elif user_input == 'n':
                main.status = 'exit'
                exit()

    with keyboard.Listener(on_press=on_press) as listener:
        listener.join()

def main():
    main.status = 'run'

    while True:
        print('running')
        time.sleep(1)

        while main.status == 'pause':
            time.sleep(1)

        if main.status == 'exit':
            print('Main program closing')
            break

Thread(target=main).start()
Thread(target=exit_program).start()