import pyautogui as pag
import pandas as pd
import os
from PIL import ImageGrab
from functools import partial
import time
from threading import Thread
from pynput import keyboard
ImageGrab.grab = partial(ImageGrab.grab, all_screens=True) # pag.locateOnScreen이 듀얼 모니터에서 안되는 문제 해결

#%% 

# def property_assign_macro(input_path, input_xlsx, drag_duration=0.15, offset=2, start_material_index=761): 

############################ 처음에 입력해야 할 부분 ###########################
### 초기 경로 설정
input_path = r'C:\Users\hwlee\Desktop\Python\내진성능설계' # Input Sheets 가 있는 폴더 경로
input_xlsx = 'Data Conversion_Ver.1.0_220930.xlsx' # Input Sheets 이름

### 초기 좌표 획득(해당 위치에 마우스 올려놓고 ctrl+Enter 로 실행하면 순서대로 좌표 획득 가능)
# 반드시 H1 view 에서 좌표 획득하기
position_lefttop = pag.position() # 좌상단점
position_righttop = pag.position() # 우상단점
position_leftbot = pag.position() # 좌하단점
position_rightbot = pag.position() # 우하단점

################ 옵션 ###############
drag_duration = 0.15 # drag 하는 속도(너무 빨리하면 팅길 수 있으므로 적당한 속도 권장)
offset = 2 # 픽셀 오차 방지용 여유치, 단위 : pixel
start_material_index = 761 # wall_material_info_repeat 에서 시작하고자 하는 material name의 index 입력, 처음부터일때는 0 입력
#####################################
# 매크로 중간에 중단하고 싶을 때는 (조금 없어보이지만) ctrl+alt+Delete 누르기
# 매크로 작업중 다른 작업 불가능하므로 점심시간이나 퇴근시간 이용 추천
# Property Assign의 경우 혹여나 잘못 입력해도 다시 그 부재부터 다시 입력하면 됨.
###############################################################################

# 자동 좌표 획득
position_p3dbar = pag.locateCenterOnScreen('images/p3d_status_bar.png') # 퍼폼 상태바
position_AssCom = pag.locateCenterOnScreen('images/assign_component.png') # Assign Component 버튼
position_CleSelEle = pag.locateCenterOnScreen('images/clear_selected_elements.png') # Clear Selected Elements 버튼
position_assign = pag.locateCenterOnScreen('images/assign.png') # Assign 버튼
position_cancel = pag.locateCenterOnScreen('images/cancel.png') # Cancel 버튼
position_missingdata = pag.locateCenterOnScreen('images/confirm_kr.png') # Missing Data 알람 확인(아무 부재도 선택하지 않고 Assign 누를 시 뜨는 팝업 창에서의 확인 버튼)
position_next = []
for i in pag.locateAllOnScreen('images/next.png'):
    position_next_x = i[0] + i[2]/2
    position_next_y = i[1] + i[3]/2
    
    position_next.append([position_next_x, position_next_y])
   
position_nextprop = position_next[2] # 다음 property로 넘어가기 화살표
position_nextfram = pag.locateCenterOnScreen('images/next_frame.png') # 다음 프레임 넘어가기 화살표

# Input Sheet 정보 Load
story_info = pd.DataFrame()
wall_material_info = pd.DataFrame()

input_xlsx_sheet = 'Output_Wall Properties'
input_data_raw = pd.ExcelFile(input_path + '\\' + input_xlsx)
input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', 'Naming', input_xlsx_sheet]\
                                  , skiprows=[0,2,3])
input_data_raw.close()

story_info = input_data_sheets['Story Data'].iloc[:,1:3]
wall_material_info = input_data_sheets[input_xlsx_sheet].iloc[:,0]
naming_info = input_data_sheets['Naming'].iloc[:,8:12]

naming_info.columns = ['Wall Name', 'Story(from)', 'Story(to)', 'Amount']

# story 좌표 정의
story_info['mouse_coord'] = position_leftbot.y \
                            + (position_lefttop.y-position_leftbot.y)\
                              *(story_info['Level']-story_info.iloc[len(story_info)-1,1])\
                              /(story_info.iloc[0, 1]-story_info.iloc[len(story_info)-1,1])
story_info = story_info.loc[::-1] # 편의를 위해 역순으로 배치
story_info = story_info.reset_index(drop=True, inplace=False)
story_info.reset_index(level=0, inplace=True) # level은 index를 몇번째 column으로 지정할 것인가

# wall material name에서 wall name split
wall_name_split = []
for i in range(len(wall_material_info)):
    wall_name_split.append(wall_material_info.iloc[i].split('_')[0])
wall_name_split_dummy = wall_name_split.copy()
wall_name_split_dummy.append('empty')

# 각 wall 의 property 개수 count
wall_name_count = pd.DataFrame()
count = 1
for i in range(len(wall_name_split_dummy)-1): # dummy를 불러옴!
    if wall_name_split[i] == wall_name_split_dummy[i+1]:
        count = count+1
    else:
        wall_name_count = pd.concat([wall_name_count, pd.Series(count)], ignore_index=True)
        count = 1
wall_name_count = wall_name_count.astype(int)

naming_info = pd.concat([naming_info, wall_name_count], axis=1)
naming_info.columns.values[4] = 'material num'

wall_material_info_repeat = pd.DataFrame()
a = 0
for i, j  in zip(naming_info.iloc[:,3], naming_info.iloc[:,4]):
    for k in range(i):
        wall_material_info_repeat = pd.concat([wall_material_info_repeat, wall_material_info[a:a+j]], ignore_index=True)
    a = a+j

# wall material data repeat에서 벽 이름 부분만 다시 split
wall_material_info_repeat_split = pd.DataFrame()
for i in range(len(wall_material_info_repeat)):
    wall_material_info_repeat_split = pd.concat([wall_material_info_repeat_split, pd.Series(wall_material_info_repeat.iloc[i, 0].split('_')[0])], ignore_index=True)
wall_material_info_repeat = pd.concat([wall_material_info_repeat, wall_material_info_repeat_split], axis=1)
wall_material_info_repeat.columns = ['wall', 'Head']

# 각 property 층 정보 얻기
split = []
for i in wall_material_info:
    split.append(i.split('_')[-1])

story_from = []
for i in range(len(split)):
    if '-' in split[i]:
        story_from.append(split[i].split('-')[0])
    else:
        story_from.append(split[i])

story_to = []
for i in range(len(split)):
    if '-' in split[i]:
        story_to.append(split[i].split('-')[1])
    else:
        story_to.append(split[i])

story_total = pd.concat([pd.Series(wall_material_info), pd.Series(story_from)\
                         , pd.Series(story_to)], axis=1)
story_total.columns = ['material name', 'Story(from)', 'Story(to)']

# wall material info repeat 과 story total join/ story info index-match 하기
wall_material_info_repeat = wall_material_info_repeat.join(story_total.set_index('material name')['Story(from)'], on='wall')
wall_material_info_repeat = wall_material_info_repeat.join(story_total.set_index('material name')['Story(to)'], on='wall')

wall_material_info_repeat = wall_material_info_repeat.join(story_info.set_index('Story Name')['index'], on='Story(from)'); wall_material_info_repeat.rename({'index' : 'Story(from)_order'}, axis=1, inplace=True)
wall_material_info_repeat = wall_material_info_repeat.join(story_info.set_index('Story Name')['index'], on='Story(to)'); wall_material_info_repeat.rename({'index' : 'Story(to)_order'}, axis=1, inplace=True)

# wall material info repeat 과 naming_info의 각 벽체별 material 개수 join 하기
wall_material_info_repeat = wall_material_info_repeat.join(naming_info.set_index('Wall Name')['material num'], on='Head')

########################### Property Assign 매크로 ############################

for i in range(start_material_index, len(wall_material_info_repeat)):
    if i == start_material_index:
        pag.click(position_p3dbar)
        pag.moveTo(position_lefttop.x-50, story_info.iloc[wall_material_info_repeat.iloc[i, 4], 3]-offset)
        pag.dragTo(position_righttop.x+50, story_info.iloc[wall_material_info_repeat.iloc[i, 5]+1, 3]+offset, duration=drag_duration)
        pag.click(position_assign)
        pag.click(position_missingdata)
        pag.click(position_cancel)
        pag.click(position_cancel)
        pag.click(position_AssCom)
        pag.click(position_nextprop)
    elif wall_material_info_repeat.iloc[i-1, 1] == wall_material_info_repeat.iloc[i, 1]:
        if wall_material_info_repeat.iloc[i-1, 4] < wall_material_info_repeat.iloc[i, 4]:
            pag.click(position_p3dbar)
            pag.moveTo(position_lefttop.x-50, story_info.iloc[wall_material_info_repeat.iloc[i, 4], 3]-offset)
            pag.dragTo(position_righttop.x+50, story_info.iloc[wall_material_info_repeat.iloc[i, 5]+1, 3]+offset, duration=drag_duration)
            pag.click(position_assign)
            pag.click(position_missingdata)
            pag.click(position_cancel)
            pag.click(position_cancel)
            pag.click(position_AssCom)
            pag.click(position_nextprop)
        else:
            pag.click(position_p3dbar)
            pag.click(position_nextfram)
            for prop_num in range(wall_material_info_repeat.iloc[i, 6]):
                pag.rightClick(position_nextprop)
            pag.click(position_p3dbar)
            pag.moveTo(position_lefttop.x-50, story_info.iloc[wall_material_info_repeat.iloc[i, 4], 3]-offset)
            pag.dragTo(position_righttop.x+50, story_info.iloc[wall_material_info_repeat.iloc[i, 5]+1, 3]+offset, duration=drag_duration)
            pag.click(position_assign)
            pag.click(position_missingdata)
            pag.click(position_cancel)
            pag.click(position_cancel)
            pag.click(position_AssCom)
            pag.click(position_nextprop)
    else:
        pag.click(position_p3dbar)
        pag.click(position_nextfram)
        pag.moveTo(position_lefttop.x-50, story_info.iloc[wall_material_info_repeat.iloc[i, 4], 3]-offset)
        pag.dragTo(position_righttop.x+50, story_info.iloc[wall_material_info_repeat.iloc[i, 5]+1, 3]+offset, duration=drag_duration)
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