import pyautogui as pag
import pandas as pd
import numpy as np
import mss
import cv2

################################# USER INPUT ##################################

### Data Path
data_path = r'D:\이형우\성능기반 내진설계\22-RM-200 창원 신월2구역 재건축 아파트 신축공사 성능기반 내진설계\06. Data Conversion Sheets\103D_Data Conversion_Ver.1.5.xlsx'

### Mouse Coordinates 
# Locate mouse on the specific location on the monitor and Run
# Recommended to get the coords from H1 view
position_lefttop = pag.position() # 좌상단점
position_righttop = pag.position() # 우상단점
position_leftbot = pag.position() # 좌하단점
position_rightbot = pag.position() # 우하단점
position_p3dbar = pag.position() # 퍼폼 상태바
position_AddCuts = pag.position() # Add Cuts 탭
position_DeleteCuts = pag.position() # Delete Cuts 탭
position_OK = pag.position() # Add Cuts에서의 OK 버튼
position_nextsection = pag.position() # 다음 섹션으로 넘어가는 화살표
position_missingdata = pag.position() # Missing Data 알람 확인(Add Cuts에서 아무 부재도 선택하지 않고 OK 누를 시 뜨는 팝업 창에서 ok 버튼)
position_nextframe = pag.position() # 다음 프레임으로 넘어가는 화살표
position_OK_delete = pag.position() # Delete Cuts 에서 OK 버튼

###############################################################################

# Load Story Data
story_info = pd.read_excel(data_path, sheet_name='Story Data', keep_default_na=False
                           , index_col= None, skiprows=[0,2,3]).iloc[:, [1,2,4]]
# Load Section Name
section_info = pd.read_excel(data_path, sheet_name = 'Input_S.Wall', skiprows=[0,2,3]).iloc[:,0]

section_info_splited_1 = section_info.apply(lambda x: x.split('_')[0])
section_info_splited_2 = section_info.apply(lambda x: x.split('_')[1])
section_info_splited_3 = section_info.apply(lambda x: x.split('_')[2])
section_info = pd.concat([section_info, section_info_splited_1, section_info_splited_3, section_info_splited_2], axis=1)
section_info.columns = ['Name', 'Wall Name', 'Story Name', 'Divide']

# Calculate Story Coordinates
story_info['mouse_coord'] = position_leftbot.y + (position_lefttop.y - position_leftbot.y) * (story_info['Level'] - story_info.iloc[len(story_info)-1,1])/(story_info.iloc[0, 1]-story_info.iloc[len(story_info)-1,1])
story_info = story_info.loc[::-1] # 편의를 위해 역순으로 배치
story_info = story_info.reset_index(drop=True, inplace=False)
story_info.reset_index(level=0, inplace=True) # level은 index를 몇번째 column으로 지정할 것인가

# # 벽체별 층수 파악하기
# section_repeat = pd.DataFrame()
# for i in range(len(section_info)):
#     amount = section_info.iloc[i, 3]
#     for j in range(amount):
#         section_repeat = section_repeat.append(section_info.iloc[i, [0, 1, 2]])

# section_repeat = section_repeat.reset_index(drop=True, inplace=False)

# Merge Section Name with Story Data
section_info = pd.merge(section_info, story_info[['Story Name', 'index']], how='left')

# 벽체 2개 이상을 나눠지는 것 반영하기 위해 section_repeat에 section_info에 있던 amount 정보 join 하기
#section_repeat = section_repeat.join(section_info.set_index(''))


############################ USER INPUT for Macro #############################

wall_name = 'W1_1_B2' # index of the first wall (in "section_info") to be considered
# ex) 0 if every wall from the beginning of section_info should be considered
drag_duration = 0.3 # drag speed (there will be an error if too fast) ## RECOMMENDED TO USE THE DEFAULT VALUE
offset = 3  # increase if multiple stories are assigned at once / decrease if no story assigned
# change with increment/decrement of 0.1      ## RECOMMENDED TO USE THE DEFAULT VALUE

################################# Run Macro ###################################

# 해당 Wall name의 index 찾기
wall_index = section_info[section_info['Name'] == wall_name]['index'][0]

wall_name_count = section_info.iloc[wall_index,0].split('_') # loop 돌면서 전 row와 비교하고, 부재 바뀔 때 프레임 넘기기
for i in range(wall_index, len(section_info)):
    story_idx = section_info.iloc[i,4]
    pag.click(position_p3dbar)
    
    # 이전, 현재 부재 비교 후, 다르면 프레임 넘기기
    wall_name_current = section_info.iloc[i,0].split('_')
    if (wall_name_current[0] != wall_name_count[0]) | (wall_name_current[1] != wall_name_count[1]):
        pag.click(position_nextframe)
        wall_name_count = wall_name_current
        
    pag.click(position_AddCuts)
    pag.moveTo(position_leftbot.x-200, story_info.iloc[story_idx, 4] - offset) # 시작하는 층의 좌하단 좌표
    pag.dragTo(position_rightbot.x+200, story_info.iloc[story_idx, 4]
               -(story_info.iloc[story_idx, 4]-story_info.iloc[story_idx+1, 4])
               /story_info.iloc[story_idx, 3] + offset, duration=drag_duration) # 윗점 찍기, divide 개수 고려
    pag.click(position_OK) # OK click
    pag.click(position_missingdata)
    pag.click(position_nextsection) # next section click
    
    # ESC 누르면 멈춤
    # https://stackoverflow.com/questions/65399258/how-do-i-break-out-of-the-loop-at-anytime-with-a-keypress
    # import threading
    # import keyboard
    # try:
    #     if keyboard.is_pressed('esc'):
    #         break
    #     else: 
    #         pass
    # finally: 
    #     pass

# 매크로 중간에 중단하고 싶을 때는 (조금 없어보이지만) ctrl+alt+Delete 누르기
# 매크로 작업중 다른 작업 불가능하므로 점심시간이나 퇴근시간 이용 추천
###############################################################################




###################### Macro for Deletion of Section ##########################

# 혹여나, section을 잘못 입력하여 삭제해야 하는 경우 "연달아 n개"를 순서대로 삭제하는 매크로
drag_duration = 0.3 # drag 속도 조정(너무 짧게 하면 팅길 수 있음)

n = 60 # 연달아 삭제하고 싶은 section 개수
for i in range(n):
    pag.click(position_p3dbar)
    pag.click(position_DeleteCuts)
    pag.moveTo(position_lefttop.x-200, position_lefttop.y-10) # 3은 offset
    pag.dragTo(position_rightbot.x+200, position_rightbot.y+10, duration=drag_duration) # 3은 offset
    pag.click(position_OK_delete)
    pag.click(position_missingdata)
    pag.click(position_nextsection)
    
###############################################################################




################### Macro for Assignment of Story Section #####################

# 순서는 아랫층에서 윗층으로, 이미 section name 은 Base 부터 최상층까지 입력된 상태에서 실행!
story_index = 0 # 맨 아래층부터 시작할 경우 0, 특정 층부터 시작할 경우 story_info 에서 해당 층 index 입력!
drag_duration = 0.3 # drag 속도 조정(너무 짧게 하면 팅길 수 있음)
for i in range(story_index, len(story_info)-1):
    pag.click(position_p3dbar)
    pag.click(position_AddCuts)
    pag.moveTo(position_leftbot.x-200, story_info.iloc[i, 4])
    pag.dragTo(position_rightbot.x+200, story_info.iloc[i, 4] - (story_info.iloc[i, 4]-story_info.iloc[i+1, 4])/story_info.iloc[i, 3] + offset, duration=drag_duration)
    pag.click(position_OK)
    pag.click(position_missingdata)
    pag.click(position_nextsection)
    
###############################################################################

#################### Macro for Assignment of Constraints ######################

################################# 추가 좌표 획득 ###############################
position_nextconst = pag.position() # 다음 constraints 화살표
position_Horizontal = pag.position() # Horizontal rigid floor 체크
position_Addnodes = pag.position() # Add Nodes 탭
position_constOK = pag.position() # OK 버튼
position_constmissingdata = pag.position() # 아무것도 안잡혔을 때 뜨는 Missing data 확인

############################## 매크로 실행 부분 ################################
# 순서는 아랫층에서 윗층으로, 이미 constraint name 은 입력된 상태(층 분할된 경우 아래부터 1F-1, 1F-2, 1F-3, ...)
drag_duration = 0.3 # drag 속도 조정(너무 짧게 하면 팅길 수 있음)
story_index = 1 # Constraint는 Base 빼고 잡기 때문에 1부터 시작!!
for i in range(story_index, len(story_info)-1):
    for j in range(story_info.iloc[i, 3]): # 층 분할 고려하기 위해 추가
        pag.click(position_p3dbar)
        pag.moveTo(position_leftbot.x-200, story_info.iloc[i, 4] - (story_info.iloc[i, 4] - story_info.iloc[i+1, 4])/story_info.iloc[i, 3]*j+ 5)
        pag.dragTo(position_rightbot.x+200, story_info.iloc[i, 4] - (story_info.iloc[i, 4] - story_info.iloc[i+1, 4])/story_info.iloc[i, 3]*j - 5, duration=drag_duration)
        pag.click(position_constOK)
        pag.click(position_constmissingdata)
        pag.click(position_nextconst)