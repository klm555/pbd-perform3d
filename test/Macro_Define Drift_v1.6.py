import pyautogui as pag
import pyperclip
import pandas as pd
import numpy as np
import time

################################# USER INPUT ##################################
### Initial Data Path
data_path = r'D:\이형우\성능기반 내진설계\용현학익7단지\708D\YH-708_Data Conversion_Ver.3.2_구조심의.xlsx'

### Mouse Coordinates (Locate mouse on the specific location on the monitor and Run)
# 반드시 H1 view 에서 좌표 획득하기
position_lefttop = pag.position() # 좌상단점
position_righttop = pag.position() # 우상단점
position_leftbot = pag.position() # 좌하단점
position_rightbot = pag.position() # 우하단점
position_drift1_x = pag.position() # H1 view 에서 drift1 frame의 x좌표
position_drift2_x = pag.position() # H1 view 에서 drift2 frame의 x좌표
position_drift3_x = pag.position() # H1 view 에서 drift3 frame의 x좌표
position_drift4_x = pag.position() # (optional)
position_drift5_x = pag.position() # (optional)
position_drift6_x = pag.position() # (optional)
position_p3dbar = pag.position() # 퍼폼 상태바
position_New = pag.position() # New 버튼
position_H1 = pag.position() # H1 check circle
position_H2 = pag.position() # H2 check circle
position_Test = pag.position() # Test 버튼
position_nextfram = pag.position() # 다음 프레임 버튼
#############################################################################################################

# Load Story Data
story_data = pd.read_excel(data_path, sheet_name='Story Data', keep_default_na=False, skiprows=[0, 2, 3], index_col= None).iloc[:, [1, 2]]

# story 좌표 정의
story_data['mouse_coord'] = position_leftbot.y + (position_lefttop.y - position_leftbot.y) * (story_data['Level'] - story_data.iloc[len(story_data)-1,1])/(story_data.iloc[0, 1]-story_data.iloc[len(story_data)-1,1])
story_data = story_data.loc[::-1] # 편의를 위해 역순으로 배치
story_data = story_data.reset_index(drop=True, inplace=False)
story_data.reset_index(level=0, inplace=True) # level은 index를 몇번째 column으로 지정할 것인가

# Load Drift Name
drift_name = pd.read_excel(data_path, sheet_name='Input_Naming', keep_default_na=False, skiprows=[0, 2, 3], index_col= None).iloc[:, 6]
drift_name = drift_name.to_frame(name='Drift Name')
drift_name = drift_name.replace(r'', np.nan, regex=True) # 빈 셀을 nan으로 바꾸기
drift_name = drift_name.dropna() # nan 제거

# Drift name 을 층, 시, 방향으로 split (for 문 쓰는것보다 훨씬 빠름)
drift_name['Story Name'] = drift_name['Drift Name'].str.split('_').str[0]
drift_name['Clock'] = drift_name['Drift Name'].str.split('_').str[1]
drift_name['Direction'] = drift_name['Drift Name'].str.split('_').str[-1]

# drift 위치 개수
drift_name_unique = drift_name['Clock'].unique()
drift_num = len(drift_name_unique)

# Drift 종류별 x좌표 모으기
drift_position_x = pd.DataFrame()
for i in range(drift_num):
    clock = drift_name_unique[i]
    coord = globals()['position_drift{}_x'.format(i+1)]
    drift_position_x_each = pd.Series([clock, coord.x])
    drift_position_x = pd.concat([drift_position_x, drift_position_x_each], axis=1)
drift_position_x = drift_position_x.T
drift_position_x.columns = ['Clock', 'X_coord']
# drift_position_x['X_coord'] = drift_position_x['X_coord'].astype(int)

# Drift name에 story 좌표 및 종류별 x좌표 join
drift_name = pd.merge(drift_name, story_data, how='left')
drift_name = pd.merge(drift_name, drift_position_x, how='left')

###################################### Drift 입력 매크로 ##########################################
### 중요 ### 반드시 매크로 돌리기 전에 CapsLock 켜져 있는지 확인!!
drift_index = 0 # 처음부터 시작할때는 0, 특정 drift 부터 시작할 때는 drift_name에서 해당 drift의 index 입력!
for i in range(drift_index, len(drift_name)):
    pag.click(position_p3dbar)
    pag.click(position_New)
    pyperclip.copy(drift_name.iloc[i,0])
    pag.keyDown('ctrl')
    pag.keyDown('v')
    pag.keyUp('v')
    pag.keyUp('ctrl')
    # pag.write(drift_name.iloc[i, 1], interval=0.25)
    # pag.keyDown('shift')
    # pag.press('-')
    # pag.keyUp('Shift')
    # time.sleep(0.5) # shift 밀림 방지
    # pag.write(drift_name.iloc[i, 2])
    # pag.keyDown('shift')
    # pag.press('-')
    # pag.keyUp('Shift')
    # time.sleep(0.5)  # shift 밀림 방지
    # pag.write(drift_name.iloc[i, 3], interval=0.25)
    if drift_name.iloc[i, 3] == 'X':
        pag.click(position_H1)
    else:
        pag.click(position_H2)
    pag.click(drift_name.iloc[i, 7], story_data.iloc[drift_name.iloc[i, 4] + 1, 3])
    pag.click(drift_name.iloc[i, 7], story_data.iloc[drift_name.iloc[i, 4], 3])
    pag.click(position_Test)
    pag.click(position_Test)
    if drift_name.iloc[i, 2] != drift_name.iloc[i+1, 2]:
        time.sleep(0.5) # 튕김 방지
        pag.click(position_nextfram)

# 매크로 중간에 중단하고 싶을 때는 ctrl+alt+Delete 누르기
###################################################################################################
