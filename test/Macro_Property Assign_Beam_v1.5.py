import pyautogui as pag
import pandas as pd

########################################### 처음에 입력해야 할 부분 ############################################
### Data Path
data_path = r'D:\이형우\성능기반 내진설계\22-RM-200 창원 신월2구역 재건축 아파트 신축공사 성능기반 내진설계\06. Data Conversion Sheets\103D_Data Conversion_Ver.1.5.xlsx'
beam_sheet = 'Input_C.Beam'

### 초기 좌표 획득(해당 위치에 마우스 올려놓고 ctrl+Enter 로 실행하면 순서대로 좌표 획득 가능)
# 반드시 H1 view 에서 좌표 획득하기
position_lefttop = pag.position() # 좌상단점
position_righttop = pag.position() # 우상단점
position_leftbot = pag.position() # 좌하단점
position_rightbot = pag.position() # 우하단점
position_p3dbar = pag.position() # 퍼폼 상태바
position_AssCom = pag.position() # Assign Component 버튼
position_CleSelEle = pag.position() # Clear Selected Elements 버튼
position_assign = pag.position() # Assign 버튼
position_cancel = pag.position() # Cancel 버튼
position_nextprop = pag.position() # 다음 property로 넘어가기 화살표
position_missingdata = pag.position() # Missing Data 알람 확인(아무 부재도 선택하지 않고 Assign 누를 시 뜨는 팝업 창에서의 확인 버튼)
position_nextframe = pag.position() # 다음 프레임 넘어가기 화살표

################ 옵션 ###############
drag_duration = 0.1 # drag 하는 속도(너무 빨리하면 팅길 수 있으므로 적당한 속도 권장)
offset = 10 # 픽셀 오차 방지용 여유치, 단위 : pixel
beam_index = 0 # elem_info 에서 시작하고자 하는 material name의 index 입력, 처음부터일때는 0 입력
#####################################
##############################################################################################################

# Story 정보 load
story_info = pd.read_excel(data_path, sheet_name='Story Data', keep_default_na=False
                           , index_col= None, skiprows=[0,2,3]).iloc[:,[1,2]]

# story 좌표 정의
story_info['mouse_coord'] = position_leftbot.y + (position_lefttop.y - position_leftbot.y) * (story_info['Level'] - story_info.iloc[len(story_info)-1,1])/(story_info.iloc[0, 1]-story_info.iloc[len(story_info)-1,1])
story_info = story_info.loc[::-1] # 편의를 위해 역순으로 배치
story_info = story_info.reset_index(drop=True, inplace=False)
story_info.reset_index(level=0, inplace=True) # level은 index를 몇번째 column으로 지정할 것인가

# E.Column 정보 불러오기
elem_info = pd.read_excel(data_path, sheet_name = beam_sheet, skiprows=[0,2,3]).iloc[:,0]

# underscore 2개 들어간 자료만 slice
elem_info = elem_info.loc[elem_info.str.count('_') == 2]

elem_info_splited_1 = elem_info.apply(lambda x: x.split('_')[0])
elem_info_splited_2 = elem_info.apply(lambda x: x.split('_')[1])
elem_info_splited_3 = elem_info.apply(lambda x: x.split('_')[2])
elem_info = pd.concat([elem_info, elem_info_splited_1, elem_info_splited_3, elem_info_splited_2], axis=1)
elem_info.columns = ['Name', 'Column Name', 'Story Name', 'Divide']

elem_info = pd.merge(elem_info, story_info[['Story Name', 'index']], how='left')

###################################### Property Assign 매크로 ##########################################

elem_name_count = elem_info.iloc[beam_index,0].split('_') # loop 돌면서 전 row와 비교하고, 부재 바뀔 때 프레임 넘기기
for i in range(beam_index, len(elem_info)):
    story_idx = elem_info.iloc[i,4]
    pag.click(position_p3dbar)
    
    # 이전, 현재 부재 비교 후, 다르면 프레임 넘기기
    elem_name_current = elem_info.iloc[i,0].split('_')
    if (elem_name_current[0] != elem_name_count[0]) | (elem_name_current[1] != elem_name_count[1]):
        pag.click(position_nextframe)
        elem_name_count = elem_name_current
     
    pag.moveTo(position_lefttop.x-50, story_info.iloc[story_idx, 3]+offset)
    pag.dragTo(position_righttop.x+50, story_info.iloc[story_idx, 3]-offset, duration=drag_duration)    
    pag.click(position_assign)
    pag.click(position_missingdata)
    pag.click(position_cancel)
    pag.click(position_cancel)
    pag.click(position_AssCom)
    pag.click(position_nextprop)

# 매크로 중간에 중단하고 싶을 때는 (조금 없어보이지만) ctrl+alt+Delete 누르기
# 매크로 작업중 다른 작업 불가능하므로 점심시간이나 퇴근시간 이용 추천
# Property Assign의 경우 혹여나 잘못 입력해도 다시 그 부재부터 다시 입력하면 됨.
###################################################################################################

