import pyautogui as pag
import pandas as pd



########################################### 처음에 입력해야 할 부분 ############################################
### 초기 경로 설정
data_path = r'C:\Users\khpark\Desktop\21-RM-513 광명 4R구역\퍼폼 모델링\111동\코드\data' # Input Sheets 가 있는 폴더 경로
Input_sheet_name = '111D_Data Conversion_Shear Wall Type_Ver.1.4.xlsx' # Input Sheets 이름

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
position_nextfram = pag.position() # 다음 프레임 넘어가기 화살표
##############################################################################################################

# Story 정보 load
story_info = pd.read_excel(data_path + '\\' + Input_sheet_name, sheet_name='Story Data', keep_default_na=False, index_col= None, skiprows=[0, 2, 3]).iloc[:, [1, 2]]
story_info.columns = ['Floor Name', 'Height(mm)']
# Section 정보 불러오기
section_info = pd.read_excel(data_path + '\\' + Input_sheet_name, sheet_name = 'Wall Naming', skiprows=[0, 2, 3]).iloc[:, [0, 1, 2, 3]]
section_info.columns = ['Wall Name', 'Story(from)', 'Story(to)', 'Amount']

# story 좌표 정의
story_info['mouse_coord'] = position_leftbot.y + (position_lefttop.y - position_leftbot.y) * (story_info['Height(mm)'] - story_info.iloc[len(story_info)-1,1])/(story_info.iloc[0, 1]-story_info.iloc[len(story_info)-1,1])
story_info = story_info.loc[::-1] # 편의를 위해 역순으로 배치
story_info = story_info.reset_index(drop=True, inplace=False)
story_info.reset_index(level=0, inplace=True) # level은 index를 몇번째 column으로 지정할 것인가

# wall material 정보 load
wall_material_data = pd.read_excel(data_path + '\\' + Input_sheet_name, sheet_name='Output_Wall Properties', skiprows=[0, 2, 3]).iloc[:, 0]
wall_material_data = pd.DataFrame(wall_material_data) # series to df
wall_material_data.columns = ['wall']

# wall material name에서 wall name split
wall_name_split = pd.DataFrame()
for i in range(len(wall_material_data)):
    wall_name_split = wall_name_split.append([wall_material_data.iloc[i, 0].split('_')[0]])
wall_name_split = wall_name_split.reset_index(drop=True, inplace=False)
wall_name_split_dummy = wall_name_split.append(pd.Series('empty'), ignore_index=True)

# 각 wall 의 property 개수 count
wall_name_count = pd.DataFrame()
count = 1
for i in range(len(wall_name_split_dummy)-1): # dummy를 불러옴!
    if wall_name_split.iloc[i, 0] == wall_name_split_dummy.iloc[i+1, 0]:
        count = count+1
    else:
        wall_name_count = wall_name_count.append(pd.Series(count), ignore_index=True)
        count = 1
wall_name_count = wall_name_count.astype(int)

section_info = pd.concat([section_info, wall_name_count], axis=1)
section_info.columns.values[4] = 'material num'

wall_material_data_repeat = pd.DataFrame()
a = 0
for i, j  in zip(section_info.iloc[:, 3], section_info.iloc[:, 4]):
    for k in range(i):
        wall_material_data_repeat = wall_material_data_repeat.append(pd.DataFrame(wall_material_data.iloc[a:a+j, 0]), ignore_index=True)
    a = a+j

# wall material data repeat에서 벽 이름 부분만 다시 split
wall_material_data_repeat_split = pd.DataFrame()
for i in range(len(wall_material_data_repeat)):
    wall_material_data_repeat_split = wall_material_data_repeat_split.append(pd.Series(wall_material_data_repeat.iloc[i, 0].split('_')[0]), ignore_index=True)
wall_material_data_repeat_split.columns= ['Head']
wall_material_data_repeat = pd.concat([wall_material_data_repeat, wall_material_data_repeat_split], axis=1)


# 각 property 층 정보 얻기
split = pd.DataFrame()
for i in wall_material_data.iloc[:, 0]:
    split = split.append([i.split('_')[-1]])
split = split.reset_index(drop=True, inplace=False)

story_from = pd.DataFrame()
for i in range(len(split)):
    if '-' in split.iloc[i, 0]:
        story_from = story_from.append([split.iloc[i, 0].split('-')[0]])
    else:
        story_from = story_from.append([split.iloc[i, 0]])
story_from = story_from.reset_index(drop=True, inplace=False)

story_to = pd.DataFrame()
for i in range(len(split)):
    if '-' in split.iloc[i, 0]:
        story_to = story_to.append([split.iloc[i, 0].split('-')[1]])
    else:
        story_to = story_to.append([split.iloc[i, 0]])
story_to = story_to.reset_index(drop=True, inplace=False)


story_total = pd.concat([wall_material_data, story_from, story_to], axis=1)
story_total.columns = ['material name', 'Story(from)', 'Story(to)']


# wall material data repeat 과 story total join/ story info index-match 하기
wall_material_data_repeat = wall_material_data_repeat.join(story_total.set_index('material name')['Story(from)'], on='wall')
wall_material_data_repeat = wall_material_data_repeat.join(story_total.set_index('material name')['Story(to)'], on='wall')

wall_material_data_repeat = wall_material_data_repeat.join(story_info.set_index('Floor Name')['index'], on='Story(from)'); wall_material_data_repeat.rename({'index' : 'Story(from)_order'}, axis=1, inplace=True)
wall_material_data_repeat = wall_material_data_repeat.join(story_info.set_index('Floor Name')['index'], on='Story(to)'); wall_material_data_repeat.rename({'index' : 'Story(to)_order'}, axis=1, inplace=True)

# wall material data repeat 과 section_info의 각 벽체별 material 개수 join 하기
wall_material_data_repeat = wall_material_data_repeat.join(section_info.set_index('Wall Name')['material num'], on='Head')

###################################### Property Assign 매크로 ##########################################
################ 옵션 ###############
drag_duration = 0.15 # drag 하는 속도(너무 빨리하면 팅길 수 있으므로 적당한 속도 권장)
offset = 2 # 픽셀 오차 방지용 여유치, 단위 : pixel
start_material_index = 761 # wall_material_data_repeat 에서 시작하고자 하는 material name의 index 입력, 처음부터일때는 0 입력
#####################################

for i in range(start_material_index, len(wall_material_data_repeat)):
    if i == start_material_index:
        pag.click(position_p3dbar)
        pag.moveTo(position_lefttop.x-50, story_info.iloc[wall_material_data_repeat.iloc[i, 4], 3]-offset)
        pag.dragTo(position_righttop.x+50, story_info.iloc[wall_material_data_repeat.iloc[i, 5]+1, 3]+offset, duration=drag_duration)
        pag.click(position_assign)
        pag.click(position_missingdata)
        pag.click(position_cancel)
        pag.click(position_cancel)
        pag.click(position_AssCom)
        pag.click(position_nextprop)
    elif wall_material_data_repeat.iloc[i-1, 1] == wall_material_data_repeat.iloc[i, 1]:
        if wall_material_data_repeat.iloc[i-1, 4] < wall_material_data_repeat.iloc[i, 4]:
            pag.click(position_p3dbar)
            pag.moveTo(position_lefttop.x-50, story_info.iloc[wall_material_data_repeat.iloc[i, 4], 3]-offset)
            pag.dragTo(position_righttop.x+50, story_info.iloc[wall_material_data_repeat.iloc[i, 5]+1, 3]+offset, duration=drag_duration)
            pag.click(position_assign)
            pag.click(position_missingdata)
            pag.click(position_cancel)
            pag.click(position_cancel)
            pag.click(position_AssCom)
            pag.click(position_nextprop)
        else:
            pag.click(position_p3dbar)
            pag.click(position_nextfram)
            for prop_num in range(wall_material_data_repeat.iloc[i, 6]):
                pag.rightClick(position_nextprop)
            pag.click(position_p3dbar)
            pag.moveTo(position_lefttop.x-50, story_info.iloc[wall_material_data_repeat.iloc[i, 4], 3]-offset)
            pag.dragTo(position_righttop.x+50, story_info.iloc[wall_material_data_repeat.iloc[i, 5]+1, 3]+offset, duration=drag_duration)
            pag.click(position_assign)
            pag.click(position_missingdata)
            pag.click(position_cancel)
            pag.click(position_cancel)
            pag.click(position_AssCom)
            pag.click(position_nextprop)
    else:
        pag.click(position_p3dbar)
        pag.click(position_nextfram)
        pag.moveTo(position_lefttop.x-50, story_info.iloc[wall_material_data_repeat.iloc[i, 4], 3]-offset)
        pag.dragTo(position_righttop.x+50, story_info.iloc[wall_material_data_repeat.iloc[i, 5]+1, 3]+offset, duration=drag_duration)
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

