import pyautogui as pag
import pandas as pd

input_xlsx_path = r'D:\이형우\성능기반 내진설계\22-RM-200 창원 신월2구역 재건축 아파트 신축공사 성능기반 내진설계\06. Data Conversion Sheets\103D_Data Conversion_Ver.1.5.xlsx'

def macro(input_xlsx_path, start_or_end, macro_mode, *args):
    
    #%% 변수 정리
    pos_lefttop = args[0]
    pos_righttop = args[1]
    pos_leftbot = args[2]
    pos_rightbot = args[3]
    pos_p3dbar = args[4]
    pos_addcuts = args[5]
    pos_deletecuts = args[6]
    pos_ok = args[7]
    pos_nextsection = args[8]
    pos_nextframe = args[9]
    pos_ok_delete = args[10]
    pos_missingdata = args[11]
    pos_assigncom = args[12]
    pos_clearelem = args[13]
    drag_duration = args[14]
    offset = args[15]
    wall_name = args[16]

    #%% 변수, 이름 지정
    story_info_xlsx_sheet = 'Story Data'
    elem_info_xlsx_sheet = 'Input_S.Wall'
    
    # str -> int
    pos_lefttop = list(map(int, pos_lefttop.split(',')))
    print('OK')

    # Load Story Data
    story_info = pd.read_excel(input_xlsx_path, sheet_name=story_info_xlsx_sheet
                               , keep_default_na=False, index_col= None
                               , skiprows=[0,2,3], usecols=[1,2,4])
    # Load Section Name
    section_info = pd.read_excel(input_xlsx_path, sheet_name = elem_info_xlsx_sheet
                                 , skiprows=[0,2,3]).iloc[:,0]
    
    section_info_splited_1 = section_info.apply(lambda x: x.split('_')[0])
    section_info_splited_2 = section_info.apply(lambda x: x.split('_')[1])
    section_info_splited_3 = section_info.apply(lambda x: x.split('_')[2])
    section_info = pd.concat([section_info, section_info_splited_1, section_info_splited_3, section_info_splited_2], axis=1)
    section_info.columns = ['Name', 'Wall Name', 'Story Name', 'Divide']
    
    # Calculate Story Coordinates
    story_info['mouse_coord'] = pos_leftbot[1] + (pos_lefttop[1] - pos_leftbot[1]) * (story_info['Level'] - story_info.iloc[len(story_info)-1,1])/(story_info.iloc[0, 1]-story_info.iloc[len(story_info)-1,1])
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
    
    ################################# Run Macro ###################################
    
    # 해당 Wall name의 index 찾기
    wall_index = section_info[section_info['Name'] == wall_name]['index'][0]
    
    wall_name_count = section_info.iloc[wall_index,0].split('_') # loop 돌면서 전 row와 비교하고, 부재 바뀔 때 프레임 넘기기
    for i in range(wall_index, len(section_info)):
        story_idx = section_info.iloc[i,4]
        pag.click(pos_p3dbar)
        
        # 이전, 현재 부재 비교 후, 다르면 프레임 넘기기
        wall_name_current = section_info.iloc[i,0].split('_')
        if (wall_name_current[0] != wall_name_count[0]) | (wall_name_current[1] != wall_name_count[1]):
            pag.click(pos_nextframe)
            wall_name_count = wall_name_current
            
        pag.click(pos_addcuts)
        pag.moveTo(pos_leftbot[0]-200, story_info.iloc[story_idx, 4] - offset) # 시작하는 층의 좌하단 좌표
        pag.dragTo(pos_rightbot[0]+200, story_info.iloc[story_idx, 4]
                   -(story_info.iloc[story_idx, 4]-story_info.iloc[story_idx+1, 4])
                   /story_info.iloc[story_idx, 3] + offset, duration=drag_duration) # 윗점 찍기, divide 개수 고려
        pag.click(pos_ok) # OK click
        pag.click(pos_missingdata)
        pag.click(pos_nextsection) # next section click
        
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




# ###################### Macro for Deletion of Section ##########################

# # 혹여나, section을 잘못 입력하여 삭제해야 하는 경우 "연달아 n개"를 순서대로 삭제하는 매크로
# drag_duration = 0.3 # drag 속도 조정(너무 짧게 하면 팅길 수 있음)

# n = 60 # 연달아 삭제하고 싶은 section 개수
# for i in range(n):
#     pag.click(pos_p3dbar)
#     pag.click(pos_deletecuts)
#     pag.moveTo(pos_lefttop[0]-200, pos_lefttop[1]-10) # 3은 offset
#     pag.dragTo(pos_rightbot[0]+200, pos_rightbot[1]+10, duration=drag_duration) # 3은 offset
#     pag.click(pos_ok_delete)
#     pag.click(pos_missingdata)
#     pag.click(pos_nextsection)
    
# ###############################################################################




# ################### Macro for Assignment of Story Section #####################

# # 순서는 아랫층에서 윗층으로, 이미 section name 은 Base 부터 최상층까지 입력된 상태에서 실행!
# story_index = 0 # 맨 아래층부터 시작할 경우 0, 특정 층부터 시작할 경우 story_info 에서 해당 층 index 입력!
# drag_duration = 0.3 # drag 속도 조정(너무 짧게 하면 팅길 수 있음)
# for i in range(story_index, len(story_info)-1):
#     pag.click(pos_p3dbar)
#     pag.click(pos_addcuts)
#     pag.moveTo(pos_leftbot[0]-200, story_info.iloc[i, 4])
#     pag.dragTo(pos_rightbot[0]+200, story_info.iloc[i, 4] - (story_info.iloc[i, 4]-story_info.iloc[i+1, 4])/story_info.iloc[i, 3] + offset, duration=drag_duration)
#     pag.click(pos_ok)
#     pag.click(pos_missingdata)
#     pag.click(pos_nextsection)
    
# ###############################################################################

# #################### Macro for Assignment of Constraints ######################

# ################################# 추가 좌표 획득 ###############################
# pos_nextconst = pag.position() # 다음 constraints 화살표
# pos_Horizontal = pag.position() # Horizontal rigid floor 체크
# pos_Addnodes = pag.position() # Add Nodes 탭
# pos_constOK = pag.position() # OK 버튼
# pos_constmissingdata = pag.position() # 아무것도 안잡혔을 때 뜨는 Missing data 확인

# ############################## 매크로 실행 부분 ################################
# # 순서는 아랫층에서 윗층으로, 이미 constraint name 은 입력된 상태(층 분할된 경우 아래부터 1F-1, 1F-2, 1F-3, ...)
# drag_duration = 0.3 # drag 속도 조정(너무 짧게 하면 팅길 수 있음)
# story_index = 1 # Constraint는 Base 빼고 잡기 때문에 1부터 시작!!
# for i in range(story_index, len(story_info)-1):
#     for j in range(story_info.iloc[i, 3]): # 층 분할 고려하기 위해 추가
#         pag.click(pos_p3dbar)
#         pag.moveTo(pos_leftbot[0]-200, story_info.iloc[i, 4] - (story_info.iloc[i, 4] - story_info.iloc[i+1, 4])/story_info.iloc[i, 3]*j+ 5)
#         pag.dragTo(pos_rightbot[0]+200, story_info.iloc[i, 4] - (story_info.iloc[i, 4] - story_info.iloc[i+1, 4])/story_info.iloc[i, 3]*j - 5, duration=drag_duration)
#         pag.click(pos_constOK)
#         pag.click(pos_constmissingdata)
#         pag.click(pos_nextconst)