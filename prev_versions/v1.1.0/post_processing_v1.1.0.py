#%% Import

import os
import pandas as pd
import time
from io import BytesIO # 파일처럼 취급되는 문자열 객체 생성(메모리 낭비 down)
import multiprocessing as mp

import PBD_p3d.output_to_docx as otd
import PBD_p3d as pbd

#%% 시간 측정(START)
time_start = time.time()

#%% USER INPUT

###########################   FILE 경로    ####################################
# Analysis Result
result_xlsx_1 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\비선형해석모델\Results\107\KHSM_107_2_Analysis Result_1.xlsx'"
result_xlsx_2 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\비선형해석모델\Results\107\KHSM_107_2_Analysis Result_2.xlsx'"
result_xlsx_3 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\비선형해석모델\Results\107\KHSM_107_2_Analysis Result_3.xlsx'"
# result_xlsx_4 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\Results\107\KHSM_107_2_Analysis Result_4.xlsx'"
# result_xlsx_5 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\Results\107\KHSM_107_2_Analysis Result_5.xlsx'"
result_xlsx_path = result_xlsx_1 + ',' + result_xlsx_2 + ',' + result_xlsx_3  # + ',' + result_xlsx_4 + ',' + result_xlsx_5
result_xlsx_path = result_xlsx_path.split(',')
result_xlsx_path = [i.strip("'") for i in result_xlsx_path]
result_xlsx_path = [i.strip('"') for i in result_xlsx_path]
to_load_list = result_xlsx_path

# Data Conversion Sheet, Column Sheet, Beam Sheet
input_xlsx_path = r'K:\2105-이형우\성능기반 내진설계\KHSM\107\KHSM_107_Data Conversion_Ver.1.3M_내진상세_변경.xlsx'
# input_xlsx_path = r'K:\1110-김학범\김해신문1지구 A17-1BL\내진성능설계\107D\input\Data Conversion_Ver.1.3M_107D-6.xlsx'
retrofit_xlsx_path = r'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\최종보강안\04-벽체\230530-GIMHAE-107D-WALL LIST-0005.xlsx'
#########################   DOCX 출력 변수    ##################################
bldg_name = '107동'
DCR = 0.91

# 전체 결과
output_docx = '107_2_base_SF_test.docx' 
# 벽체 전단(보강)           *** 보강 Sheet를 작성해야 부재별 결과확인 함수 실행가능
WSF_docx = '107_2_벽체_수평배근_보강.docx'
wall_retrofit_sheet = 'Results_Wall_보강 (1)'
# 연결보 전단(보강)
BSF_docx = '107_2_연결보_수평배근_보강.docx'
beam_retrofit_sheet = 'Results_C.Beam_보강 (1)'
# 기둥 전단(보강)
CSF_docx = '107_2_기둥_수평배근_보강.docx'
col_retrofit_sheet = 'Results_G.Column_보강 (1)'
# 소성힌지 생성여부
beam_p_hinge_docx = '107_2_연결보_소성힌지.docx'
col_p_hinge_docx = '107_2_기둥_소성힌지.docx'

column_xlsx_path = 'KHSM_107_2_Results_E.Column_Ver.1.3.xlsx'
column_pdf_name = '107_2_E.Column_결과' # 출력될 pdf 파일의 이름 (확장자X)

#######################   WALL 전단 보강 변수    ###############################
rebar_limit = ['D13', 100] # 벽체 수평배근 보강 최소한계

###########################   ETC 변수    #####################################
# Group Name (Perform-3D)
beam_group = 'BEAM'
col_group = 'COLUMN'
AS_gage_group = 'AX' # P3D에서 벽체 축게이지 Group Name (gage 이름과 gage group 이름이 동일해야함)

max_shear = 90000 #kN, 그래프의 Y-limit in <Base Shear><Story Shear>
story_gap = 2 # 층간격

#%% Post Processing - TOTAL (Word로 출력)

# Execute functions for data analysis
base_SF = result.base_SF(result_xlsx_path, ylim=max_shear) # 밑면 전단력
story_SF = result.story_SF(result_xlsx_path, input_xlsx_path, yticks=story_gap, xlim=max_shear)


# base_SF = pbd.base_SF(result_xlsx_path, ylim=max_shear) # 밑면 전단력
# story_SF = pbd.story_SF(input_xlsx_path, result_xlsx_path, yticks=story_gap, xlim=max_shear) # 층 전단력
# IDR = pbd.IDR(input_xlsx_path, result_xlsx_path, yticks=story_gap) # 층간변위비

# AS = pbd.AS(input_xlsx_path, result_xlsx_path, yticks=story_gap, min_criteria=-0.002, AS_gage_group=AS_gage_group) # 벽체 축 변형률
# WR_DCR = pbd.SWR_DCR(input_xlsx_path, result_xlsx_path, yticks=story_gap, xlim=3) # 벽체 소성회전각(DCR)
# wall_SF = pbd.wall_SF(input_xlsx_path, result_xlsx_path, graph=True, yticks=story_gap, xlim=3) # 벽체 전단강도

# BR_DCR = pbd.BR_DCR(input_xlsx_path, result_xlsx_path, yticks=story_gap, xlim=3, c_beam_group=beam_group) # 연결보 소성회전각(DCR)
# BSF = pbd.BSF(input_xlsx_path, result_xlsx_path) # 연결보 전단강도

# CR_DCR = pbd.CR_DCR(input_xlsx_path, result_xlsx_path, yticks=story_gap, xlim=3, col_group=col_group) # 일반기둥 소성회전각(DCR)                                                    
# CSF = pbd.CSF(input_xlsx_path, result_xlsx_path) # 일반기둥 전단강도 
# E_CSF = pbd.E_CSF(input_xlsx_path, result_xlsx_path, column_xlsx_path, export_to_pdf=True)

# output_to_docx 이용해서 결과 출력
# 객체 생성(전체)
result = otd.OutputDocx(bldg_name, 'total')

result.base_SF_docx(base_SF)
result.story_SF_docx(story_SF)
# result.IDR_docx(IDR)
# result.WAS_docx(AS)
# result.WR_docx(WR_DCR)
# result.BR_docx(BR_DCR)
# result.CR_docx(CR_DCR)
# result.WSF_docx(wall_SF)
# result.BSF_docx(BSF)
# result.CSF_docx(CSF)

result.save_docx(result_xlsx_path, output_docx)

#%% 벽체 전단검토에 따른 수평배근 보강 (Excel에 자동 입력)
# (미완성)간격 변경만 됨. 직경은 수동으로 바꿔야 함.
WSF_retrofit = pbd.WSF_retrofit(input_xlsx_path, rebar_limit=rebar_limit) # 벽체 수평배근 보강 

# *** 기둥이나 보는 보강량이 적어서 코드 안만들었음. ***                         

#%% Post Processing - INDIVIDUAL (Word 출력) - 벽체 전단(보강)

# Execute functions for data analysis
WSF_each = pbd.WSF_each_HMW(input_xlsx_path, retrofit_xlsx_path, retrofit_sheet=wall_retrofit_sheet)
# output_to_docx 이용해서 결과 출력
result_each = otd.OutputDocx(bldg_name, 'each')
result_each.WSF_each_docx(WSF_each, DCR=DCR)
result_each.save_docx(result_xlsx_path, WSF_docx)

#%% Post Processing - INDIVIDUAL (Word 출력) - 연결보 전단(보강)

# Execute functions for data analysis
BSF_each = pbd.BSF_each(input_xlsx_path, retrofit_sheet=beam_retrofit_sheet)
# output_to_docx 이용해서 결과 출력
result_each = otd.OutputDocx(bldg_name, 'each')
result_each.BSF_each_docx(BSF_each, DCR=DCR)
result_each.save_docx(result_xlsx_path, BSF_docx)

#%% Post Processing - INDIVIDUAL (Word 출력) - 일반기둥 전단(보강)

# Execute functions for data analysis
CSF_each = pbd.CSF_each(input_xlsx_path, retrofit_sheet=col_retrofit_sheet)
# output_to_docx 이용해서 결과 출력
result_each = otd.OutputDocx(bldg_name, 'each')
result_each.CSF_each_docx(CSF_each, DCR=DCR)
result_each.save_docx(result_xlsx_path, CSF_docx)

#%% Post Processing - INDIVIDUAL (excel 입력 / Word 출력) - 소성힌지 생성 여부

# Execute functions for data analysis      *** excel 입력만 원할 시(word 출력 X),이 함수만 실행할 것! ***
beam_p_hinge, col_p_hinge = pbd.p_hinge(input_xlsx_path, result_xlsx_path)

# output_to_docx 이용해서 결과 출력 - Beam
p_hinge_result = otd.OutputDocx(bldg_name, 'beam_plastic_hinge')
p_hinge_result.p_hinge_docx(beam_p_hinge, 'beam')
p_hinge_result.save_docx(result_xlsx_path, beam_p_hinge_docx)
# output_to_docx 이용해서 결과 출력 - Column
p_hinge_result = otd.OutputDocx(bldg_name, 'column_plastic_hinge')
p_hinge_result.p_hinge_docx(col_p_hinge, 'column')
p_hinge_result.save_docx(result_xlsx_path, col_p_hinge_docx)

#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%% 허무원 박사님용(TEMPORARY)
pbd.CSF_HMW(input_xlsx_path, result_xlsx_path)
pbd.BSF_HMW(input_xlsx_path, result_xlsx_path)
#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#

#%% 시간 측정(END)
time_end = time.time()
time_run = (time_end-time_start)/60
print('\n', 'total time = %0.7f min' %(time_run))

#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%#%%

#########################     will be depriecated     #########################
# 연결보 소성회전각
# BR = pbd.BR(result_path, result_xlsx, input_xlsx_path\
#               , m_hinge_group_name, s_hinge_group_name=s_hinge_group_name\
#               , s_cri_DE=0.01, s_cri_MCE=0.025/1.2, yticks=2, xlim=0.03)
# 
# 연결보 소성회전각(Gage 설치 X)    
# BR_no_gage = pbd.BR_no_gage(input_xlsx_path, result_xlsx_path
#                             , cri_DE=0.01, cri_MCE=0.025/1.2, yticks=2, xlim=0.03)
# 
# 벽체 소성회전각
# SWR = pbd.SWR(input_xlsx_path, result_xlsx_path, yticks=story_gap\
                # , DE_criteria=0.002, MCE_criteria=0.004/1.2)
# 
# 일반기둥 소성회전각
# CR = pbd.CR(input_xlsx_path, result_xlsx_path
#             , yticks=2)  
# 
# 전이보 전단력(구버전)
# trans_beam_SF_old = pbd.trans_beam_SF(result_path, result_xlsx\
#                                         , input_xlsx_path, beam_xlsx)

##############################     In Process     #############################\
# 벽체 전단력(only graph)
# wall_SF_graph = pbd.wall_SF_graph(input_xlsx_path, input_xlsx_sheet='Results_Wall_보강', yticks=story_gap)
# 탄성보 전단력
# E_BSF = pbd.E_BSF(input_xlsx_path, result_xlsx_path, beam_xlsx, contour=True)
# 탄성기둥 전단강도
# E_CSF = pbd.E_CSF(input_xlsx_path, result_xlsx_path, column_xlsx_path
#                   , export_to_pdf=True, pdf_name=column_pdf_name)
# 탄성기둥 전단강도(only pdf)
# E_CSF_pdf = pbd.E_CSF_pdf(column_xlsx_path, pdf_name=column_pdf_name)
# Pushover
# pushover = system.pushover(result_xlsx_path, y_result_txt, base_SF_design, pp_x=, pp_y=)