#%% Import

import os
import pandas as pd
import time
from io import BytesIO # 파일처럼 취급되는 문자열 객체 생성(메모리 낭비 down)
import multiprocessing as mp

import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Cm
from docx.oxml.ns import qn

import PBD_p3d.output_to_docx as otd
import PBD_p3d as pbd

#%% 시간 측정(START)
time_start = time.time()

###############################################################################
###############################################################################
#%% User Input

# Building Name
bldg_name = '107동'

# Analysis Result
result_xlsx_1 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\Results\107\KHSM_107_2_Analysis Result_1.xlsx'"
result_xlsx_2 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\Results\107\KHSM_107_2_Analysis Result_2.xlsx'"
result_xlsx_3 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\Results\107\KHSM_107_2_Analysis Result_3.xlsx'"
# result_xlsx_3 = r"'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\Results\105_3(no_DE_MCE6)\KHSM_105_3_Analysis Result(no_DE_MCE6).xlsx'"
result_xlsx_path = result_xlsx_1 + ',' + result_xlsx_2 + ',' + result_xlsx_3
result_xlsx_path = result_xlsx_path.split(',')
result_xlsx_path = [i.strip("'") for i in result_xlsx_path]
result_xlsx_path = [i.strip('"') for i in result_xlsx_path]
to_load_list = result_xlsx_path

# Data Conversion Sheet, Column Sheet, Beam Sheet
# input_xlsx_path = r'D:\이형우\성능기반 내진설계\22-GR-167 김해 신문1지구 도시개발사업 A1블록 공동주택 신축 성능기반내진설계\105\KHSM_105_Data Conversion_Ver.1.3M_수평배근_변경.xlsx'
input_xlsx_path = r'K:\2105-이형우\성능기반 내진설계\KHSM\107\KHSM_107_Data Conversion_Ver.1.3M_내진상세_변경.xlsx'

# Post-processed Result
output_docx = '107_2_해석결과.docx'
# column_xlsx_path = '101D_Results_E.Column_N1_Ver.1.3.xlsx'
# column_pdf_name = '101D_N1_E.Column_결과' # 출력될 pdf 파일의 이름 (확장자X)
appendix_docx = 'Appendix. Wall SF(elementwise).docx'

beam_group = 'C.Beam'
col_group = 'G.Column'
AS_gage_group = 'AS' # gage 이름과 gage group 이름이 동일해야함

# Base Shear
max_shear = 70000 #kN, 그래프의 y limit

# 층간격
story_gap = 2

#%% Post Processing

# 밑면 전단력
base_SF = pbd.base_SF(result_xlsx_path, ylim=max_shear)
# 층 전단력
story_SF = pbd.story_SF(input_xlsx_path, result_xlsx_path
                            , yticks=story_gap, xlim=max_shear)
# 층간변위비
IDR = pbd.IDR(input_xlsx_path, result_xlsx_path, yticks=story_gap)


# 벽체 압축/인장 변형률
AS = pbd.AS(input_xlsx_path, result_xlsx_path, yticks=story_gap, min_criteria=-0.002, AS_gage_group=AS_gage_group)
# 벽체 전단강도
wall_SF = pbd.wall_SF(input_xlsx_path, result_xlsx_path, graph=True\
                        , yticks=story_gap, xlim=3)
# 벽체 전단력(only graph)
# wall_SF_graph = pbd.wall_SF_graph(input_xlsx_path, input_xlsx_sheet='Results_Wall_보강', yticks=story_gap)
# 벽체 소성회전각(DCR)
WR_DCR = pbd.SWR_DCR(input_xlsx_path, result_xlsx_path
                        , yticks=story_gap, xlim=3)

# 연결보 소성회전각(DCR)
BR_DCR = pbd.BR_DCR(input_xlsx_path, result_xlsx_path
                    , yticks=story_gap, xlim=3, c_beam_group=beam_group)

# 일반기둥 소성회전각(DCR)
CR_DCR = pbd.CR_DCR(input_xlsx_path, result_xlsx_path
                    , yticks=story_gap, xlim=3, col_group=col_group)      
# 일반기둥 전단강도                                                  
CSF = pbd.CSF(input_xlsx_path, result_xlsx_path)

plastic_hinge = pbd.plastic_hinge(input_xlsx_path, result_xlsx_path)

# 탄성보 전단력
# E_BSF = pbd.E_BSF(input_xlsx_path, result_xlsx_path, beam_xlsx, contour=True)
# 탄성기둥 전단강도
# E_CSF = pbd.E_CSF(input_xlsx_path, result_xlsx_path, column_xlsx_path
#                   , export_to_pdf=True, pdf_name=column_pdf_name)
# 탄성기둥 전단강도(only pdf)
# E_CSF_pdf = pbd.E_CSF_pdf(column_xlsx_path, pdf_name=column_pdf_name)
                       
                                                  
# Pushover
# pushover = system.pushover(result_xlsx_path, y_result_txt, base_SF_design, pp_x=, pp_y=)

###############################################################################
###############################################################################
#########################     will be depriecated     #########################
# =============================================================================
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
# =============================================================================

#%% output_to_docx 이용해서 결과 출력
# 객체 생성
result = otd.OutputDocx(bldg_name)

result.base_SF_docx(base_SF)
result.story_SF_docx(story_SF)
result.IDR_docx(IDR)
result.WAS_docx(AS)
result.WR_docx(WR_DCR)
result.BR_docx(BR_DCR)
result.WSF_docx(wall_SF)
result.CR_docx(CR_DCR)
result.CSF_docx(CSF)

# 반드시 Results_C.Beam, Results_G.Column 시트를 완성하고 실행할 것
# result.plastic_hinge_docx(plastic_hinge) 

result.save_docx(result_xlsx_path, output_docx)

#%% 시간 측정(END)
time_end = time.time()
time_run = (time_end-time_start)/60
print('\n', 'total time = %0.7f min' %(time_run))


#%% 그래프 & df 리스트 만들기

function_list = ['base_SF', 'story_SF', 'IDR', 'AS', 'wall_SF'\
                 , 'wall_SF_graph', 'SWR', 'SWR_DCR', 'BR', 'BR_no_gage'\
                 , 'BR_DCR', 'CR', 'CR_DCR', 'trans_beam_SF', 'trans_column_SF', 'pushover']

# plot & dataframe 합쳐진 list 만들기
plot_df_list = []
for i in function_list:
    if i in globals():
        plot_df_list.append(globals()['{}'.format(i)])
        
plot_df_list = list(filter(None, plot_df_list))

# generator -> tuple
plot_df_tuple_list = []
for i in plot_df_list:
    if not isinstance(i, tuple):
        plot_df_tuple_list.append(tuple(i))        
    else:
        plot_df_tuple_list.append(i)
        
# =============================================================================
# generator -> tuple
# def create_plot_df(plot_df_list):
#     plot_df_tuple_list = []    
#     for i in plot_df_list:
#         if not isinstance(i, tuple):
#             plot_df_tuple_list.append(tuple(i))        
#         else:
#             plot_df_tuple_list.append(i)            
#     return plot_df_tuple_list
# 
# Multiprocessing으로 figure append
# num_core = mp.cpu_count()
# 
# pool = mp.Pool(processes = num_core)
# plot_df_mp_result = pool.map(create_plot_df, plot_df_list)
# pool.close()
# pool.join()
# =============================================================================


# tuple을 list로 펼치기
plot_df_list_flat = list(sum(plot_df_tuple_list, ()))

# 전체 plot
plot_list = [x for x in plot_df_list_flat if not isinstance(x, pd.DataFrame)]

# 전체 dataframe
df_list = [x for x in plot_df_list_flat if isinstance(x, pd.DataFrame)]

#%% Word로 결과 정리

# template 불러와서 Document 생성
output_word = docx.Document("template/report_template.docx")

# 제목
output_word_title_para = output_word.add_paragraph()
output_word_title_run = output_word_title_para.add_run('• ' + bldg_name)
output_word_title_run.font.size = Pt(9)
output_word_title_run.bold = True

# 표 삽입  # int(-(-x//1)) = math.ceil()
output_word_table = output_word.add_table(int(-(-len(plot_list)//2)), 2)
output_word_table_faster = output_word_table._cells

# 그래프 사이즈(inch)
figsize_x, figsize_y = 8.7, 2.3 # cm

#%% 전체 결과 그래프 그리기

count = 0
for i in plot_list:
    
    memfile = BytesIO()
    i.savefig(memfile)
    
    output_word_table_faster_para = output_word_table_faster[count]\
                                    .paragraphs[0]
    output_word_table_faster_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    output_word_table_faster_run = output_word_table_faster_para.add_run()
    # output_word_table_faster_run.add_picture(memfile, height = figsize_y\
    #                                          , width = figsize_x)
    # output_word_table_faster_run.add_picture(memfile, width=Cm(figsize_x)
    output_word_table_faster_run.add_picture(memfile, width=Cm(figsize_x))
    output_word_table_faster_run.add_break(WD_BREAK.PAGE)
    
    memfile.close()    
    count += 1

# Table 스타일  
output_word_table.style = 'no_borderlines'
output_word_table.autofit = False
output_word_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
# 스타일 지정(global)
output_word_style = output_word.styles['Normal']
output_word_style.font.name = '맑은 고딕'
output_word_style._element.rPr.rFonts\
    .set(qn('w:eastAsia'), '맑은 고딕') # 한글 폰트를 따로 설정해 준다
output_word_style.font.size = Pt(8) 

# 저장~
output_word.save(output_path + '\\' + output_docx)