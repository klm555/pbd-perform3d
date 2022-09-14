# 기존 변수 제거
all = [var for var in globals() if var[0] != "_"]
for var in all:
    del globals()[var]

#%% Import

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

import PBD_p3d as pbd

#%% 시간 측정(START)
time_start = time.time()

print('\n########## Result ##########\n')

###############################################################################
###############################################################################
#%% User Input

# Building Name
bldg_name = '101동'

# Analysis Result
result_path = r'D:\이형우\내진성능평가\광명 4R\해석 결과\101_6'
result_xlsx = 'Analysis Result' # 해석결과에 공통으로 포함되는 이름 (확장자X)

# Data Conversion Sheet, Column Sheet, Beam Sheet
input_path = r'D:\이형우\내진성능평가\광명 4R\101'
input_xlsx = 'Input Sheets(101_6)_v.1.8.xlsx'
column_xlsx = 'Results_E.Column(101_6)_Ver.1.3.xlsx'
beam_xlsx = 'Results_E.Beam(101_6).xlsx'

# Post-processed Result
output_path = result_path # result_path와 동일하게 설정. 바꿔도 됨
output_docx = '101_6_해석결과.docx'
column_pdf_name = '101_6_전이기둥 결과' # 출력될 pdf 파일의 이름 (확장자X)

# Base Shear
ylim = 60000 #kN, 그래프의 y limit

# Story Shear
xlim = 60000 # kN, 그래프의 x limit
story_gap = 2 # 층간격

# BR
yticks = 2

#%% Post Processing

# 밑면 전단력
base_SF = pbd.base_SF(result_path, ylim=ylim)

# 층 전단력
story_SF = pbd.story_SF(input_path, input_xlsx, result_path\
                            , yticks=story_gap, xlim=xlim)

# 층간변위비
IDR = pbd.IDR(input_path, input_xlsx, result_path, yticks=story_gap)

# 벽체 압축/인장 변형률
AS = pbd.AS(input_path, input_xlsx, result_path, yticks=story_gap)

# 벽체 전단강도
wall_SF = pbd.wall_SF(input_path, input_xlsx, result_path, graph=True\
                        , yticks=story_gap, xlim=3)

# 벽체 전단력(only graph)
# wall_SF_graph = pbd.wall_SF_graph(input_path, input_xlsx, yticks=story_gap)

# 벽체 소성회전각
SWR = pbd.SWR(input_path, input_xlsx, result_path, yticks=story_gap\
                , DE_criteria=0.002, MCE_criteria=0.004/1.2)

# 벽체 소성회전각(DCR)
SWR_DCR = pbd.SWR_DCR(input_path, input_xlsx, result_path\
                        , yticks=story_gap, xlim=3)

# 연결보 소성회전각(Gage 설치 X)    
BR_no_gage = pbd.BR_no_gage(result_path, result_xlsx, input_path\
                                , input_xlsx, cri_DE=0.01, cri_MCE=0.025/1.2\
                                , yticks=2, xlim=0.03)

# 연결보 소성회전각(DCR)
BR_DCR = pbd.BR_DCR(result_path, result_xlsx, input_path, input_xlsx\
                        , yticks=3, xlim=3)

# 전이보 전단력
trans_beam_SF = pbd.trans_beam_SF_2(result_path, result_xlsx, input_path\
                                        , input_xlsx, beam_xlsx, contour=True)

# 전이기둥 전단강도
trans_column_SF = pbd.trans_column_SF(result_path, result_xlsx, input_path\
                                , input_xlsx, column_xlsx, export_to_pdf=True\
                                    , pdf_name=column_pdf_name)

# 전이기둥 전단강도(only pdf)
# trans_column_SF_pdf = pbd.trans_column_SF_pdf(input_path, column_xlsx\
                                                  # , pdf_name=column_pdf_name)

###############################################################################
###############################################################################
##### will be depriecated #####
# 연결보 소성회전각
# BR = pbd.BR(result_path, result_xlsx, input_path, input_xlsx\
#              , m_hinge_group_name, s_hinge_group_name=s_hinge_group_name\
#              , s_cri_DE=0.01, s_cri_MCE=0.025/1.2, yticks=2, xlim=0.03)

# 전이보 전단력(구버전)
# trans_beam_SF_old = pbd.trans_beam_SF(result_path, result_xlsx\
#                                        , input_path, input_xlsx, beam_xlsx)

#%% 그래프 & df 리스트 만들기

function_list = ['base_SF', 'story_SF', 'IDR', 'AS', 'wall_SF'\
                 , 'wall_SF_graph', 'SWR', 'SWR_DCR', 'BR', 'BR_no_gage'\
                 , 'BR_DCR', 'trans_beam_SF']

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
        
# generator -> tuple
# def create_plot_df(plot_df_list):
#     plot_df_tuple_list = []    
#     for i in plot_df_list:
#         if not isinstance(i, tuple):
#             plot_df_tuple_list.append(tuple(i))        
#         else:
#             plot_df_tuple_list.append(i)            
#     return plot_df_tuple_list

# Multiprocessing으로 figure append
# num_core = mp.cpu_count()

# pool = mp.Pool(processes = num_core)
# plot_df_mp_result = pool.map(create_plot_df, plot_df_list)
# pool.close()
# pool.join()


# tuple을 list로 펼치기
plot_df_list_flat = list(sum(plot_df_tuple_list, ()))

# 전체 plot
plot_list = [x for x in plot_df_list_flat if not isinstance(x, pd.DataFrame)]

# 전체 dataframe
df_list = [x for x in plot_df_list_flat if isinstance(x, pd.DataFrame)]

#%% Word로 결과 정리

# Document 생성
output_word = docx.Document()

# Changing the page margins
output_word_sections = output_word.sections
for section in output_word_sections:
    section.top_margin = Cm(1)
    section.bottom_margin = Cm(0.44)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(1.5)

# 제목
output_word_title_para = output_word.add_paragraph()
output_word_title_run = output_word_title_para.add_run(bldg_name)
output_word_title_run.font.size = Pt(12)
output_word_title_run.bold = True

# 표 삽입  # int(-(-x//1)) = math.ceil()
output_word_table = output_word.add_table(int(-(-len(plot_list)//2)), 2)
output_word_table_faster = output_word_table._cells

# 그래프 사이즈(inch)
figsize_x, figsize_y = 8.7, 2.3 # cm

#%% Story Shear 그래프 그리기

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
output_word_table.style = 'Table Grid'
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

#%% 시간 측정(END)
time_end = time.time()
time_run = (time_end-time_start)/60
print('\n', 'total time = %0.7f min' %(time_run))