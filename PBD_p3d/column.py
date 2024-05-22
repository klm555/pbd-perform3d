import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from collections import deque # Double-ended Queue : 자료의 앞, 뒤 양 방향에서 자료를 추가하거나 제거가능
import win32com.client
import pythoncom
from PyPDF2 import PdfMerger, PdfFileReader

#%% Elastic Column SF (DCR)
def E_CSF(self, input_xlsx_path, col_design_xlsx_path, export_to_pdf=True\
          , pdf_name='E.Column Results', DCR_criteria=1, yticks=2, xlim=3):
    ''' 

    Perform-3D 해석 결과에서 기둥의 축력, 전단력, 모멘트를 불러와 Results_E.Column 엑셀파일을 작성. \n
    result_path : Perform-3D에서 나온 해석 파일의 경로. \n
    result_xlsx : Perform-3D에서 나온 해석 파일의 이름. 해당 파일 이름이 포함된 파일들을 모두 불러온다. \n
    input_path : Data Conversion 엑셀 파일의 경로 \n
    input_xlsx : Data Conversion 엑셀 파일의 이름. result_xlsx와는 달리 확장자명(.xlsx)까지 기입해줘야한다. 하나의 파일만 불러온다. \n
    column_xlsx : Results_E.Column 엑셀 파일의 이름.확장자명(.xlsx)까지 기입해줘야한다. \n
    export_to_pdf : 입력된 값에 따른 각 부재들의 결과 시트를 pdf로 출력. True = pdf 출력, False = pdf 미출력(Results_E.Column 엑셀파일만 작성됨).
    pdf_name = 출력할 pdf 파일 이름.
    
    '''
#%% Input Sheet 정보 load
    # Data Conversion Sheets
    story_info = self.story_info
    col_info = self.ecol_info.copy()
    rebar_info = self.rebar_info

    # Analysis Result Sheets
    node_data = self.node_data
    element_data = self.frame_data
    force_info_data = self.beam_shear_force_data

    # Seismic Loads List
    load_name_list = self.load_name_list
    gravity_load_name = self.gravity_load_name
    seismic_load_name_list = self.seismic_load_name_list
    DE_load_name_list = self.DE_load_name_list
    MCE_load_name_list = self.MCE_load_name_list
    
    # Data Conversion Sheets
    # story_info = result.story_info
    # col_info = result.ecol_info.copy()
    # rebar_info = result.rebar_info

    # # Analysis Result Sheets
    # node_data = result.node_data
    # element_data = result.frame_data
    # force_info_data = result.beam_shear_force_data

    # # Seismic Loads List
    # load_name_list = result.load_name_list
    # gravity_load_name = result.gravity_load_name
    # seismic_load_name_list = result.seismic_load_name_list
    # DE_load_name_list = result.DE_load_name_list
    # MCE_load_name_list = result.MCE_load_name_list

#%% Process Data
    # Rebar info 필요한 열만 추출
    rebar_info = rebar_info.iloc[:,[0,1,2]]

    # 필요한 부재만 선별
    prop_name = col_info.iloc[:,0]
    prop_name.name = 'Property Name'
    element_data = element_data[element_data['Property Name'].isin(prop_name)]

    element_data = element_data.drop_duplicates()
    node_data = node_data.drop_duplicates()   
    
    # Analysis Result에 Element, Node 정보 매칭
    element_data = pd.merge(element_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
    force_ongoing = pd.merge(element_data, force_info_data, how='left')
    force_ongoing.reset_index(inplace=True, drop=True)

#%% Calculate and Input Rebar Strength
    rebar_info = rebar_info.set_index('Type')

    main_rebar_strength = []
    hoop_rebar_strength = []
    for idx, row in col_info.iterrows():
        main_rebar_strength.append(rebar_info.loc[row['Main Rebar(DXX)'], row['Main Rebar Type']])
        hoop_rebar_strength.append(rebar_info.loc[row['Hoop(DXX)'], row['Hoop Type']])
        
    col_info['Main Rebar Strength'] = main_rebar_strength
    col_info['Hoop Strength'] = hoop_rebar_strength
    col_info['Cover Thickness(mm)'] = 40
    
    # 재정렬
    col_output = col_info.iloc[:,np.r_[0,1,2,15,18,3,5,16,7,17,8:15]]

#%% V, M 값 뽑기

    # 필요한 정보들만 다시 모아서 new dataframe
    force_ongoing = force_ongoing.iloc[:, [1,0,7,9,10,11,14,12,15,16]]    

    # 필요한 하중만 포함된 데이터 slice (MCE)
    force_ongoing_MCE = force_ongoing[force_ongoing['Load Case'].str.contains('|'.join(MCE_load_name_list))]
    # function equivalent of a combination of df.isin() and df.str.contains()

    # 결과값 모두 absolute
    columns_to_abs = ['P I-End', 'V3 I-End', 'V2 I-End', 'M2 I-End', 'M3 I-End']
    force_ongoing_MCE[columns_to_abs] = force_ongoing_MCE[columns_to_abs].abs()

    # 지진하중에 따라 Shear Force 데이터 Grouping
    force_grouped_list = list(force_ongoing_MCE.groupby('Load Case'))    

    # 해석 결과 상관없이 Full 지진하중 이름 list 만들기
    full_MCE_load_name_list = 'MCE' + pd.Series([11,12,21,22,31,32,41,42,51,52,61,62,71,72]).astype(str)
    
    # 이름만 들어간 Dataframe 만들기
    E_CSF_output = pd.DataFrame(prop_name)
    
    # 지진하중, i,j 노드, Max,Min loop 돌리기
    for load_name in full_MCE_load_name_list:
        # 만들어진 Group List loop 돌리기
        for force_grouped in force_grouped_list:
            if load_name in force_grouped[0]:
                # 같은 결과가 2개씩 있어서 drop_duplicates
                force_grouped_df = force_grouped[1].drop_duplicates()
                # Element Name이 같은 경우(부재가 잘려서 모델링 된 경우 등), 큰 값만 선택
                force_grouped_df = force_grouped_df.sort_values(by='P I-End')
                force_grouped_df = force_grouped_df.drop_duplicates(subset=['Element Name', 'Step Type'], keep='last')
                # P, V, M의 최대값 뽑기
                force_grouped_df_max = force_grouped_df.groupby('Element Name')\
                    .agg({'Property Name':'first', 'V':'first', 'P I-End':'max', 'V3 I-End':'max', 'V2 I-End':'max'
                          , 'M2 I-End':'max', 'M3 I-End':'max'})\
                        [['Property Name', 'V', 'P I-End', 'V3 I-End', 'V2 I-End', 'M2 I-End', 'M3 I-End']]

                # Input 시트의 부재 순서대로 재정렬
                force_grouped_df_max = pd.merge(prop_name, force_grouped_df_max, how='left')
                force_grouped_df_max.reset_index(inplace=True, drop=True)
                E_CSF_output = pd.concat([E_CSF_output, force_grouped_df_max\
                                          [['P I-End', 'V3 I-End', 'V2 I-End', 'M2 I-End'
                                           , 'M3 I-End']]], axis=1)
            
        # 해당 지진하중의 해석결과가 없는 경우 Blank Column 생성
        if load_name not in MCE_load_name_list: 
            blank_col = pd.Series([''] * len(prop_name))
            E_CSF_output = pd.concat([E_CSF_output, blank_col], axis=1)      

#%% 결과 정리 후 Input Sheets에 넣기
    
    # Design_E.Column 시트    
    # 결과값 없는 부재 제거
    idx_to_slice = E_CSF_output.iloc[:,1:].dropna().index # dropna로 결과값(DE,MCE) 있는 부재만 남긴 후 idx 추출
    idx_to_slice2 = col_info['Name'].iloc[idx_to_slice].index # 결과값 있는 부재만 slice 후 idx 추출
    col_info = col_info.iloc[idx_to_slice2,:]
    col_info.reset_index(inplace=True, drop=True)

    # nan인 칸을 ''로 바꿔주기
    E_CSF_output = E_CSF_output.replace(np.nan, '', regex=True)
    col_output = col_output.replace(np.nan, '', regex=True)

#%% 출력 (Using win32com...)    
    # Using win32com...
    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게

    wb = excel.Workbooks.Open(col_design_xlsx_path)
    ws1 = wb.Sheets('Results_E.Column')
    ws2 = wb.Sheets('Design_E.Column')
    
    startrow, startcol = 5, 1
    
    # Results_E.Column 시트 입력
    ws1.Range('A%s:BS%s' %(startrow, 5000)).ClearContents()
    ws1.Range('A%s:BS%s' %(startrow, startrow + E_CSF_output.shape[0] - 1)).Value\
        = list(E_CSF_output.itertuples(index=False, name=None))
        
    # Design_E.Column 시트 입력
    ws2.Range('B%s:R%s' %(startrow, 5000)).ClearContents()
    ws2.Range('B%s:R%s' %(startrow, startrow + col_output.shape[0] - 1)).Value\
        = list(col_output.itertuples(index=False, name=None))
    
    wb.Save()           
    # wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application

#%% Column Rotation (DCR)
def CR(self, DCR_criteria=1, yticks=2, xlim=3):

#%% Load Data
    # Data Conversion Sheets
    story_info = self.story_info
    deformation_cap = self.col_deform_cap
    
    # Analysis Result Sheets
    node_data = self.node_data
    element_data = self.frame_data
    beam_rot_data = self.beam_rot_data

    # Seismic Loads List
    load_name_list = self.load_name_list
    gravity_load_name = self.gravity_load_name
    seismic_load_name_list = self.seismic_load_name_list
    DE_load_name_list = self.DE_load_name_list
    MCE_load_name_list = self.MCE_load_name_list
    
#%% Process Data
    # temporary ((L), (R) 등 지우기)
    element_data.loc[:, 'Property Name'] = element_data.loc[:, 'Property Name'].str.split('(').str[0]
    
    # 필요한 부재만 선별
    prop_name = deformation_cap.iloc[:,0]
    prop_name.name = 'Property Name'
    element_data = element_data[element_data['Property Name'].isin(prop_name)]

    element_data = element_data.drop_duplicates()
    node_data = node_data.drop_duplicates()
    
    # Analysis Result에 Element, Node 정보 매칭
    beam_rot_data = pd.merge(beam_rot_data, element_data, how='left')
    beam_rot_data = pd.merge(beam_rot_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
    
    # 필요없는 부재 빼기, 필요한 부재만 추출
    beam_rot_data = beam_rot_data[beam_rot_data['Distance from I-End'] == 0]
    beam_rot_data = beam_rot_data[beam_rot_data['Property Name'].notna()]
     
    # 필요한 정보들만 다시 모아서 new dataframe
    column_name_to_slice = ['Group Name', 'Element Name', 'Property Name', 'V', 'Load Case', 'Step Type', 'H2 Rotation(rad)', 'H3 Rotation(rad)']
    beam_rot_data = beam_rot_data.loc[:, column_name_to_slice]    
    beam_rot_data.reset_index(inplace=True, drop=True)
    
#%% 성능기준(LS, CP) 정리해서 merge
    
    beam_rot_data = pd.merge(beam_rot_data, deformation_cap, how='left', left_on='Property Name', right_on='Name')
    
    beam_rot_data['DE_X Rotation(rad)'] = beam_rot_data['H2 Rotation(rad)'].abs() / beam_rot_data['LS(X)']
    beam_rot_data['DE_Y Rotation(rad)'] = beam_rot_data['H3 Rotation(rad)'].abs() / beam_rot_data['LS(Y)']
    beam_rot_data['MCE_X Rotation(rad)'] = beam_rot_data['H2 Rotation(rad)'].abs() / beam_rot_data['CP(X)']
    beam_rot_data['MCE_Y Rotation(rad)'] = beam_rot_data['H3 Rotation(rad)'].abs() / beam_rot_data['CP(Y)']
    
    beam_rot_data = beam_rot_data[beam_rot_data['Name'].notna()]
    
#%% 조작용 코드
    # 없애고 싶은 부재의 이름 입력(error_beam 확인 후!, DE, MCE에서 다 없어짐)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('AC405_1_17F'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('B15_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB4A_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB4B_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB3D_'))].index)

#%% DE 결과 Plot
    count = 1
    if len(DE_load_name_list) != 0:
        
        beam_rot_data_total_DE = pd.DataFrame()    
        
        for load_name in DE_load_name_list:
        
            temp_df_X_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['DE_X Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_DE['{}_X_max'.format(load_name)] = temp_df_X_max.tolist()
            
            temp_df_X_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['DE_X Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_DE['{}_X_min'.format(load_name)] = temp_df_X_min.tolist()
            
            temp_df_Y_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['DE_Y Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_DE['{}_Y_max'.format(load_name)] = temp_df_Y_max.tolist()
            
            temp_df_Y_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['DE_Y Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_DE['{}_Y_min'.format(load_name)] = temp_df_Y_min.tolist()
            
        beam_rot_data_total_DE['Element Name'] = temp_df_X_max.index
        
        beam_rot_data_total_DE.reset_index(inplace=True, drop=True)
        
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, element_data, how='left')
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, story_info, how='left', left_on='V', right_on='Height(mm)')
        beam_rot_data_total_DE.sort_values('Height(mm)', inplace=True)
        # beam_rot_data_total_DE.reset_index(inplace=True, drop=True)
        
        # 평균 열 생성        
        beam_rot_data_total_DE['DE_X Max avg'] = beam_rot_data_total_DE.iloc[:,list(range(0,len(DE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_DE['DE_X Min avg'] = beam_rot_data_total_DE.iloc[:,list(range(1,len(DE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_DE['DE_Y Max avg'] = beam_rot_data_total_DE.iloc[:,list(range(2,len(DE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_DE['DE_Y Min avg'] = beam_rot_data_total_DE.iloc[:,list(range(3,len(DE_load_name_list)*4,4))].mean(axis=1)
    
        # 전체 Plot            
        ### DE X ###
        fig1 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        plt.scatter(beam_rot_data_total_DE['DE_X Max avg'], beam_rot_data_total_DE.loc[:,'V'], color='k', s=1)
        plt.scatter(beam_rot_data_total_DE['DE_X Min avg'], beam_rot_data_total_DE.loc[:,'V'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Column Rotation (X DE)')
        
        plt.tight_layout()   
        plt.close()

        # 기준 넘는 점 확인
        error_beam_DE_X = beam_rot_data_total_DE[['Element Name', 'Property Name', 'Story Name', 'DE_X Max avg', 'DE_X Min avg']]\
                      [(beam_rot_data_total_DE['DE_X Max avg'] >= DCR_criteria) | (beam_rot_data_total_DE['DE_X Min avg'] >= DCR_criteria)]
        
        count += 1
        
        yield fig1
        yield error_beam_DE_X
        
        ### DE Y ###
        fig2 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        plt.scatter(beam_rot_data_total_DE['DE_Y Max avg'], beam_rot_data_total_DE.loc[:,'V'], color='k', s=1)
        plt.scatter(beam_rot_data_total_DE['DE_Y Min avg'], beam_rot_data_total_DE.loc[:,'V'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Column Rotation (Y DE)')
        
        plt.tight_layout()   
        plt.close()

        # 기준 넘는 점 확인
        error_beam_DE_Y = beam_rot_data_total_DE[['Element Name', 'Property Name', 'Story Name', 'DE_Y Max avg', 'DE_Y Min avg']]\
                      [(beam_rot_data_total_DE['DE_Y Max avg'] >= DCR_criteria) | (beam_rot_data_total_DE['DE_Y Min avg'] >= DCR_criteria)]
        
        count += 1
        
        yield fig2
        yield error_beam_DE_Y
        yield 'DE' # Marker 출력
        
#%% MCE 결과 Plot
    
    if len(MCE_load_name_list) != 0:
        
        beam_rot_data_total_MCE = pd.DataFrame()    
        
        for load_name in MCE_load_name_list:
        
            temp_df_X_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['MCE_X Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_MCE['{}_X_max'.format(load_name)] = temp_df_X_max.tolist()
            
            temp_df_X_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['MCE_X Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_MCE['{}_X_min'.format(load_name)] = temp_df_X_min.tolist()
            
            temp_df_Y_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['MCE_Y Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_MCE['{}_Y_max'.format(load_name)] = temp_df_Y_max.tolist()
            
            temp_df_Y_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['MCE_Y Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_MCE['{}_Y_min'.format(load_name)] = temp_df_Y_min.tolist()
            
        beam_rot_data_total_MCE['Element Name'] = temp_df_X_max.index
        
        beam_rot_data_total_MCE.reset_index(inplace=True, drop=True)
        
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, element_data, how='left')
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, story_info, how='left', left_on='V', right_on='Height(mm)')
        beam_rot_data_total_MCE.sort_values('Height(mm)', inplace=True)
        # beam_rot_data_total_MCE.reset_index(inplace=True, drop=True)
        
        # 평균 열 생성        
        beam_rot_data_total_MCE['MCE_X Max avg'] = beam_rot_data_total_MCE.iloc[:,list(range(0,len(MCE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_MCE['MCE_X Min avg'] = beam_rot_data_total_MCE.iloc[:,list(range(1,len(MCE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_MCE['MCE_Y Max avg'] = beam_rot_data_total_MCE.iloc[:,list(range(2,len(MCE_load_name_list)*4,4))].mean(axis=1)
        beam_rot_data_total_MCE['MCE_Y Min avg'] = beam_rot_data_total_MCE.iloc[:,list(range(3,len(MCE_load_name_list)*4,4))].mean(axis=1)     

        # 전체 Plot 
        ### MCE X ###
        fig3 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        # 평균 plot
        plt.scatter(beam_rot_data_total_MCE['MCE_X Max avg'], beam_rot_data_total_MCE.loc[:,'V'], color='k', s=1)
        plt.scatter(beam_rot_data_total_MCE['MCE_X Min avg'], beam_rot_data_total_MCE.loc[:,'V'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Column Rotation (X MCE)')
        
        plt.tight_layout()
        plt.close()
    
        # 기준 넘는 점 확인    
        error_beam_MCE_X = beam_rot_data_total_MCE[['Element Name', 'Property Name', 'Story Name', 'MCE_X Max avg', 'MCE_X Min avg']]\
                      [(beam_rot_data_total_MCE['MCE_X Max avg'] >= DCR_criteria) | (beam_rot_data_total_MCE['MCE_X Min avg'] >= DCR_criteria)]
        
        count += 1
        
        yield fig3
        yield error_beam_MCE_X
        
        ### MCE X ###
        fig4 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        # 평균 plot
        plt.scatter(beam_rot_data_total_MCE['MCE_Y Max avg'], beam_rot_data_total_MCE.loc[:,'V'], color='k', s=1)
        plt.scatter(beam_rot_data_total_MCE['MCE_Y Min avg'], beam_rot_data_total_MCE.loc[:,'V'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x= DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Column Rotation (Y MCE)')
        
        plt.tight_layout()
        plt.close()
    
        # 기준 넘는 점 확인    
        error_beam_MCE_Y = beam_rot_data_total_MCE[['Element Name', 'Property Name', 'Story Name', 'MCE_Y Max avg', 'MCE_Y Min avg']]\
                      [(beam_rot_data_total_MCE['MCE_Y Max avg'] >= DCR_criteria) | (beam_rot_data_total_MCE['MCE_Y Min avg'] >= DCR_criteria)]
        
        count += 1
        
        yield fig4
        yield error_beam_MCE_Y
        yield 'MCE' # Marker 출력

#%% General Column SF (DCR)
def CSF(self, input_xlsx_path, DCR_criteria=1, yticks=2, xlim=3):
    ''' 

    Perform-3D 해석 결과에서 일반기둥의 축력, 전단력을 불러와 Results_G.Column 엑셀파일을 작성. \n
    result_path : Perform-3D에서 나온 해석 파일의 경로. \n
    result_xlsx : Perform-3D에서 나온 해석 파일의 이름. 해당 파일 이름이 포함된 파일들을 모두 불러온다. \n
    input_path : Data Conversion 엑셀 파일의 경로 \n
    input_xlsx : Data Conversion 엑셀 파일의 이름. result_xlsx와는 달리 확장자명(.xlsx)까지 기입해줘야한다. 하나의 파일만 불러온다. \n
    column_xlsx : Results_E.Column 엑셀 파일의 이름.확장자명(.xlsx)까지 기입해줘야한다. \n
    export_to_pdf : 입력된 값에 따른 각 부재들의 결과 시트를 pdf로 출력. True = pdf 출력, False = pdf 미출력(Results_E.Column 엑셀파일만 작성됨).
    pdf_name = 출력할 pdf 파일 이름.
    
    '''
#%% Load Data
    # Data Conversion Sheets
    story_info = self.story_info
    deform_cap = self.col_deform_cap

    # Analysis Result Sheets
    node_data = self.node_data
    element_data = self.frame_data
    SF_info_data = self.beam_shear_force_data

    # Seismic Loads List
    load_name_list = self.load_name_list
    gravity_load_name = self.gravity_load_name
    seismic_load_name_list = self.seismic_load_name_list
    DE_load_name_list = self.DE_load_name_list
    MCE_load_name_list = self.MCE_load_name_list

#%% Process Data
    # 필요한 부재만 선별
    prop_name = deform_cap.iloc[:,0]
    prop_name.name = 'Property Name'
    element_data = element_data[element_data['Property Name'].isin(prop_name)]

    element_data = element_data.drop_duplicates()
    
    # Analysis Result에 Element, Node 정보 매칭
    element_data = pd.merge(element_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
    SF_ongoing = pd.merge(element_data, SF_info_data, how='left')
    SF_ongoing.reset_index(inplace=True, drop=True)

#%% V, M값에 절대값, 최대값, 평균값 뽑기

    # 절대값
    SF_ongoing[['V2 I-End', 'V3 I-End']] = SF_ongoing[['V2 I-End', 'V3 I-End']].abs()

    # V2, V3의 최대값을 저장하기 위해 필요한 데이터 slice
    SF_ongoing_max = SF_ongoing.iloc[[2*x for x in range(int(SF_ongoing.shape[0]/2))]]
    SF_ongoing_max = SF_ongoing_max.loc[:,['Element Name', 'Property Name', 'V', 'Load Case']]
    # [2*x for x in range(int(SF_ongoing.shape[0]/2] -> [짝수 index]
    
    # V2, V3의 최대값을 저장
    SF_ongoing_max['V2 max'] = SF_ongoing.groupby(SF_ongoing.index // 2)['V2 I-End'].max().tolist()
    SF_ongoing_max['V3 max'] = SF_ongoing.groupby(SF_ongoing.index // 2)['V3 I-End'].max().tolist()

    # 필요한 하중만 포함된 데이터 slice (MCE)
    SF_ongoing_max_MCE = SF_ongoing_max[SF_ongoing_max['Load Case']\
                                        .str.contains('|'.join(MCE_load_name_list))]
    SF_ongoing_max_G = SF_ongoing_max[SF_ongoing_max['Load Case']\
                                      .str.contains('|'.join(gravity_load_name))]
    # function equivalent of a combination of df.isin() and df.str.contains()
    
    # 부재별(Element Name) 평균값을 저장하기 위해 필요한 데이터프레임 생성
    SF_ongoing_max_avg = SF_ongoing_max_MCE.iloc[:,[0,1,2]]
    SF_ongoing_max_avg = SF_ongoing_max_avg.drop_duplicates()
    SF_ongoing_max_avg.set_index('Element Name', inplace=True)    
    # 부재별(Element Name) 평균값 뽑기
    SF_ongoing_max_avg['V2 max(MCE)'] = SF_ongoing_max_MCE.groupby(['Element Name'])['V2 max'].mean()
    SF_ongoing_max_avg['V2 max(G)'] = SF_ongoing_max_G.groupby(['Element Name'])['V2 max'].mean()
    SF_ongoing_max_avg['V3 max(MCE)'] = SF_ongoing_max_MCE.groupby(['Element Name'])['V3 max'].mean()
    SF_ongoing_max_avg['V3 max(G)'] = SF_ongoing_max_G.groupby(['Element Name'])['V3 max'].mean()
 
    
    # 이름별(Property Name) 최대값을 저장하기 위해 필요한 데이터프레임 생성
    SF_ongoing_max_avg_max = SF_ongoing_max_avg.copy()
    SF_ongoing_max_avg_max = SF_ongoing_max_avg_max.drop_duplicates(subset=['Property Name'], ignore_index=True)
    SF_ongoing_max_avg_max.set_index('Property Name', inplace=True) 
    # 같은 부재(그러나 잘려있는) 경우(Property Name) 최대값 뽑기
    SF_ongoing_max_avg_max = pd.merge(SF_ongoing_max_avg_max
                                      , SF_ongoing_max_avg.groupby(['Property Name'])['V2 max(MCE)'].max()
                                      , left_on='Property Name', right_index=True, suffixes=('_before', '_after'))
    SF_ongoing_max_avg_max = pd.merge(SF_ongoing_max_avg_max
                                      , SF_ongoing_max_avg.groupby(['Property Name'])['V2 max(G)'].max()
                                      , left_on='Property Name', right_index=True, suffixes=('_before', '_after'))
    SF_ongoing_max_avg_max = pd.merge(SF_ongoing_max_avg_max
                                      , SF_ongoing_max_avg.groupby(['Property Name'])['V3 max(MCE)'].max()
                                      , left_on='Property Name', right_index=True, suffixes=('_before', '_after'))
    SF_ongoing_max_avg_max = pd.merge(SF_ongoing_max_avg_max
                                      , SF_ongoing_max_avg.groupby(['Property Name'])['V3 max(G)'].max()
                                      , left_on='Property Name', right_index=True, suffixes=('_before', '_after'))
    
    # MCE에 대해 1.2배, G에 대해 0.2배
    SF_ongoing_max_avg_max['V2 max(MCE)_after'] = SF_ongoing_max_avg_max['V2 max(MCE)_after'] * 1.2
    SF_ongoing_max_avg_max['V2 max(G)_after'] = SF_ongoing_max_avg_max['V2 max(G)_after'] * 0.2
    SF_ongoing_max_avg_max['V3 max(MCE)_after'] = SF_ongoing_max_avg_max['V3 max(MCE)_after'] * 1.2
    SF_ongoing_max_avg_max['V3 max(G)_after'] = SF_ongoing_max_avg_max['V3 max(G)_after'] * 0.2
    
    SF_ongoing_max_avg_max.reset_index(inplace=True, drop=False)

#%% 결과값 정리
    
    SF_output = pd.merge(prop_name, SF_ongoing_max_avg_max, how='left')
        
    SF_output = SF_output.dropna()
    SF_output.reset_index(inplace=True, drop=True)
        
    # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
    SF_output = SF_output.replace(np.nan, '', regex=True)
    
    # 기존 시트에 V값 넣기
    SF_output1 = SF_output.iloc[:,0]
    SF_output2 = SF_output.iloc[:,[6,7,8,9]]

#%% 출력 (Using win32com...)
    
    # Using win32com...
    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게

    wb = excel.Workbooks.Open(input_xlsx_path)
    ws = wb.Sheets('Results_G.Column')
    
    startrow, startcol = 5, 1    
    ws.Range(ws.Cells(startrow, startcol),\
              ws.Cells(startrow+SF_output1.shape[0]-1, startcol)).Value\
    = [[i] for i in SF_output1]
    
    startrow, startcol = 5, 18    
    ws.Range(ws.Cells(startrow, startcol),\
              ws.Cells(startrow+SF_output2.shape[0]-1,\
                      startcol+SF_output2.shape[1]-1)).Value\
    = list(SF_output2.itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능
    
    wb.Save()            
    # wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application
    

#%% Column Shear Force 결과(DCR) 불러오기

    input_data_raw = pd.ExcelFile(input_xlsx_path)
    input_data_sheets = pd.read_excel(input_data_raw, ['Results_G.Column'], skiprows=3)
    input_data_raw.close()
    
    SF_DCR = input_data_sheets['Results_G.Column'].iloc[:,[0,31,33]]
    
    SF_DCR.columns = ['Property Name', 'DCR_MCE(X)', 'DCR_MCE(Y)']

    # SF_output에 DCR값 merge(그래프 그릴 때 height(mm) 정보가 필요함)
    SF_output = pd.merge(SF_output, SF_DCR, how='left') 

#%% 그래프
    count = 1
    
    # ### DE 그래프
    # if len(DE_load_name_list) != 0:
        
    #     fig1 = plt.figure(count, dpi=150, figsize=(5,6))
    #     plt.xlim(0, xlim)
        
    #     plt.scatter(SWR_avg_total['DCR_DE_min'], SWR_avg_total['Height'], color='k', s=1)
    #     plt.scatter(SWR_avg_total['DCR_DE_max'], SWR_avg_total['Height'], color='k', s=1)
    #     plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
    #     plt.axvline(x = DCR_criteria, color='r', linestyle='--')
        
    #     # 기타
    #     plt.grid(linestyle='-.')
    #     plt.xlabel('D/C Ratios')
    #     plt.ylabel('Story')
    #     plt.title('Wall Rotation (DE)')
        
    #     plt.close()
    #     count += 1
                            
    #     yield fig1
    #     yield 'DE' # Marker 출력
        
    ### MCE 그래프
    if len(MCE_load_name_list) != 0:
        
        ### H1 MCE 그래프 ###
        fig3 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        plt.scatter(SF_output['DCR_MCE(X)'], SF_output['V'], color='k', s=1)
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        plt.axvline(x = DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Shear Strength (X MCE)')
        
        plt.close()
        count += 1
        
        yield fig3
        
        ### H2 MCE 그래프 ###
        fig4 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(0, xlim)
        
        plt.scatter(SF_output['DCR_MCE(Y)'], SF_output['V'], color='k', s=1)
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        plt.axvline(x = DCR_criteria, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('D/C Ratios')
        plt.ylabel('Story')
        plt.title('Shear Strength (Y MCE)')
        
        plt.close()
        count += 1
        
        yield fig4
        yield 'MCE' # Marker 출력