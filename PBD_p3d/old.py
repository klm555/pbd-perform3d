#%% Beam Rotation

def BR(input_xlsx_path, result_xlsx_path
       , m_hinge_group_name, **kwargs): 
    # arguments가 너무 많을 때, 함수를 사용할 때 직접 매개변수를 명시해주는 방식

#%% 변수 정리
    s_hinge_group_name = kwargs['s_hinge_group_name'] if 's_hinge_group_name' in kwargs.keys() else None
    m_cri_DE = kwargs['moment_cri_DE'] if 'moment_cri_DE' in kwargs.keys() else 0.015/1.2
    m_cri_MCE = kwargs['moment_cri_MCE'] if 'moment_cri_MCE' in kwargs.keys() else 0.03/1.2
    s_cri_DE = kwargs['shear_cri_DE'] if 'shear_cri_DE' in kwargs.keys() else 0.015/1.2
    s_cri_MCE = kwargs['shear_cri_MCE'] if 'shear_cri_MCE' in kwargs.keys() else 0.03/1.2
    yticks = kwargs['yticks'] if 'yticks' in kwargs.keys() else 3
    xlim = kwargs['xlim'] if 'xlim' in kwargs.keys() else 0.03

#%% Analysis Result 불러오기(BR)
    to_load_list = result_xlsx_path
    
    # Gage data
    gage_data = pd.read_excel(to_load_list[0], sheet_name='Gage Data - Beam Type'
                              , skiprows=[0, 2], header=0, usecols=[0, 2, 7, 9]) # usecols로 원하는 열만 불러오기
    
    BR_M_gage_data = gage_data[gage_data['Group Name'] == m_hinge_group_name]
    BR_S_gage_data = gage_data[gage_data['Group Name'] == s_hinge_group_name]
    
    
    # Gage result data
    result_data = pd.DataFrame()
    for i in to_load_list:
        result_data_temp = pd.read_excel(i, sheet_name='Gage Results - Beam Type'
                                         , skiprows=[0, 2], header=0, usecols=[0, 2, 5, 7, 8, 9])
        result_data = pd.concat([result_data, result_data_temp])
    
    result_data = result_data.sort_values(by=['Load Case', 'Element Name', 'Step Type']) # 지진파 순서가 섞여있을 때 sort
    
    result_data = result_data[(result_data['Load Case'].str.contains('DE')) | (result_data['Load Case'].str.contains('MCE'))]
    
    BR_M_result_data = result_data[result_data['Group Name'] == m_hinge_group_name]
    BR_S_result_data = result_data[result_data['Group Name'] == s_hinge_group_name]
    
#%% Node, Story 정보 불러오기
    
    # Node Coord data
    Node_coord_data = pd.read_excel(to_load_list[0], sheet_name='Node Coordinate Data'
                                    , skiprows=[0, 2], header=0, usecols=[1, 2, 3, 4])
    
    # Story Info data
    story_info_xlsx_sheet = 'Story Data'
    story_info = pd.read_excel(input_xlsx_path, sheet_name=story_info_xlsx_sheet, skiprows=3, usecols=[0, 1, 2], keep_default_na=False)
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    story_name = story_info.loc[:, 'Story Name']

#%% 지진파 이름 list 만들기
    load_name_list = []
    for i in result_data['Load Case'].drop_duplicates():
        new_i = i.split('+')[1]
        new_i = new_i.strip()
        load_name_list.append(new_i)
    
    gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
    seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]
    
    seismic_load_name_list.sort()
    
    DE_load_name_list = [x for x in load_name_list if 'DE' in x] # base shear로 사용할 지진파 개수 산정을 위함
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]
    
#%% 데이터 매칭 후 결과뽑기
    
    ### Gage data에서 Element Name, I-Node ID 불러와서 v좌표 match하기
    BR_M_gage_data = BR_M_gage_data[['Element Name', 'I-Node ID']]; BR_M_gage_num = len(BR_M_gage_data) # gage 개수 얻기
    BR_S_gage_data = BR_S_gage_data[['Element Name', 'I-Node ID']]; BR_S_gage_num = len(BR_S_gage_data) # gage 개수 얻기
    
    # I-Node의 v좌표 match해서 추가
    gage_data = gage_data.join(Node_coord_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')
    BR_M_gage_data = BR_M_gage_data.join(Node_coord_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')
    BR_S_gage_data = BR_S_gage_data.join(Node_coord_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')
    BR_S_gage_data.reset_index(drop=True, inplace=True)
    
    ### BR_total data 만들기
    BR_M_max = BR_M_result_data[(BR_M_result_data['Step Type'] == 'Max') & (BR_M_result_data['Performance Level'] == 1)][['Rotation']].values # dataframe을 array로
    BR_M_max = BR_M_max.reshape(BR_M_gage_num, len(DE_load_name_list)+len(MCE_load_name_list), order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
    BR_M_max = pd.DataFrame(BR_M_max) # array를 다시 dataframe으로
    BR_M_min = BR_M_result_data[(BR_M_result_data['Step Type'] == 'Min') & (BR_M_result_data['Performance Level'] == 1)][['Rotation']].values
    BR_M_min = BR_M_min.reshape(BR_M_gage_num, len(DE_load_name_list)+len(MCE_load_name_list), order='F')
    BR_M_min = pd.DataFrame(BR_M_min)
    BR_M_total = pd.concat([BR_M_max, BR_M_min], axis=1) # DE11_max~MCE72_max, DE11_min~MCE72_min 각각 28개씩

    BR_S_max = BR_S_result_data[(BR_S_result_data['Step Type'] == 'Max') & (BR_S_result_data['Performance Level'] == 1)][['Rotation']].values # dataframe을 array로
    BR_S_max = BR_S_max.reshape(BR_S_gage_num, len(DE_load_name_list)+len(MCE_load_name_list), order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
    BR_S_max = pd.DataFrame(BR_S_max) # array를 다시 dataframe으로
    BR_S_min = BR_S_result_data[(BR_S_result_data['Step Type'] == 'Min') & (BR_S_result_data['Performance Level'] == 1)][['Rotation']].values
    BR_S_min = BR_S_min.reshape(BR_S_gage_num, len(DE_load_name_list)+len(MCE_load_name_list), order='F')
    BR_S_min = pd.DataFrame(BR_S_min)
    BR_S_total = pd.concat([BR_S_max, BR_S_min], axis=1) # DE11_max~MCE72_max, DE11_min~MCE72_min 각각 28개씩
            
    ### BR_avg_data 만들기
    BR_M_DE_max_avg = BR_M_total.iloc[:, 0:len(DE_load_name_list)].mean(axis=1)
    BR_M_MCE_max_avg = BR_M_total.iloc[:, len(DE_load_name_list) : len(DE_load_name_list)+len(MCE_load_name_list)].mean(axis=1)
    BR_M_DE_min_avg = BR_M_total.iloc[:, len(DE_load_name_list)+len(MCE_load_name_list) : 2*len(DE_load_name_list)+len(MCE_load_name_list)].mean(axis=1)
    BR_M_MCE_min_avg = BR_M_total.iloc[:, 2*len(DE_load_name_list)+len(MCE_load_name_list) : 2*len(DE_load_name_list) + 2*len(MCE_load_name_list)].mean(axis=1)
    BR_M_avg_total = pd.concat([BR_M_gage_data.iloc[:,[2,3,4]], BR_M_DE_max_avg, BR_M_DE_min_avg, BR_M_MCE_max_avg, BR_M_MCE_min_avg], axis=1)
    BR_M_avg_total.columns = ['X', 'Y', 'Height', 'DE_max_avg', 'DE_min_avg', 'MCE_max_avg', 'MCE_min_avg']

    BR_S_DE_max_avg = BR_S_total.iloc[:, 0:len(DE_load_name_list)].mean(axis=1)
    BR_S_MCE_max_avg = BR_S_total.iloc[:, len(DE_load_name_list) : len(DE_load_name_list)+len(MCE_load_name_list)].mean(axis=1)
    BR_S_DE_min_avg = BR_S_total.iloc[:, len(DE_load_name_list)+len(MCE_load_name_list) : 2*len(DE_load_name_list)+len(MCE_load_name_list)].mean(axis=1)
    BR_S_MCE_min_avg = BR_S_total.iloc[:, 2*len(DE_load_name_list)+len(MCE_load_name_list) : 2*len(DE_load_name_list) + 2*len(MCE_load_name_list)].mean(axis=1)
    BR_S_avg_total = pd.concat([BR_S_gage_data.iloc[:,[2,3,4]], BR_S_DE_max_avg, BR_S_DE_min_avg, BR_S_MCE_max_avg, BR_S_MCE_min_avg], axis=1, ignore_index=True)
    BR_S_avg_total.columns = ['X', 'Y', 'Height', 'DE_max_avg', 'DE_min_avg', 'MCE_max_avg', 'MCE_min_avg']
    
#%% ***조작용 코드
    # 없애고 싶은 부재의 x좌표 입력
    # BR_M_avg_total = BR_M_avg_total.drop(BR_M_avg_total[(BR_M_avg_total['X'] == -2700)].index)
    # BR_M_avg_total = BR_M_avg_total.drop(BR_M_avg_total[(BR_M_avg_total['X'] == -6.1e-05)].index)
    # BR_M_avg_total = BR_M_avg_total.drop(BR_M_avg_total[(BR_M_avg_total['X'] == -4725)].index)
    
#%% BR (Moment Hinge) 그래프

    count = 1

    if BR_M_avg_total.shape[0] != 0:

# DE 그래프        
        if len(DE_load_name_list) != 0:
        
            fig1 = plt.figure(count, figsize=(4,5), dpi=150)  # 그래프 사이즈
            plt.xlim(-xlim, xlim)
            
            plt.scatter(BR_M_avg_total['DE_min_avg'], BR_M_avg_total['Height'], color = 'k', s=1) # s=1 : point size
            plt.scatter(BR_M_avg_total['DE_max_avg'], BR_M_avg_total['Height'], color = 'k', s=1)
            
            # height값에 대응되는 층 이름으로 y축 눈금 작성
            plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
            
            # reference line 그려서 허용치 나타내기
            plt.axvline(x= -m_cri_DE, color='r', linestyle='--')
            plt.axvline(x= m_cri_DE, color='r', linestyle='--')
            
            plt.grid(linestyle='-.')
            plt.xlabel('Rotation(rad)')
            plt.ylabel('Story')
            plt.title('DE (Moment Hinge)')
            
            plt.close()
            count += 1
            
            error_coord_M_DE = BR_M_avg_total[(BR_M_avg_total['DE_max_avg'] >= m_cri_DE)\
                                              | (BR_M_avg_total['DE_min_avg'] <= -m_cri_DE)]
                
            yield fig1
            yield error_coord_M_DE
            yield 'DE' # Marker 출력
            

# MCE 그래프    
        if len(MCE_load_name_list) != 0:
    
            fig2 = plt.figure(count, figsize=(4,5), dpi=150)
            plt.xlim(-xlim, xlim)
            plt.scatter(BR_M_avg_total['MCE_min_avg'], BR_M_avg_total['Height'], color = 'k', s=1)
            plt.scatter(BR_M_avg_total['MCE_max_avg'], BR_M_avg_total['Height'], color = 'k', s=1)
            
            plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
            
            plt.axvline(x= -m_cri_MCE, color='r', linestyle='--')
            plt.axvline(x= m_cri_MCE, color='r', linestyle='--')
            
            plt.grid(linestyle='-.')
            plt.xlabel('Rotation(rad)')
            plt.ylabel('Story')
            plt.title('MCE (Moment Hinge)')
            
            plt.close()
            count += 1
            
            error_coord_M_MCE = BR_M_avg_total[(BR_M_avg_total['MCE_max_avg'] >= m_cri_MCE)\
                                               | (BR_M_avg_total['MCE_min_avg'] <= -m_cri_MCE)]
    
            yield fig2
            yield error_coord_M_MCE
            yield 'MCE' # Marker 출력
            
#%% BR(Shear Hinge) 그래프

    if BR_S_avg_total.shape[0] != 0:

# DE 그래프        
        if len(DE_load_name_list) != 0:

            fig3 = plt.figure(count, figsize=(4,5), dpi=150)  # 그래프 사이즈
            plt.xlim(-xlim, xlim)
            
            plt.scatter(BR_S_avg_total['DE_min_avg'], BR_S_avg_total['Height'], color = 'k', s=1) # s=1 : point size
            plt.scatter(BR_S_avg_total['DE_max_avg'], BR_S_avg_total['Height'], color = 'k', s=1)
            
            # height값에 대응되는 층 이름으로 y축 눈금 작성
            plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
            
            # reference line 그려서 허용치 나타내기
            # plt.axvline(x= -s_cri_DE, color='r', linestyle='--')
            # plt.axvline(x= s_cri_DE, color='r', linestyle='--')
            
            plt.grid(linestyle='-.')
            plt.xlabel('Rotation(rad)')
            plt.ylabel('Story')
            plt.title('DE (Shear Hinge)')
            
            plt.close()
            count += 1
            
            error_coord_S_DE = BR_S_avg_total[(BR_S_avg_total['DE_max_avg'] >= s_cri_DE)\
                                              | (BR_S_avg_total['DE_min_avg'] <= -s_cri_DE)]
    
            yield fig3
            yield error_coord_S_DE
            yield 'DE' # Marker 출력

# MCE 그래프    
        if len(MCE_load_name_list) != 0:
            
            fig4 = plt.figure(count, figsize=(4,5), dpi=150)
            plt.xlim(-xlim, xlim)
            plt.scatter(BR_S_avg_total['MCE_min_avg'], BR_S_avg_total['Height'], color = 'k', s=1)
            plt.scatter(BR_S_avg_total['MCE_max_avg'], BR_S_avg_total['Height'], color = 'k', s=1)
            
            plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
            
            # plt.axvline(x= -s_cri_MCE, color='r', linestyle='--')
            # plt.axvline(x= s_cri_MCE, color='r', linestyle='--')
            
            plt.grid(linestyle='-.')
            plt.xlabel('Rotation(rad)')
            plt.ylabel('Story')
            plt.title('MCE (Shear Hinge)')
            
            plt.close()
            count += 1
            
            error_coord_S_MCE = BR_S_avg_total[(BR_S_avg_total['MCE_max_avg'] >= s_cri_MCE)\
                                               | (BR_S_avg_total['MCE_min_avg'] <= -s_cri_MCE)]

            yield fig4
            yield error_coord_S_MCE
            yield 'MCE' # Marker 출력

#%% Beam Rotation (N) GAGE)

def BR_no_gage(input_xlsx_path, result_xlsx_path, cri_DE=0.01
               , cri_MCE=0.025/1.2, yticks=2, xlim=0.04):

#%% Input Sheets 정보 load
    
    story_info = pd.DataFrame()
    deformation_cap = pd.DataFrame()
    
    input_data_raw = pd.ExcelFile(input_xlsx_path)
    input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', 'Output_C.Beam Properties'], skiprows=3)
    input_data_raw.close()
    
    story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
    deformation_cap = input_data_sheets['Output_C.Beam Properties'].iloc[:,[0,48,49]]
    
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    deformation_cap.columns = ['Name', 'LS', 'CP']
    
#%% Analysis Result 불러오기
    
    to_load_list = result_xlsx_path
    
    beam_rot_data = pd.DataFrame()
    
    for i in to_load_list:
        result_data_raw = pd.ExcelFile(i)
        result_data_sheets = pd.read_excel(result_data_raw, ['Frame Results - Bending Deform', 'Node Coordinate Data',\
                                                         'Element Data - Frame Types'], skiprows=2)
        
        beam_rot_data_temp = result_data_sheets['Frame Results - Bending Deform'].iloc[:,[0,2,5,7,10,13,14]]
        beam_rot_data = pd.concat([beam_rot_data, beam_rot_data_temp])
        
    node_data = result_data_sheets['Node Coordinate Data'].iloc[:,[1,4]]
    
    element_data = result_data_sheets['Element Data - Frame Types'].iloc[:,[2,5,7]] # beam의 양 nodes중 한 node에서의 rotation * 2
    
                
    beam_rot_data.columns = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'Distance from I-End', 'H2 Rotation(rad)', 'H3 Rotation(rad)']
    node_data.columns = ['Node ID', 'V(mm)']
    element_data.columns = ['Element Name', 'Property Name', 'I-Node ID']
    
#%% temporary ((L), (R) 등 지우기)
    element_data.loc[:, 'Property Name'] = element_data.loc[:, 'Property Name'].str.split('(').str[0]
    
    #%% 필요없는 부재 빼기, 필요한 부재만 추출
    
    beam_rot_data = beam_rot_data[beam_rot_data['Distance from I-End'] == 0]
    
#%% Analysis Result에 Element, Node 정보 매칭
    
    element_data = element_data.drop_duplicates()
    node_data = node_data.drop_duplicates()
    
    beam_rot_data = pd.merge(beam_rot_data, element_data, how='left')
    beam_rot_data = pd.merge(beam_rot_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
    
    beam_rot_data = beam_rot_data[beam_rot_data['Property Name'].notna()]
    
    beam_rot_data.reset_index(inplace=True, drop=True)
    
#%% beam_rot_data의 값 수정(H1, H2 방향 중 major한 방향의 rotation값만 추출, 그리고 2배)
    major_rot = []
    for i, j in zip(beam_rot_data['H2 Rotation(rad)'], beam_rot_data['H3 Rotation(rad)']):
        if abs(i) >= abs(j):
            major_rot.append(i)
        else: major_rot.append(j)
    
    beam_rot_data['Major Rotation(rad)'] = major_rot
     
    # 필요한 정보들만 다시 모아서 new dataframe
    beam_rot_data = beam_rot_data.iloc[:, [0,1,7,10,2,3,11]]
    
#%% 지진파 이름 list 만들기
    load_name_list = []
    for i in beam_rot_data['Load Case'].drop_duplicates():
        new_i = i.split('+')[1]
        new_i = new_i.strip()
        load_name_list.append(new_i)
    
    gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
    seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]
    
    seismic_load_name_list.sort()
    
    DE_load_name_list = [x for x in load_name_list if 'DE' in x]
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]
    
#%% 성능기준(LS, CP) 정리해서 merge
    
    beam_rot_data = pd.merge(beam_rot_data, deformation_cap, how='left', left_on='Property Name', right_on='Name')
    
    beam_rot_data['DE Rotation(rad)'] = beam_rot_data['Major Rotation(rad)']
    beam_rot_data['MCE Rotation(rad)'] = beam_rot_data['Major Rotation(rad)']
    
    beam_rot_data = beam_rot_data[beam_rot_data['Name'].notna()]
    
#%% 조작용 코드
    # 없애고 싶은 부재의 이름 입력(error_beam 확인 후!, DE, MCE에서 다 없어짐)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('B11_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('B15_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB4A_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB4B_'))].index)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('WB3D_'))].index)

#%% DE 결과 Plot
    count = 1
    
    if len(DE_load_name_list) != 0:
        
        beam_rot_data_total_DE = pd.DataFrame()    
        
        for load_name in DE_load_name_list:
        
            temp_df_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['DE Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']
                          
            beam_rot_data_total_DE['{}_max'.format(load_name)] = temp_df_max.tolist()
            
            temp_df_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['DE Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']
                          
            beam_rot_data_total_DE['{}_min'.format(load_name)] = temp_df_min.tolist()
            
        beam_rot_data_total_DE['Element Name'] = temp_df_max.index
        
        beam_rot_data_total_DE.reset_index(inplace=True, drop=True)
        
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, element_data, how='left')
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, story_info, how='left', left_on='V(mm)', right_on='Height(mm)')
        beam_rot_data_total_DE.sort_values('Height(mm)', inplace=True)
        # beam_rot_data_total_DE.reset_index(inplace=True, drop=True)
        
    # 평균 열 생성
        
        beam_rot_data_total_DE['DE Max avg'] = beam_rot_data_total_DE.iloc[:,list(range(0,len(DE_load_name_list)*2,2))].mean(axis=1)
        beam_rot_data_total_DE['DE Min avg'] = beam_rot_data_total_DE.iloc[:,list(range(1,len(DE_load_name_list)*2,2))].mean(axis=1)
        
    # 전체 Plot
            
        ### DE 
        fig1 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(-xlim, xlim)
        
        plt.scatter(beam_rot_data_total_DE['DE Max avg'], beam_rot_data_total_DE.loc[:,'V(mm)'], color='k', s=1)
        plt.scatter(beam_rot_data_total_DE['DE Min avg'], beam_rot_data_total_DE.loc[:,'V(mm)'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x= cri_DE, color='r', linestyle='--')
        plt.axvline(x= -cri_DE, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('Rotation(rad)')
        plt.ylabel('Story')
        plt.title('Beam Rotation (DE)')
        
        plt.tight_layout()   
        plt.close()

    # 기준 넘는 점 확인
        error_beam_DE = beam_rot_data_total_DE[['Element Name', 'Property Name', 'Story Name', 'DE Max avg', 'DE Min avg']]\
                      [(beam_rot_data_total_DE['DE Max avg'].abs() >= cri_DE) | (beam_rot_data_total_DE['DE Min avg'].abs() >= cri_DE)]
        
        count += 1
        
        yield fig1
        yield error_beam_DE
        
#%% MCE 결과 Plot
    
    if len(MCE_load_name_list) != 0:
        
        beam_rot_data_total_MCE = pd.DataFrame()    
        
        for load_name in MCE_load_name_list:
        
            temp_df_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['MCE Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']
                          
            beam_rot_data_total_MCE['{}_max'.format(load_name)] = temp_df_max.tolist()
            
            temp_df_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['MCE Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']
                          
            beam_rot_data_total_MCE['{}_min'.format(load_name)] = temp_df_min.tolist()
            
        beam_rot_data_total_MCE['Element Name'] = temp_df_max.index
        
        beam_rot_data_total_MCE.reset_index(inplace=True, drop=True)
        
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, element_data, how='left')
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, story_info, how='left', left_on='V(mm)', right_on='Height(mm)')
        beam_rot_data_total_MCE.sort_values('Height(mm)', inplace=True)
        # beam_rot_data_total_MCE.reset_index(inplace=True, drop=True)
        
    # 평균 열 생성
        
        beam_rot_data_total_MCE['MCE Max avg'] = beam_rot_data_total_MCE.iloc[:,list(range(0,len(MCE_load_name_list)*2,2))].mean(axis=1)
        beam_rot_data_total_MCE['MCE Min avg'] = beam_rot_data_total_MCE.iloc[:,list(range(1,len(MCE_load_name_list)*2,2))].mean(axis=1)
            

        # 전체 Plot 
        ### MCE 
        fig2 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(xlim, -xlim)
        
        # 평균 plot
        plt.scatter(beam_rot_data_total_MCE['MCE Max avg'], beam_rot_data_total_MCE.loc[:,'V(mm)'], color='k', s=1)
        plt.scatter(beam_rot_data_total_MCE['MCE Min avg'], beam_rot_data_total_MCE.loc[:,'V(mm)'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        plt.axvline(x=cri_MCE, color='r', linestyle='--')
        plt.axvline(x=-cri_MCE, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('Rotation(rad)')
        plt.ylabel('Story')
        plt.title('Beam Rotation (MCE)')
        
        plt.tight_layout()
        plt.close()
    
        # 기준 넘는 점 확인    
        error_beam_MCE = beam_rot_data_total_MCE[['Element Name', 'Property Name', 'Story Name', 'MCE Max avg', 'MCE Min avg']]\
                      [(beam_rot_data_total_MCE['MCE Max avg'].abs() >= cri_MCE) | (beam_rot_data_total_MCE['MCE Min avg'].abs() >= cri_MCE)]
        
        yield fig2
        yield error_beam_MCE

#%% Transfer Beam SF (DCR)

def trans_beam_SF(input_xlsx_path, result_xlsx_path):

#%% Input Sheet 정보 load
        
    story_info = pd.DataFrame()
    transfer_element_info = pd.DataFrame()

    input_xlsx_sheet = 'Output_E.Beam Properties'
    input_data_raw = pd.ExcelFile(input_xlsx_path)
    input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', input_xlsx_sheet], skiprows=3)
    input_data_raw.close()

    story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
    transfer_element_info = input_data_sheets[input_xlsx_sheet].iloc[:,0]
    story_info = story_info[::-1]
    story_info.reset_index(inplace=True, drop=True)

    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    transfer_element_info.name = 'Name'

    #%% Analysis Result 불러오기
    to_load_list = result_xlsx_path

    # 전단력 Data
    SF_info_data = pd.DataFrame()
    for i in to_load_list:
        SF_info_data_temp = pd.read_excel(i, sheet_name='Frame Results - End Forces'
                                          , skiprows=[0, 2], header=0
                                          , usecols=[0,2,5,7,10,11,17,18]) # usecols로 원하는 열만 불러오기
        SF_info_data = pd.concat([SF_info_data, SF_info_data_temp])

    SF_info_data = SF_info_data.sort_values(by=['Load Case', 'Element Name', 'Step Type']) # 지진파 순서가 섞여있을 때 sort

    # 부재 이름 Matching을 위한 Element 정보
    element_info_data = pd.DataFrame()
    for i in to_load_list:
        element_info_data_temp = pd.read_excel(i, sheet_name='Element Data - Frame Types'
                                               , skiprows=[0, 2], header=0, usecols=[0, 2, 5, 7]) # usecols로 원하는 열만 불러오기
        element_info_data = pd.concat([element_info_data, element_info_data_temp])

    # 필요한 부재만 선별
    element_info_data = element_info_data[element_info_data['Property Name'].isin(transfer_element_info)]
    
    # 층 정보 Matching을 위한 Node 정보
    height_info_data = pd.DataFrame()    
    for i in to_load_list:
        height_info_data_temp = pd.read_excel(i, sheet_name='Node Coordinate Data'
                                              , skiprows=[0, 2], header=0, usecols=[1, 4]) # usecols로 원하는 열만 불러오기
        height_info_data = pd.concat([height_info_data, height_info_data_temp])

    element_info_data = pd.merge(element_info_data, height_info_data, how='left', left_on='I-Node ID', right_on='Node ID')

    element_info_data = element_info_data.drop_duplicates()

    # 전단력, 부재 이름 Matching (by Element Name)
    SF_ongoing = pd.merge(element_info_data.iloc[:, [1,2,5]], SF_info_data.iloc[:, 1:], how='left')

    SF_ongoing = SF_ongoing.sort_values(by=['Element Name', 'Load Case', 'Step Type'])

    SF_ongoing.reset_index(inplace=True, drop=True)

    #%% 지진파 이름 list 만들기

    load_name_list = []
    for i in SF_ongoing['Load Case'].drop_duplicates():
        new_i = i.split('+')[1]
        new_i = new_i.strip()
        load_name_list.append(new_i)

    gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
    seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]

    seismic_load_name_list.sort()

    DE_load_name_list = [x for x in load_name_list if 'DE' in x]
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]

    #%% V, M값에 절대값, 최대값, 평균값 뽑기

    # 절대값, 1.2배
    SF_ongoing.iloc[:,[5,6,7,8]] = SF_ongoing.iloc[:,[5,6,7,8]].abs() * 1.2

    # i, j 노드 중 최대값 뽑기
    SF_ongoing['V2 max'] = SF_ongoing[['V2 I-End', 'V2 J-End']].max(axis = 1)
    SF_ongoing['M3 max'] = SF_ongoing[['M3 I-End', 'M3 J-End']].max(axis = 1)

    # max, min 중 최대값 뽑기
    SF_ongoing_max = SF_ongoing.loc[SF_ongoing.groupby(SF_ongoing.index // 2)['V2 max'].idxmax()]
    SF_ongoing_max['M3 max'] = SF_ongoing.groupby(SF_ongoing.index // 2)['M3 max'].max().tolist()

    # 필요한 하중만 포함된 데이터 slice (MCE)
    SF_ongoing_max = SF_ongoing_max[SF_ongoing_max['Load Case']\
                                    .str.contains('|'.join(MCE_load_name_list))] # function equivalent of a combination of df.isin() and df.str.contains()
    
    # 부재별 평균값 뽑기
    SF_ongoing_max_avg = SF_ongoing_max.iloc[:,[0,1,2]]
    SF_ongoing_max_avg = SF_ongoing_max_avg.drop_duplicates()
    SF_ongoing_max_avg.set_index('Element Name', inplace=True)
    
    SF_ongoing_max_avg['V2 max'] = SF_ongoing_max.groupby(['Element Name'])['V2 max'].mean()
    SF_ongoing_max_avg['M3 max'] = SF_ongoing_max.groupby(['Element Name'])['M3 max'].mean()
 
    # 같은 부재(그러나 잘려있는) 경우 최대값 뽑기
    SF_ongoing_max_avg_max = SF_ongoing_max_avg.loc[SF_ongoing_max_avg.groupby(['Property Name'])['V2 max'].idxmax()]
    SF_ongoing_max_avg_max['M3 max'] = SF_ongoing_max_avg.groupby(['Property Name'])['M3 max'].max().tolist()
    
    SF_ongoing_max_avg_max.reset_index(inplace=True, drop=True)

#%% 결과값 정리 후 input sheets에 넣기
    
    SF_ongoing_max_avg_max = pd.merge(transfer_element_info.rename('Property Name'),\
                                      SF_ongoing_max_avg_max, how='left')
        
    SF_ongoing_max_avg_max = SF_ongoing_max_avg_max.dropna()
    SF_ongoing_max_avg_max.reset_index(inplace=True, drop=True)
    
    # SF_ongoing_max_avg 재정렬
    SF_output = SF_ongoing_max_avg_max.iloc[:,[0,2,3]] 
    
    # nan인 칸을 ''로 바꿔주기 (win32com으로 nan입력시 임의의 숫자가 입력되기때문 ㅠ)
    SF_output = SF_output.replace(np.nan, '', regex=True)
    
    # Using win32com...
    # Call CoInitialize function before using any COM object
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize()) # 엑셀 실행
    excel.Visible = True # 엑셀창 안보이게

    wb = excel.Workbooks.Open(input_xlsx_path)
    ws = wb.Sheets('Results_E.Beam')
    
    startrow, startcol = 5, 1
    
    ws.Range(ws.Cells(startrow, startcol),\
             ws.Cells(startrow + SF_output.shape[0]-1,\
                      startcol + SF_output.shape[1]-1)).Value\
    = list(SF_output.itertuples(index=False, name=None)) # dataframe -> tuple list 형식만 입력가능
    
    wb.Close(SaveChanges=1) # Closing the workbook
    # excel.Quit() # Closing the application
    
    print('Done!')

#%% Column Rotation
def CR(input_xlsx_path, result_xlsx_path
       , col_group_name='G.Column', **kwargs):

#%% 변수 정리
    m_cri_DE = kwargs['shear_cri_DE'] if 'shear_cri_DE' in kwargs.keys() else 0.003
    m_cri_MCE = kwargs['shear_cri_MCE'] if 'shear_cri_MCE' in kwargs.keys() else 0.004/1.2
    yticks = kwargs['yticks'] if 'yticks' in kwargs.keys() else 3
    xlim = kwargs['xlim'] if 'xlim' in kwargs.keys() else 0.005
        
#%% Input Sheets 정보 load
    story_info = pd.DataFrame()
    
    input_data_raw = pd.ExcelFile(input_xlsx_path)
    input_data_sheets = pd.read_excel(input_data_raw, ['Story Data', 'Output_G.Column Properties'], skiprows=3)
    input_data_raw.close()
    
    story_info = input_data_sheets['Story Data'].iloc[:,[0,1,2]]
    
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    
#%% Analysis Result 불러오기   
    to_load_list = result_xlsx_path
    
    beam_rot_data = pd.DataFrame()
    
    for i in to_load_list:
        result_data_raw = pd.ExcelFile(i)
        result_data_sheets = pd.read_excel(result_data_raw, ['Frame Results - Bending Deform', 'Node Coordinate Data',\
                                                         'Element Data - Frame Types'], skiprows=2)
        
        beam_rot_data_temp = result_data_sheets['Frame Results - Bending Deform'].iloc[:,[0,2,5,7,10,13,14]]
        beam_rot_data = pd.concat([beam_rot_data, beam_rot_data_temp])
        
    node_data = result_data_sheets['Node Coordinate Data'].iloc[:,[1,4]]
    
    element_data = result_data_sheets['Element Data - Frame Types'].iloc[:,[2,5,7]] # beam의 양 nodes중 한 node에서의 rotation * 2
    
                
    beam_rot_data.columns = ['Group Name', 'Element Name', 'Load Case', 'Step Type', 'Distance from I-End', 'H2 Rotation(rad)', 'H3 Rotation(rad)']
    node_data.columns = ['Node ID', 'V(mm)']
    element_data.columns = ['Element Name', 'Property Name', 'I-Node ID']
    
#%% temporary ((L), (R) 등 지우기)
    element_data.loc[:, 'Property Name'] = element_data.loc[:, 'Property Name'].str.split('(').str[0]
    
    #%% 필요없는 부재 빼기, 필요한 부재만 추출
    beam_rot_data = beam_rot_data[beam_rot_data['Group Name'] == col_group_name]
    beam_rot_data = beam_rot_data[beam_rot_data['Distance from I-End'] == 0]
    
#%% Analysis Result에 Element, Node 정보 매칭    
    element_data = element_data.drop_duplicates()
    node_data = node_data.drop_duplicates()
    
    beam_rot_data = pd.merge(beam_rot_data, element_data, how='left')
    beam_rot_data = pd.merge(beam_rot_data, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
    
    beam_rot_data = beam_rot_data[beam_rot_data['Property Name'].notna()]
    
    beam_rot_data.reset_index(inplace=True, drop=True)
    
#%% beam_rot_data의 값 수정 (H1, H2 방향 중 major한 방향의 rotation값만 추출, 그리고 2배)
    major_rot = []
    for i, j in zip(beam_rot_data['H2 Rotation(rad)'], beam_rot_data['H3 Rotation(rad)']):
        if abs(i) >= abs(j):
            major_rot.append(i)
        else: major_rot.append(j)
    
    beam_rot_data['Major Rotation(rad)'] = major_rot
     
    # 필요한 정보들만 다시 모아서 new dataframe
    beam_rot_data = beam_rot_data.iloc[:, [0,1,7,10,2,3,5,6]]
    
#%% 지진파 이름 list 만들기
    load_name_list = []
    for i in beam_rot_data['Load Case'].drop_duplicates():
        new_i = i.split('+')[1]
        new_i = new_i.strip()
        load_name_list.append(new_i)
    
    gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
    seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]
    
    seismic_load_name_list.sort()
    
    DE_load_name_list = [x for x in load_name_list if 'DE' in x]
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]
    
#%% 조작용 코드
    # 없애고 싶은 부재의 이름 입력(error_beam 확인 후!, DE, MCE에서 다 없어짐)
    # beam_rot_data = beam_rot_data.drop(beam_rot_data[(beam_rot_data['Property Name'].str.contains('LB4_'))].index)
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
                          .groupby(['Element Name'])['H2 Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_DE['{}_X_max'.format(load_name)] = temp_df_X_max.tolist()
            
            temp_df_X_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['H2 Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_DE['{}_X_min'.format(load_name)] = temp_df_X_min.tolist()
            
            temp_df_Y_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['H3 Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_DE['{}_Y_max'.format(load_name)] = temp_df_Y_max.tolist()
            
            temp_df_Y_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['H3 Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_DE['{}_Y_min'.format(load_name)] = temp_df_Y_min.tolist()
            
        beam_rot_data_total_DE['Element Name'] = temp_df_X_max.index
        
        beam_rot_data_total_DE.reset_index(inplace=True, drop=True)
        
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, element_data, how='left')
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
        beam_rot_data_total_DE = pd.merge(beam_rot_data_total_DE, story_info, how='left', left_on='V(mm)', right_on='Height(mm)')
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
        plt.xlim(-xlim, xlim)
        
        plt.scatter(beam_rot_data_total_DE['DE_X Max avg'], beam_rot_data_total_DE.loc[:,'V(mm)'], color='k', s=1)
        plt.scatter(beam_rot_data_total_DE['DE_X Min avg'], beam_rot_data_total_DE.loc[:,'V(mm)'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        # plt.axvline(x= m_cri_DE, color='r', linestyle='--')
        # plt.axvline(x= -m_cri_DE, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('Rotation(rad)')
        plt.ylabel('Story')
        plt.title('Column Rotation (X DE)')
        
        plt.tight_layout()   
        plt.close()

        # 기준 넘는 점 확인
        error_beam_DE_X = beam_rot_data_total_DE[['Element Name', 'Property Name', 'Story Name', 'DE_X Max avg', 'DE_X Min avg']]\
                      [(beam_rot_data_total_DE['DE_X Max avg'] >= m_cri_DE) | (beam_rot_data_total_DE['DE_X Min avg'] <= -m_cri_DE)]
        
        count += 1
        
        yield fig1
        yield error_beam_DE_X
        
        ### DE Y ###
        fig2 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(-xlim, xlim)
        
        plt.scatter(beam_rot_data_total_DE['DE_Y Max avg'], beam_rot_data_total_DE.loc[:,'V(mm)'], color='k', s=1)
        plt.scatter(beam_rot_data_total_DE['DE_Y Min avg'], beam_rot_data_total_DE.loc[:,'V(mm)'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        # plt.axvline(x= m_cri_DE, color='r', linestyle='--')
        # plt.axvline(x= -m_cri_DE, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('Rotation(rad)')
        plt.ylabel('Story')
        plt.title('Column Rotation (Y DE)')
        
        plt.tight_layout()   
        plt.close()

        # 기준 넘는 점 확인
        error_beam_DE_Y = beam_rot_data_total_DE[['Element Name', 'Property Name', 'Story Name', 'DE_Y Max avg', 'DE_Y Min avg']]\
                      [(beam_rot_data_total_DE['DE_Y Max avg'] >= m_cri_DE) | (beam_rot_data_total_DE['DE_Y Min avg'] <= -m_cri_DE)]
        
        count += 1
        
        yield fig2
        yield error_beam_DE_Y
        
#%% MCE 결과 Plot
    
    if len(MCE_load_name_list) != 0:
        
        beam_rot_data_total_MCE = pd.DataFrame()    
        
        for load_name in MCE_load_name_list:
        
            temp_df_X_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['H2 Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_MCE['{}_X_max'.format(load_name)] = temp_df_X_max.tolist()
            
            temp_df_X_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['H2 Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_MCE['{}_X_min'.format(load_name)] = temp_df_X_min.tolist()
            
            temp_df_Y_max = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Max')]\
                          .groupby(['Element Name'])['H3 Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_MCE['{}_Y_max'.format(load_name)] = temp_df_Y_max.tolist()
            
            temp_df_Y_min = beam_rot_data[(beam_rot_data['Load Case'].str.contains('{}'.format(load_name)))\
                                        & (beam_rot_data['Step Type'] == 'Min')]\
                          .groupby(['Element Name'])['H3 Rotation(rad)']\
                          .agg(**{'Rotation avg':'mean'})['Rotation avg']                          
            beam_rot_data_total_MCE['{}_Y_min'.format(load_name)] = temp_df_Y_min.tolist()
            
        beam_rot_data_total_MCE['Element Name'] = temp_df_X_max.index
        
        beam_rot_data_total_MCE.reset_index(inplace=True, drop=True)
        
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, element_data, how='left')
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, node_data, how='left', left_on='I-Node ID', right_on='Node ID')
        beam_rot_data_total_MCE = pd.merge(beam_rot_data_total_MCE, story_info, how='left', left_on='V(mm)', right_on='Height(mm)')
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
        plt.xlim(-xlim, xlim)
        
        # 평균 plot
        plt.scatter(beam_rot_data_total_MCE['MCE_X Max avg'], beam_rot_data_total_MCE.loc[:,'V(mm)'], color='k', s=1)
        plt.scatter(beam_rot_data_total_MCE['MCE_X Min avg'], beam_rot_data_total_MCE.loc[:,'V(mm)'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        # plt.axvline(x= m_cri_MCE, color='r', linestyle='--')
        # plt.axvline(x= -m_cri_MCE, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('Rotation(rad)')
        plt.ylabel('Story')
        plt.title('Column Rotation (X MCE)')
        
        plt.tight_layout()
        plt.close()
    
        # 기준 넘는 점 확인    
        error_beam_MCE_X = beam_rot_data_total_MCE[['Element Name', 'Property Name', 'Story Name', 'MCE_X Max avg', 'MCE_X Min avg']]\
                      [(beam_rot_data_total_MCE['MCE_X Max avg'] >= m_cri_MCE) | (beam_rot_data_total_MCE['MCE_X Min avg'] <= -m_cri_MCE)]
        
        count += 1
        
        yield fig3
        yield error_beam_MCE_X
        
        ### MCE X ###
        fig4 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(-xlim, xlim)
        
        # 평균 plot
        plt.scatter(beam_rot_data_total_MCE['MCE_Y Max avg'], beam_rot_data_total_MCE.loc[:,'V(mm)'], color='k', s=1)
        plt.scatter(beam_rot_data_total_MCE['MCE_Y Min avg'], beam_rot_data_total_MCE.loc[:,'V(mm)'], color='k', s=1)
        
        plt.yticks(story_info['Height(mm)'][::-yticks], story_info['Story Name'][::-yticks])
        # plt.xticks(range(14), range(1,15))
        
        # plt.axvline(x= m_cri_MCE, color='r', linestyle='--')
        # plt.axvline(x= -m_cri_MCE, color='r', linestyle='--')
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('Rotation(rad)')
        plt.ylabel('Story')
        plt.title('Column Rotation (Y MCE)')
        
        plt.tight_layout()
        plt.close()
    
        # 기준 넘는 점 확인    
        error_beam_MCE_Y = beam_rot_data_total_MCE[['Element Name', 'Property Name', 'Story Name', 'MCE_Y Max avg', 'MCE_Y Min avg']]\
                      [(beam_rot_data_total_MCE['MCE_Y Max avg'] >= m_cri_MCE) | (beam_rot_data_total_MCE['MCE_Y Min avg'] <= -m_cri_MCE)]
        
        count += 1
        
        yield fig4
        yield error_beam_MCE_Y

#%% Shear Wall Rotation

def SWR(input_xlsx_path, result_xlsx_path, DE_criteria=0.002
        , MCE_criteria=0.004/1.2, yticks=2, xlim=0.005):
    ''' 

    각각의 벽체의 회전각을 산포도 그래프 형식으로 출력.
    
    Parameters
    ----------
    input_path : str
                 Data Conversion 엑셀 파일의 경로.
                 
    input_xlsx : str
                 Data Conversion 엑셀 파일의 이름. result_xlsx와는 달리 확장자명(.xlsx)까지 기입해줘야한다. 하나의 파일만 불러온다.
                 
    result_path : str
                  Perform-3D에서 나온 해석 파일의 경로.
                  
    result_xlsx : str, optional, default='Analysis Result'
                  Perform-3D에서 나온 해석 파일의 이름. 해당 파일 이름이 포함된 파일들을 모두 불러온다.
                 
    DE_criteria : float, optional, default=0.002/1.2
                  LS(인명안전)에 대한 벽체 회전각 허용기준. default=가장 보수적인 값
    
    MCE_criteria : float, optional, default=-0.004/1.2
                   CP(붕괴방지)에 대한 벽체 회전각 허용기준. default=가장 보수적인 값
                   
    yticks : int, optional, default=2
             그래프의 y축 눈금 간격(층간격). 층이 너무 높으면 y축에 너무 많은 층이 표기되기 때문에, 층간격을 조절해서 정돈된 그래프를 표기할 수 있다.

    xlim : int, optional, default=0.005
           그래프의 x축 limit 값. x축 limit 안의 값만 표기되므로, limit를 넘어가는 값을 확인하고 싶을 시에는 더 큰 xlim 값을 사용하면 된다.

    Yields
    -------
    Min, Max값 모두 출력됨. 
    
    fig1 : matplotlib.pyplot.figure or None
           DE(설계지진) 발생 시 벽체 회전각 그래프
    
    fig2 : matplotlib.pyplot.figure or None
           MCE(최대고려지진) 발생 시 벽체 회전각 그래프
    
    error_coord_DE : pandas.core.frame.DataFrame or None
                     DE(설계지진) 발생 시 기준값을 초과하는 벽체의 좌표
                     
    error_coord_MCE : pandas.core.frame.DataFrame or None
                     MCE(최대고려지진) 발생 시 기준값을 초과하는 벽체의 좌표                                          
    
    Raises
    -------
    
    References
    -------
    .. [1] "철근콘크리트 건축구조물의 성능기반 내진설계 지침", 대한건축학회, p.79, 2021    
    
    '''
    #%% Analysis Result 불러오기
    to_load_list = result_xlsx_path

    # Gage data
    gage_data = pd.read_excel(to_load_list[0], sheet_name='Gage Data - Wall Type'
                              , skiprows=[0, 2], header=0, usecols=[0, 2, 7, 9, 11, 13]) # usecols로 원하는 열만 불러오기

    # Gage result data
    wall_rot_data = pd.DataFrame()
    for i in to_load_list:
        wall_rot_data_temp = pd.read_excel(i, sheet_name='Gage Results - Wall Type', skiprows=[0,2])
        
        column_name_to_choose = ['Group Name', 'Element Name', 'Load Case'
                                  , 'Step Type', 'Rotation', 'Performance Level']
        wall_rot_data_temp = wall_rot_data_temp.loc[:,column_name_to_choose]
        wall_rot_data = pd.concat([wall_rot_data, wall_rot_data_temp])

    wall_rot_data.sort_values(['Load Case', 'Element Name'] , inplace=True)

    # Node Coord data
    node_data = pd.read_excel(to_load_list[0], sheet_name='Node Coordinate Data'
                              , skiprows=[0, 2], header=0, usecols=[1, 2, 3, 4])

    # Story Info data
    story_info_xlsx_sheet = 'Story Data'
    story_info = pd.read_excel(input_xlsx_path, sheet_name=story_info_xlsx_sheet, skiprows=3, usecols=[0, 1, 2], keep_default_na=False)
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    story_name = story_info.loc[:, 'Story Name']

#%% 지진파 이름 list 만들기
    load_name_list = []
    for i in wall_rot_data['Load Case'].drop_duplicates():
        new_i = i.split('+')[1]
        new_i = new_i.strip()
        load_name_list.append(new_i)
    
    gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
    seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]
    
    seismic_load_name_list.sort()
    
    DE_load_name_list = [x for x in load_name_list if 'DE' in x] # base shear로 사용할 지진파 개수 산정을 위함
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]

    #%% 데이터 매칭 후 결과뽑기
    
    wall_rot_data = wall_rot_data[wall_rot_data['Load Case']\
                                      .str.contains('|'.join(seismic_load_name_list))]
    
    ### Gage data에서 Element Name, I-Node ID 불러와서 v좌표 match하기
    gage_data = gage_data[['Element Name', 'I-Node ID']]; gage_num = len(gage_data) # gage 개수 얻기
    node_data_V = node_data[['Node ID', 'V']]

    # I-Node의 v좌표 match해서 추가
    gage_data = gage_data.join(node_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')

    ### SWR_total data 만들기

    SWR_max = wall_rot_data[(wall_rot_data['Step Type'] == 'Max') & (wall_rot_data['Performance Level'] == 1)][['Rotation']].values # dataframe을 array로
    SWR_max = SWR_max.reshape(gage_num, len(seismic_load_name_list), order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
    SWR_max = pd.DataFrame(SWR_max) # array를 다시 dataframe으로
    SWR_min = wall_rot_data[(wall_rot_data['Step Type'] == 'Min') & (wall_rot_data['Performance Level'] == 1)][['Rotation']].values
    SWR_min = SWR_min.reshape(gage_num, len(seismic_load_name_list), order='F')
    SWR_min = pd.DataFrame(SWR_min)
    SWR_total = pd.concat([SWR_max, SWR_min], axis=1) # DE11_max~MCE72_max, DE11_min~MCE72_min 각각 28개씩

    ### SWR_avg_data 만들기
    DE_max_avg = SWR_total.iloc[:, 0:len(DE_load_name_list)].mean(axis=1)
    MCE_max_avg = SWR_total.iloc[:, len(DE_load_name_list) : len(DE_load_name_list)+ len(MCE_load_name_list)].mean(axis=1)
    DE_min_avg = SWR_total.iloc[:, len(DE_load_name_list)+len(MCE_load_name_list) : 2*len(DE_load_name_list)+len(MCE_load_name_list)].mean(axis=1)
    MCE_min_avg = SWR_total.iloc[:, 2*len(DE_load_name_list)+len(MCE_load_name_list) : 2*len(DE_load_name_list) + 2*len(MCE_load_name_list)].mean(axis=1)
    SWR_avg_total = pd.concat([gage_data.loc[:,['H1', 'H2', 'V']], DE_max_avg, DE_min_avg, MCE_max_avg, MCE_min_avg], axis=1)
    SWR_avg_total.columns = ['X(mm)', 'Y(mm)', 'Height(mm)', 'DE_max_avg', 'DE_min_avg', 'MCE_max_avg', 'MCE_min_avg']

    #%% ***조작용 코드
    # SWR_avg_total = SWR_avg_total.drop(SWR_avg_total[(SWR_avg_total.iloc[:,2] < -0.0038) | (SWR_avg_total.iloc[:,1] > 0.0038)].index) # DE
    # SWR_avg_total = SWR_avg_total.drop(SWR_avg_total[(SWR_avg_total.iloc[:,4] < -0.0035) | (SWR_avg_total.iloc[:,3] > 0.0035)].index) # MCE

    #%% 그래프
    count = 1
    
    # DE 그래프
    if len(DE_load_name_list) != 0:
    
        fig1 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(-xlim, xlim)
        plt.scatter(SWR_avg_total['DE_min_avg'], SWR_avg_total['Height(mm)'], color = 'k', s=1) # s=1 : point size
        plt.scatter(SWR_avg_total['DE_max_avg'], SWR_avg_total['Height(mm)'], color = 'k', s=1)
    
        # height값에 대응되는 층 이름으로 y축 눈금 작성
        plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
    
        # reference line 그려서 허용치 나타내기
        # plt.axvline(x= -DE_criteria, color='r', linestyle='--')
        # plt.axvline(x= DE_criteria, color='r', linestyle='--')
    
        plt.grid(linestyle='-.')
        plt.xlabel('Rotation(rad)')
        plt.ylabel('Story')
        plt.title('Wall Rotation (DE)')
    
        plt.tight_layout()
        plt.close()
        count += 1
        
        error_coord_DE = SWR_avg_total[(SWR_avg_total['DE_max_avg'] >= DE_criteria)\
                                       | (SWR_avg_total['DE_min_avg'] <= -DE_criteria)]          
        
        yield fig1
        yield error_coord_DE
        yield 'DE' # Marker 출력

    # MCE 그래프
    if len(MCE_load_name_list) != 0:
        
        fig2 = plt.figure(count, dpi=150, figsize=(5,6))
        plt.xlim(-xlim, xlim)
        plt.scatter(SWR_avg_total['MCE_min_avg'], SWR_avg_total['Height(mm)'], color = 'k', s=1)
        plt.scatter(SWR_avg_total['MCE_max_avg'], SWR_avg_total['Height(mm)'], color = 'k', s=1)
    
        plt.yticks(story_info['Height(mm)'][::-yticks], story_name[::-yticks])
    
        # plt.axvline(x= -MCE_criteria, color='r', linestyle='--')
        # plt.axvline(x= MCE_criteria, color='r', linestyle='--')
    
        plt.grid(linestyle='-.')
        plt.xlabel('Rotation(rad)')
        plt.ylabel('Story')
        plt.title('Wall Rotation (MCE)')
    
        plt.tight_layout()
        plt.close()
        count += 1
        
        error_coord_MCE = SWR_avg_total[(SWR_avg_total['MCE_max_avg'] >= MCE_criteria)\
                                        | (SWR_avg_total['MCE_min_avg'] <= -MCE_criteria)]  
        
        yield fig2
        yield error_coord_MCE
        yield 'MCE' # Marker 출력

