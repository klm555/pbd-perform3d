import pandas as pd
import os
from collections import deque  # Double-ended Queue : 자료의 앞, 뒤 양 방향에서 자료를 추가하거나 제거가능
import matplotlib.pyplot as plt
from decimal import *
import io
import pickle
from collections import deque

#%% Base SF

def base_SF(self, ylim=70000) -> pd.DataFrame:
    ''' 

    Perform-3D 해석 결과에서 각 지진파에 대한 Base층의 전단력을 막대그래프 형식으로 출력. (kN)
    
    Parameters
    ----------                  
    ylim : int, optional, default=70000
           그래프의 y축 limit 값. y축 limit 안의 값만 표기되므로, limit를 넘어가는 값을 확인하고 싶을 시에는 ylim 값을 더 크게 설정하면 된다.

    Returns
    -------
    base_SF.pkl : pickle
        Base Shear Force results in pd.DataFrame type is saved as pickle in base_SF.pkl
    '''
#%% Load Data
    # Shear Force
    shear_force_data = self.shear_force_data
    # Seismic Loads List
    load_name_list = self.load_name_list
    gravity_load_name = self.gravity_load_name
    seismic_load_name_list = self.seismic_load_name_list
    DE_load_name_list = self.DE_load_name_list
    MCE_load_name_list = self.MCE_load_name_list
    
    # Base 전단력 추출
    shear_force_data = shear_force_data[shear_force_data['Name'].str.contains('base', case=False)]      
    shear_force_data.reset_index(inplace=True, drop=True)
    
#%% 데이터 Grouping
    shear_force_H1_data_grouped = pd.DataFrame()
    shear_force_H2_data_grouped = pd.DataFrame()
    
    for load_name in seismic_load_name_list:
        shear_force_H1_data_grouped['{}_H1_max'.format(load_name)] = shear_force_data[(shear_force_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (shear_force_data['Step Type'] == 'Max')]['H1(kN)'].values
            
        shear_force_H1_data_grouped['{}_H1_min'.format(load_name)] = shear_force_data[(shear_force_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (shear_force_data['Step Type'] == 'Min')]['H1(kN)'].values
    
    for load_name in seismic_load_name_list:
        shear_force_H2_data_grouped['{}_H2_max'.format(load_name)] = shear_force_data[(shear_force_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (shear_force_data['Step Type'] == 'Max')]['H2(kN)'].values
            
        shear_force_H2_data_grouped['{}_H2_min'.format(load_name)] = shear_force_data[(shear_force_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (shear_force_data['Step Type'] == 'Min')]['H2(kN)'].values   
    
    # all 절대값
    shear_force_H1_abs = shear_force_H1_data_grouped.abs()
    shear_force_H2_abs = shear_force_H2_data_grouped.abs()
    
    # Min, Max 중 최대값
    shear_force_H1_max = shear_force_H1_abs.groupby([[i//2 for i in range(0,len(seismic_load_name_list)*2)]], axis=1).max()
    shear_force_H2_max = shear_force_H2_abs.groupby([[i//2 for i in range(0,len(seismic_load_name_list)*2)]], axis=1).max()
    
    shear_force_H1_max.columns = seismic_load_name_list
    shear_force_H2_max.columns = seismic_load_name_list
    
    shear_force_H1_max.index = shear_force_data['Name'].drop_duplicates()
    shear_force_H2_max.index = shear_force_data['Name'].drop_duplicates()
    
    #%% Base Shear 그래프 그리기
    # Base Shear
    base_shear_H1 = shear_force_H1_max.copy()
    base_shear_H2 = shear_force_H2_max.copy()
    
    # 결과 dataframe -> pickle
    base_SF_result = []
    base_SF_result.append(base_shear_H1)
    base_SF_result.append(base_shear_H2)
    base_SF_result.append(DE_load_name_list)
    base_SF_result.append(MCE_load_name_list)
    with open('pkl/base_SF.pkl', 'wb') as f:
        pickle.dump(base_SF_result, f)

    count = 1

#%% Story SF

def story_SF(self, yticks=2, xlim=70000) -> pd.DataFrame:
    ''' 

    Perform-3D 해석 결과에서 각 지진파에 대한 각 층의 전단력을 그래프로 출력(kN).
    
    Parameters
    ----------
    yticks : int, optional, default=2
             그래프의 y축 눈금 간격(층간격). 층이 너무 높으면 y축에 너무 많은 층이 표기되기 때문에, 층간격을 조절해서 정돈된 그래프를 표기할 수 있다.
    
    xlim : int, optional, default=70000
           그래프의 x축 limit 값. x축 limit 안의 값만 표기되므로, limit를 넘어가는 값을 확인하고 싶을 시에는 더 큰 xlim 값을 사용하면 된다.
    
    Returns
    -------
    story_SF.pkl : pickle
        Story Shear Force results in pd.DataFrame type is saved as pickle in story_SF.pkl
    '''    
#%% Load Data
    shear_force_data = self.shear_force_data
    story_info = self.story_info

    # Seismic Loads List
    load_name_list = self.load_name_list
    gravity_load_name = self.gravity_load_name
    seismic_load_name_list = self.seismic_load_name_list
    DE_load_name_list = self.DE_load_name_list
    MCE_load_name_list = self.MCE_load_name_list

#%% Process Data   
    # 필요없는 전단력 제거(층전단력)
    shear_force_data = shear_force_data[shear_force_data['Name'].str.count('_') != 2] # underbar가 두개 들어간 행들은 제거
    # _shear 제거
    shear_force_data['Name'] = shear_force_data['Name'].str.replace('_Shear', '') 
    shear_force_data.reset_index(inplace=True, drop=True)  
    
#%% 데이터 Grouping
    shear_force_H1_data_grouped = pd.DataFrame()
    shear_force_H2_data_grouped = pd.DataFrame()
    
    for load_name in seismic_load_name_list:
        shear_force_H1_data_grouped['{}_H1_max'.format(load_name)] = shear_force_data[(shear_force_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (shear_force_data['Step Type'] == 'Max')]['H1(kN)'].values
            
        shear_force_H1_data_grouped['{}_H1_min'.format(load_name)] = shear_force_data[(shear_force_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (shear_force_data['Step Type'] == 'Min')]['H1(kN)'].values
    
    for load_name in seismic_load_name_list:
        shear_force_H2_data_grouped['{}_H2_max'.format(load_name)] = shear_force_data[(shear_force_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (shear_force_data['Step Type'] == 'Max')]['H2(kN)'].values
            
        shear_force_H2_data_grouped['{}_H2_min'.format(load_name)] = shear_force_data[(shear_force_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                      (shear_force_data['Step Type'] == 'Min')]['H2(kN)'].values   
    
    # all 절대값
    shear_force_H1_abs = shear_force_H1_data_grouped.abs()
    shear_force_H2_abs = shear_force_H2_data_grouped.abs()
    
    # Min, Max 중 최대값
    shear_force_H1_max = shear_force_H1_abs.groupby([[i//2 for i in range(0,len(seismic_load_name_list)*2)]], axis=1).max()
    shear_force_H2_max = shear_force_H2_abs.groupby([[i//2 for i in range(0,len(seismic_load_name_list)*2)]], axis=1).max()
    
    shear_force_H1_max.columns = seismic_load_name_list
    shear_force_H2_max.columns = seismic_load_name_list
    
    shear_force_H1_max.index = shear_force_data['Name'].drop_duplicates()
    shear_force_H2_max.index = shear_force_data['Name'].drop_duplicates()

    # 결과 dataframe -> pickle
    story_SF_result = []
    story_SF_result.append(shear_force_H1_max)
    story_SF_result.append(shear_force_H2_max)
    story_SF_result.append(DE_load_name_list)
    story_SF_result.append(MCE_load_name_list)
    with open('pkl/story_SF.pkl', 'wb') as f:
        pickle.dump(story_SF_result, f)

#%% IDR
def IDR(self, cri_DE=0.015, cri_MCE=0.02, yticks=2) -> pd.DataFrame:
    ''' 

    Perform-3D 해석 결과에서 각 지진파에 대한 층간변위비를 그래프로 출력.  
    
    Parameters
    ----------                  
    cri_DE : float, optional, default=0.015
             LS(인명보호)를 만족하는 층간변위비 허용기준.
             
    cri_MCE : float, optional, default=0.02
              CP(붕괴방지)를 만족하는 층간변위비 허용기준.
              
    yticks : int, optional, default=2
             그래프의 y축 눈금 간격(층간격). 층이 너무 높으면 y축에 너무 많은 층이 표기되기 때문에, 층간격을 조절해서 정돈된 그래프를 표기할 수 있다.
    
    Returns
    -------
    IDR.pkl : pickle
        Interstory Drift Ratio results in pd.DataFrame type is saved as pickle in IDR.pkl
    
    References
    -------
    [1] "철근콘크리트 건축구조물의 성능기반 내진설계 지침", 대한건축학회, p.103, 2021
    
    '''    
#%% Load Data
    drift_data = self.drift_data
    story_info = self.story_info

    # Seismic Loads List
    load_name_list = self.load_name_list
    gravity_load_name = self.gravity_load_name
    seismic_load_name_list = self.seismic_load_name_list
    DE_load_name_list = self.DE_load_name_list
    MCE_load_name_list = self.MCE_load_name_list
    
#%% Process Data

    #  Drift Name에서 story, direction 뽑아내기    
    story = [] # 1F, 2F
    direction = [] # X, Y
    position = [] # 1,5,7,11,NE,SE,SW,NW
    for i in drift_data['Drift Name']:
        i = i.strip()  # drift name 앞뒤에 있는 blank 제거
    
        if i.count('_') == 2:
            story.append(i.split('_')[0])
            direction.append(i.split('_')[-1])
            position.append(i.split('_')[1].split('_')[0])
        else:
            story.append(None)
            direction.append(None)
            
    # 지진파 순서가 섞여있을 때 sort
    drift_data = drift_data.sort_values(by=['Load Case', 'Drift ID', 'Step Type']) 
    drift_data.reset_index(inplace=True, drop=True)
    drift_data = pd.concat([pd.Series(story, name='Name'),\
                                 pd.Series(direction, name='Direction'),\
                                 pd.Series(position, name='Position'), drift_data], axis=1)
        
    # Load Case에서 지진파 이름만 뽑아서 다시 naming
    load_striped = []      
    for i in drift_data['Load Case']:
        load_striped.append(i.strip().split(' ')[-1])        
    drift_data['Load Case'] = load_striped

#%% 각 지진파에 대한 IDR 최대값
    # 각 지진파들로 변수 생성 후, 값 대입
    for load_name in seismic_load_name_list:
        globals()['IDR_x_max_{}_avg'.format(load_name)] = drift_data[(drift_data['Load Case'] == '{}'.format(load_name)) &\
                                                                      (drift_data['Direction'] == 'X') &\
                                                                      (drift_data['Step Type'] == 'Max')].groupby(['Name', 'Position'])['Drift']\
                                                                      .agg(**{'X Max avg':'max'}).groupby('Name').max() # extract 최대값
        
        globals()['IDR_x_min_{}_avg'.format(load_name)] = drift_data[(drift_data['Load Case'] == '{}'.format(load_name)) &\
                                                                      (drift_data['Direction'] == 'X') &\
                                                                      (drift_data['Step Type'] == 'Min')].groupby(['Name'])['Drift']\
                                                                      .agg(**{'X Min avg':'min'}).groupby('Name').min() # extract 최소값
            
        globals()['IDR_y_max_{}_avg'.format(load_name)] = drift_data[(drift_data['Load Case'] == '{}'.format(load_name)) &\
                                                                      (drift_data['Direction'] == 'Y') &\
                                                                      (drift_data['Step Type'] == 'Max')].groupby(['Name'])['Drift']\
                                                                      .agg(**{'Y Max avg':'max'}).groupby('Name').max() # extract 최대값
        
        globals()['IDR_y_min_{}_avg'.format(load_name)] = drift_data[(drift_data['Load Case'] == '{}'.format(load_name)) &\
                                                                      (drift_data['Direction'] == 'Y') &\
                                                                      (drift_data['Step Type'] == 'Min')].groupby(['Name'])['Drift']\
                                                                      .agg(**{'Y Min avg':'min'}).groupby('Name').min() # extract 최소값
            
        globals()['IDR_x_max_{}_avg'.format(load_name)].reset_index(inplace=True)
        globals()['IDR_x_min_{}_avg'.format(load_name)].reset_index(inplace=True)
        globals()['IDR_y_max_{}_avg'.format(load_name)].reset_index(inplace=True)
        globals()['IDR_y_min_{}_avg'.format(load_name)].reset_index(inplace=True)
   
#%% Story에 따라 IDR 정렬
    # Story 정렬하기
    story_name = story_info.loc[:, 'Story Name']
    story_name_window = drift_data['Name'].drop_duplicates()
    story_name_window_reordered = [x for x in story_name[::-1].tolist() \
                                    if x in story_name_window.tolist()]  # story name를 reference로 해서 정렬
    
    # 정렬된 Story에 따라 IDR값도 정렬
    result_each = []
    for load_name in seismic_load_name_list:   
        globals()['IDR_x_max_{}_avg'.format(load_name)]['Name'] = pd.Categorical(globals()['IDR_x_max_{}_avg'.format(load_name)]['Name'], story_name[::-1])
        globals()['IDR_x_max_{}_avg'.format(load_name)].sort_values('Name', inplace=True)
        globals()['IDR_x_max_{}_avg'.format(load_name)].reset_index(inplace=True, drop=True)
        
        globals()['IDR_x_min_{}_avg'.format(load_name)]['Name'] = pd.Categorical(globals()['IDR_x_min_{}_avg'.format(load_name)]['Name'], story_name[::-1])
        globals()['IDR_x_min_{}_avg'.format(load_name)].sort_values('Name', inplace=True)
        globals()['IDR_x_min_{}_avg'.format(load_name)].reset_index(inplace=True, drop=True)
        
        globals()['IDR_y_max_{}_avg'.format(load_name)]['Name'] = pd.Categorical(globals()['IDR_y_max_{}_avg'.format(load_name)]['Name'], story_name[::-1])
        globals()['IDR_y_max_{}_avg'.format(load_name)].sort_values('Name', inplace=True)
        globals()['IDR_y_max_{}_avg'.format(load_name)].reset_index(inplace=True, drop=True)
        
        globals()['IDR_y_min_{}_avg'.format(load_name)]['Name'] = pd.Categorical(globals()['IDR_y_min_{}_avg'.format(load_name)]['Name'], story_name[::-1])
        globals()['IDR_y_min_{}_avg'.format(load_name)].sort_values('Name', inplace=True)
        globals()['IDR_y_min_{}_avg'.format(load_name)].reset_index(inplace=True, drop=True)
        
        result_each.append(globals()['IDR_x_max_{}_avg'.format(load_name)])
        result_each.append(globals()['IDR_x_min_{}_avg'.format(load_name)])
        result_each.append(globals()['IDR_y_max_{}_avg'.format(load_name)])
        result_each.append(globals()['IDR_y_min_{}_avg'.format(load_name)])
        
#%% 모든 지진파에 대한 IDR 평균
    result_avg = []
    if len(DE_load_name_list) != 0:
            
        IDR_x_max_DE_total = pd.concat([globals()['IDR_x_max_{}_avg'.format(x)].iloc[:,-1] for x in DE_load_name_list], axis=1)
        IDR_x_min_DE_total = pd.concat([globals()['IDR_x_min_{}_avg'.format(x)].iloc[:,-1] for x in DE_load_name_list], axis=1)
        IDR_y_max_DE_total = pd.concat([globals()['IDR_y_max_{}_avg'.format(x)].iloc[:,-1] for x in DE_load_name_list], axis=1)
        IDR_y_min_DE_total = pd.concat([globals()['IDR_y_min_{}_avg'.format(x)].iloc[:,-1] for x in DE_load_name_list], axis=1)
        
        # 모든 지진파에 대한 x_max, x_min, y_max, y_min 평균
        IDR_x_max_DE_avg = IDR_x_max_DE_total.mean(axis=1)
        IDR_x_min_DE_avg = IDR_x_min_DE_total.mean(axis=1)
        IDR_y_max_DE_avg = IDR_y_max_DE_total.mean(axis=1)
        IDR_y_min_DE_avg = IDR_y_min_DE_total.mean(axis=1)
        
        # x_max, x_min, y_max, y_min 결과값을 dataframe으로 합치기
        IDR_DE_avg = pd.concat([IDR_x_max_DE_avg, IDR_x_min_DE_avg, IDR_y_max_DE_avg, IDR_y_min_DE_avg], axis=1)
        result_avg.append(IDR_DE_avg)
    
    if len(MCE_load_name_list) != 0:
        
        IDR_x_max_MCE_total = pd.concat([globals()['IDR_x_max_{}_avg'.format(x)].iloc[:,-1] for x in MCE_load_name_list], axis=1)
        IDR_x_min_MCE_total = pd.concat([globals()['IDR_x_min_{}_avg'.format(x)].iloc[:,-1] for x in MCE_load_name_list], axis=1)
        IDR_y_max_MCE_total = pd.concat([globals()['IDR_y_max_{}_avg'.format(x)].iloc[:,-1] for x in MCE_load_name_list], axis=1)
        IDR_y_min_MCE_total = pd.concat([globals()['IDR_y_min_{}_avg'.format(x)].iloc[:,-1] for x in MCE_load_name_list], axis=1)
        
        IDR_x_max_MCE_avg = IDR_x_max_MCE_total.mean(axis=1)
        IDR_x_min_MCE_avg = IDR_x_min_MCE_total.mean(axis=1)
        IDR_y_max_MCE_avg = IDR_y_max_MCE_total.mean(axis=1)
        IDR_y_min_MCE_avg = IDR_y_min_MCE_total.mean(axis=1)
        
        IDR_MCE_avg = pd.concat([IDR_x_max_MCE_avg, IDR_x_min_MCE_avg, IDR_y_max_MCE_avg, IDR_y_min_MCE_avg], axis=1)
        result_avg.append(IDR_MCE_avg)

    # 결과 dataframe -> pickle
    IDR_result = []
    IDR_result.append(result_each)
    IDR_result.append(result_avg)
    IDR_result.append(DE_load_name_list)
    IDR_result.append(MCE_load_name_list)
    IDR_result.append(story_name_window_reordered)
    with open('pkl/IDR.pkl', 'wb') as f:
        pickle.dump(IDR_result, f)
    
#%% Pushover

def Pushover(result_xlsx_path, x_result_txt, y_result_txt, base_SF_design=None, pp_x=None, pp_y=None):
    ''' 

    Perform-3D의 Pushover 해석 결과로 성능곡선 그래프 출력 (아직 코드로만 실행 가능)
    
    Parameters
    ----------
    result_xlsx_path : str
                  Perform-3D에서 나온 해석 파일의 경로.
                  
    result_xlsx : str, optional, default='Analysis Result'
                  Perform-3D에서 나온 해석 파일의 이름. 해당 파일 이름이 포함된 파일들을 모두 불러온다.
                  
    ylim : int, optional, default=70000
           그래프의 y축 limit 값. y축 limit 안의 값만 표기되므로, limit를 넘어가는 값을 확인하고 싶을 시에는 ylim 값을 더 크게 설정하면 된다.
    
    pp_x : float
    
    Returns
    -------
    '''

    # 설계밑면전단력 값 입력
    x_result_txt = 'N1_PO_X.txt'
    y_result_txt = 'N1_PO_Y.txt'
    design_base_shear_x = 10037*0.85 # kN (GEN값 * 0.85)
    design_base_shear_y = 10369*0.85 # kN
    pp_DE_x = [1.915e-4, 32480]
    pp_MCE_x = [1.856e-4, 32570]
    pp_DE_y = [0.001898, 10470]
    pp_MCE_y = [0.002489, 11820]
###############################################################################     

    # data_X = pd.read_csv(result_xlsx_path[0]+'\\'+x_result_txt, skiprows=8, header=None)
    # data_Y = pd.read_csv(result_xlsx_path[0]+'\\'+y_result_txt, skiprows=8, header=None)
    data_X = pd.read_csv(r'D:\이형우\성능기반 내진설계\21-GR-222 광명 4R구역 주택재개발사업 성능기반내진설계\해석 결과\101_N1'+'\\'+x_result_txt, skiprows=8, header=None)
    data_Y = pd.read_csv(r'D:\이형우\성능기반 내진설계\21-GR-222 광명 4R구역 주택재개발사업 성능기반내진설계\해석 결과\101_N1'+'\\'+y_result_txt, skiprows=8, header=None)
    data_X.columns = ['Drift', 'Base Shear']
    data_Y.columns = ['Drift', 'Base Shear']
    
    ### 성능곡선 그리기
    ### X Direction
    fig1 = plt.figure(1, figsize=(8,5))  # 그래프 사이즈
    plt.grid()
    plt.plot(data_X['Drift'], data_X['Base Shear'], color = 'k', linewidth = 1)
    plt.title('Capacity Curve (X-dir)', pad = 10)
    plt.xlabel('Reference Drift', labelpad= 10) # fontweight='bold'
    plt.ylabel('Base Shear(kN)', labelpad= 10)
    plt.xlim([0, max(data_X['Drift'])])
    plt.ylim([0, max(data_X['Base Shear'])+3000])
    
    # 설계 밑면전단력 그리기
    if design_base_shear_x != None:
        plt.axhline(design_base_shear_x, 0, 1, color = 'royalblue', linestyle='--', linewidth = 1.5)
    
    # 성능점 그리기
    plt.plot(pp_DE_x[0], pp_DE_x[1], color='r', marker='o')
    plt.text(pp_DE_x[0]*1.3, pp_DE_x[1], 'Performance Point at DE \n ({},{})'.format(pp_DE_x[0], pp_DE_x[1])
             , verticalalignment='top')
    
    plt.plot(pp_MCE_x[0], pp_MCE_x[1], color='g', marker='o')
    plt.text(pp_MCE_x[0]*1.3, pp_MCE_x[1], 'Performance Point at MCE \n ({},{})'.format(pp_MCE_x[0], pp_MCE_x[1])
             , verticalalignment='bottom')
    
    plt.show()
    # yield fig1
    
    
    ### Y Direction
    fig2 = plt.figure(2, figsize=(8,5))  # 그래프 사이즈
    plt.grid()
    plt.plot(data_Y['Drift'], data_Y['Base Shear'], color = 'k', linewidth = 1)
    plt.title('Capacity Curve (Y-dir)', pad = 10)
    plt.xlabel('Reference Drift', labelpad= 10) # fontweight='bold'
    plt.ylabel('Base Shear(kN)', labelpad= 10)
    plt.xlim([0, max(data_Y['Drift'])])
    plt.ylim([0, max(data_Y['Base Shear'])+3000])
    
    if design_base_shear_y != None:
        plt.axhline(design_base_shear_y, 0, 1, color='royalblue', linestyle='--', linewidth=1.5)

    plt.plot(pp_DE_y[0], pp_DE_y[1], color='r', marker='o')
    plt.text(pp_DE_y[0]*1.3, pp_DE_y[1], 'Performance Point at DE \n ({},{})'.format(pp_DE_y[0], pp_DE_y[1])
             , verticalalignment='top')
    
    plt.plot(pp_MCE_y[0], pp_MCE_y[1], color='g', marker='o')
    plt.text(pp_MCE_y[0]*1.3, pp_MCE_y[1], 'Performance Point at MCE \n ({},{})'.format(pp_MCE_y[0], pp_MCE_y[1])
             , verticalalignment='bottom')
    
    plt.show()
    # yield fig2
    
    print(max(data_X['Base Shear']), design_base_shear_x, max(data_X['Base Shear'])/design_base_shear_x)
    print(max(data_Y['Base Shear']), design_base_shear_y, max(data_Y['Base Shear'])/design_base_shear_y)