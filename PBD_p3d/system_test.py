import pandas as pd
import os
from collections import deque  # Double-ended Queue : 자료의 앞, 뒤 양 방향에서 자료를 추가하거나 제거가능
import matplotlib.pyplot as plt
from decimal import Decimal, ROUND_UP

#%% Base SF

def base_SF_test(result_path, result_xlsx='Analysis Result', ylim=70000):
    ''' 

    Perform-3D 해석 결과에서 각 지진파에 대한 Base층의 전단력을 막대그래프 형식으로 출력. (kN)
    
    Parameters
    ----------
    result_path : str
                  Perform-3D에서 나온 해석 파일의 경로.
                  
    result_xlsx : str, optional, default='Analysis Result'
                  Perform-3D에서 나온 해석 파일의 이름. 해당 파일 이름이 포함된 파일들을 모두 불러온다.
                  
    ylim : int, optional, default=70000
           그래프의 y축 limit 값. y축 limit 안의 값만 표기되므로, limit를 넘어가는 값을 확인하고 싶을 시에는 ylim 값을 더 크게 설정하면 된다.
    
    Returns
    -------
    '''
#%% Analysis Result 불러오기

    if isinstance(result_path, list):
        to_load_list = result_path
        
    else:
        to_load_list = []
        file_names = os.listdir(result_path)
        for file_name in file_names:
            if (result_xlsx in file_name) and ('~$' not in file_name):
                to_load_list.append(file_name)
    
    # 전단력 불러오기
    shear_force_data = pd.DataFrame()
    
    for i in to_load_list:
        result_data_raw = pd.ExcelFile(result_path + '\\' + i)
        result_data_sheets = pd.read_excel(result_data_raw, ['Structure Section Forces'], skiprows=[0,2])
        
        column_name_to_slice = ['StrucSec Name', 'Load Case', 'Step Type', 'FH1', 'FH2']
        shear_force_data_temp = result_data_sheets['Structure Section Forces'].loc[:,column_name_to_slice]
        shear_force_data = pd.concat([shear_force_data, shear_force_data_temp])
        
    shear_force_data.columns = ['Name', 'Load Case', 'Step Type', 'H1(kN)', 'H2(kN)']
    
    # Base 전단력 추출
    shear_force_data = shear_force_data[shear_force_data['Name'].str.contains('base', case=False)]
      
    shear_force_data.reset_index(inplace=True, drop=True)
    
    # _shear 제거
    shear_force_data['Name'] = shear_force_data['Name'].str.rstrip('_Shear')
    
#%% 지진파 이름 list 만들기
    load_name_list = []
    for i in shear_force_data['Load Case'].drop_duplicates():
        new_i = i.split('+')[1]
        new_i = new_i.strip()
        load_name_list.append(new_i)
    
    gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
    seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]
    
    seismic_load_name_list.sort()
    
    DE_load_name_list = [x for x in load_name_list if 'DE' in x] # base shear로 사용할 지진파 개수 산정을 위함
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]
    
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
    
    count = 1
    
# DE Plot
  
    if len(MCE_load_name_list) != 0:
    
# H1_DE
        fig1, ax1 = plt.subplots(1,1)
        fig1.set_dpi(150)
        ax1.set_ylim(0, ylim)
        
        ax1.bar(range(len(MCE_load_name_list)), base_shear_H1.iloc[0, 0:len(MCE_load_name_list)]\
                , color='darkblue', edgecolor='k', label = 'Max. Base Shear')
        ax1.axhline(y= base_shear_H1.iloc[0, 0:len(MCE_load_name_list)].mean(), color='r', linestyle='-', label='Average')
        ax1.set_xticks(range(14), range(1,15))
        # ax1.xticks(range(14), load_name[0:14], fontsize=8.5)
        
        ax1.set_xlabel('Ground Motion No.')
        ax1.set_ylabel('Base Shear(kN)')
        ax1.legend(loc = 2)
        ax1.set_title('H1 DE')
        
        # plt.savefig(memfile)
        # plt.close()
        # count += 1
        
        yield ax1
        print('base_shear_avg(H1_DE) =', Decimal(str(base_shear_H1.iloc[0, 0:len(DE_load_name_list)].mean()))\
              .quantize(Decimal('.01'), rounding=ROUND_UP))
        
        # H2_DE
        fig2 = plt.figure(count, dpi=150)
        plt.ylim(0, ylim)
        
        plt.bar(range(len(DE_load_name_list)), base_shear_H2.iloc[0, 0:len(DE_load_name_list)], color='darkblue', edgecolor='k', label = 'Max. Base Shear')
        plt.axhline(y= base_shear_H2.iloc[0, 0:len(DE_load_name_list)].mean(), color='r', linestyle='-', label='Average')
        plt.xticks(range(14), range(1,15))
        # plt.xticks(range(14), load_name[0:14], fontsize=8.5)
        
        plt.xlabel('Ground Motion No.')
        plt.ylabel('Base Shear(kN)')
        plt.legend(loc = 2)
        plt.title('H2 DE')
        
        # plt.savefig(memfile2)
        plt.close()
        count += 1

        yield fig2
        print('base_shear_avg(H2_DE) =', Decimal(str(base_shear_H2.iloc[0, 0:len(DE_load_name_list)].mean()))\
              .quantize(Decimal('.01'), rounding=ROUND_UP))

# MCE Plot
  
    if len(DE_load_name_list) != 0:
    
        # H1_MCE
        fig3 = plt.figure(count, dpi=150)
        plt.ylim(0, ylim)
        
        plt.bar(range(len(MCE_load_name_list)), base_shear_H1\
                .iloc[0, len(DE_load_name_list):len(DE_load_name_list)+len(MCE_load_name_list)]\
                , color='darkblue', edgecolor='k', label = 'Max. Base Shear')
        plt.axhline(y= base_shear_H1.iloc[0, len(DE_load_name_list):len(DE_load_name_list)+len(MCE_load_name_list)]\
                    .mean(), color='r', linestyle='-', label='Average')
        plt.xticks(range(14), range(1,15))
        # plt.xticks(range(14), load_name[0:14], fontsize=8.5)
        
        plt.xlabel('Ground Motion No.')
        plt.ylabel('Base Shear(kN)')
        plt.legend(loc = 2)
        plt.title('H1 MCE')
        
        # plt.savefig(memfile3)
        plt.close()
        count += 1

        yield fig3
        print('base_shear_avg(H1_MCE) =', Decimal(str(base_shear_H1.iloc[0, len(DE_load_name_list):len(DE_load_name_list)+len(MCE_load_name_list)].mean()))\
              .quantize(Decimal('.01'), rounding=ROUND_UP))

        # H2_MCE
        fig4 = plt.figure(count, dpi=150)
        plt.ylim(0, ylim)
        
        plt.bar(range(len(MCE_load_name_list)), base_shear_H2\
                .iloc[0, len(DE_load_name_list):len(DE_load_name_list)+len(MCE_load_name_list)]\
                , color='darkblue', edgecolor='k', label = 'Max. Base Shear')
        plt.axhline(y= base_shear_H2.iloc[0, len(DE_load_name_list):len(DE_load_name_list)+len(MCE_load_name_list)]\
                    .mean(), color='r', linestyle='-', label='Average')
        plt.xticks(range(14), range(1,15))
        # plt.xticks(range(14), load_name[0:14], fontsize=8.5)
        
        plt.xlabel('Ground Motion No.')
        plt.ylabel('Base Shear(kN)')
        plt.legend(loc = 2)
        plt.title('H2 MCE')
        
        # plt.savefig(memfile4)
        plt.close()
        count += 1

        yield fig4
        print('base_shear_avg(H2_MCE) =', Decimal(str(base_shear_H2.iloc[0, len(DE_load_name_list):len(DE_load_name_list)+len(MCE_load_name_list)].mean()))\
              .quantize(Decimal('.01'), rounding=ROUND_UP))

#%% Story SF

def story_SF(input_path, input_xlsx, result_path, result_xlsx='Analysis Result', yticks=2, xlim=70000):
    ''' 

    Perform-3D 해석 결과에서 각 지진파에 대한 각 층의 전단력을 그래프로 출력(kN).
    
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
                 
    yticks : int, optional, default=2
             그래프의 y축 눈금 간격(층간격). 층이 너무 높으면 y축에 너무 많은 층이 표기되기 때문에, 층간격을 조절해서 정돈된 그래프를 표기할 수 있다.
    
    xlim : int, optional, default=70000
           그래프의 x축 limit 값. x축 limit 안의 값만 표기되므로, limit를 넘어가는 값을 확인하고 싶을 시에는 더 큰 xlim 값을 사용하면 된다.
    
    Returns
    -------
    '''    
#%% Analysis Result 불러오기

    to_load_list = []
    file_names = os.listdir(result_path)
    for file_name in file_names:
        if (result_xlsx in file_name) and ('~$' not in file_name):
            to_load_list.append(file_name)
    
    # 전단력 불러오기
    shear_force_data = pd.DataFrame()
    
    for i in to_load_list:
        shear_force_data_temp = pd.read_excel(result_path + '\\' + i,
                                   sheet_name='Structure Section Forces', skiprows=2, usecols=[0,3,5,6,7])
        shear_force_data = pd.concat([shear_force_data, shear_force_data_temp])
        
    shear_force_data.columns = ['Name', 'Load Case', 'Step Type', 'H1(kN)', 'H2(kN)']
    
    # 필요없는 전단력 제거(층전단력)
    shear_force_data = shear_force_data[shear_force_data['Name'].str.count('_') != 2] # underbar가 두개 들어간 행들은 제거
      
    shear_force_data.reset_index(inplace=True, drop=True)
    
    # _shear 제거
    shear_force_data['Name'] = shear_force_data['Name'].str.rstrip('_Shear')
    
#%% 부재명, H1, H2 값 뽑기
    
    # 지진파 이름 list 만들기
    load_name_list = []
    for i in shear_force_data['Load Case'].drop_duplicates():
        new_i = i.split('+')[1]
        new_i = new_i.strip()
        load_name_list.append(new_i)
    
    gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
    seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]
    
    seismic_load_name_list.sort()
    
    DE_load_name_list = [x for x in load_name_list if 'DE' in x] # base shear로 사용할 지진파 개수 산정을 위함
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]
    
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
    
#%% Story 정보 load
    
    # Story 정보에서 층이름만 뽑아내기
    story_info_xlsx_sheet = 'Story Data'
    story_info = pd.read_excel(input_path + '\\' + input_xlsx, sheet_name=story_info_xlsx_sheet, skiprows=3, usecols=[0, 1, 2], keep_default_na=False)
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    story_name = story_info.loc[:, 'Story Name']

#%% Story Shear 그래프 그리기

    count = 1
    
    # DE Plot    
    if len(DE_load_name_list) != 0:

        ### H1_DE
        fig1 = plt.figure(count, dpi=150)
        plt.xlim(0, xlim)
        
        # 지진파별 plot
        for i in range(len(DE_load_name_list)):
            plt.plot(shear_force_H1_max.iloc[:,i], range(shear_force_H1_max.shape[0]), label=DE_load_name_list[i], linewidth=0.7)
            
        # 평균 plot
        plt.plot(shear_force_H1_max.iloc[:,0:len(DE_load_name_list)]\
                 .mean(axis=1), range(shear_force_H1_max.shape[0]), color='k', label='Average', linewidth=2)
        
        plt.yticks(range(shear_force_H1_max.shape[0])[::yticks], shear_force_H1_max.index[::yticks], fontsize=8.5)
        # plt.xticks(range(14), range(1,15))
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('Story Shear(kN)')
        plt.ylabel('Story')
        plt.legend(loc=1, fontsize=8)
        plt.title('H1 DE')
        
        plt.tight_layout()
        # plt.savefig(memfile5)
        plt.close()
        count += 1

        yield fig1
        
        # H2_DE
        fig2 = plt.figure(count, dpi=150)
        plt.xlim(0, xlim)
        
        # 지진파별 plot
        for i in range(len(DE_load_name_list)):
            plt.plot(shear_force_H2_max.iloc[:,i], range(shear_force_H2_max.shape[0]), label=DE_load_name_list[i], linewidth=0.7)
            
        # 평균 plot
        plt.plot(shear_force_H2_max.iloc[:,0:len(DE_load_name_list)]\
                 .mean(axis=1), range(shear_force_H2_max.shape[0]), color='k', label='Average', linewidth=2)
        
        plt.yticks(range(shear_force_H2_max.shape[0])[::yticks], shear_force_H2_max.index[::yticks], fontsize=8.5)
        # plt.xticks(range(14), range(1,15))
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('Story Shear(kN)')
        plt.ylabel('Story')
        plt.legend(loc=1, fontsize=8)
        plt.title('H2 DE')
        
        plt.tight_layout()
        # plt.savefig(memfile6)
        plt.close()
        count += 1

        yield fig2
    
    # DE Plot    
    if len(MCE_load_name_list) != 0:    
    
        ### H1_MCE
        fig3 = plt.figure(count, dpi=150)
        plt.xlim(0, xlim)
        
        # 지진파별 plot
        for i in range(len(MCE_load_name_list)):
            plt.plot(shear_force_H1_max.iloc[:,i+len(DE_load_name_list)], range(shear_force_H1_max.shape[0]), label=MCE_load_name_list[i], linewidth=0.7)
            
        # 평균 plot
        plt.plot(shear_force_H1_max.iloc[:,len(DE_load_name_list)\
                                         :len(DE_load_name_list)+len(MCE_load_name_list)]\
                 .mean(axis=1), range(shear_force_H1_max.shape[0]), color='k', label='Average', linewidth=2)
        
        plt.yticks(range(shear_force_H1_max.shape[0])[::yticks], shear_force_H1_max.index[::yticks], fontsize=8.5)
        # plt.xticks(range(14), range(1,15))
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('Story Shear(kN)')
        plt.ylabel('Story')
        plt.legend(loc=1, fontsize=8)
        plt.title('H1 MCE')
        
        plt.tight_layout()
        # plt.savefig(memfile7)
        plt.close()
        count += 1

        yield fig3
        
        # H2_MCE
        fig4 = plt.figure(count, dpi=150)
        plt.xlim(0, xlim)
        
        # 지진파별 plot
        for i in range(len(MCE_load_name_list)):
            plt.plot(shear_force_H2_max.iloc[:,i+len(DE_load_name_list)], range(shear_force_H2_max.shape[0]), label=MCE_load_name_list[i], linewidth=0.7)
            
        # 평균 plot
        plt.plot(shear_force_H2_max.iloc[:,len(DE_load_name_list)\
                                         :len(DE_load_name_list)+len(MCE_load_name_list)]\
                 .mean(axis=1), range(shear_force_H2_max.shape[0]), color='k', label='Average', linewidth=2)
        
        plt.yticks(range(shear_force_H2_max.shape[0])[::yticks], shear_force_H2_max.index[::yticks], fontsize=8.5)
        # plt.xticks(range(14), range(1,15))
        
        # 기타
        plt.grid(linestyle='-.')
        plt.xlabel('Story Shear(kN)')
        plt.ylabel('Story')
        plt.legend(loc=1, fontsize=8)
        plt.title('H2 MCE')
        
        plt.tight_layout()
        # plt.savefig(memfile8)
        plt.close()
        count += 1

        yield fig4

#%% IDR
def IDR(input_path, input_xlsx, result_path, result_xlsx='Analysis Result', cri_DE=0.015, cri_MCE=0.02, yticks=2):   
    ''' 

    Perform-3D 해석 결과에서 각 지진파에 대한 층간변위비를 그래프로 출력.  
    
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
                  
    cri_DE : float, optional, default=0.015
             LS(인명보호)를 만족하는 층간변위비 허용기준.
             
    cri_MCE : float, optional, default=0.02
              CP(붕괴방지)를 만족하는 층간변위비 허용기준.
              
    yticks : int, optional, default=2
             그래프의 y축 눈금 간격(층간격). 층이 너무 높으면 y축에 너무 많은 층이 표기되기 때문에, 층간격을 조절해서 정돈된 그래프를 표기할 수 있다.
    
    Yields
    -------
    fig1 : matplotlib.pyplot.figure or None
           DE(설계지진) 발생 시 x방향 층간변위비 그래프
    
    fig2 : matplotlib.pyplot.figure or None
           DE(설계지진) 발생 시 y방향 층간변위비 그래프
    
    fig3 : matplotlib.pyplot.figure or None
           MCE(최대고려지진) 발생 시 x방향 층간변위비 그래프
    
    fig4 : matplotlib.pyplot.figure or None
           MCE(최대고려지진) 발생 시 y방향 층간변위비 그래프
    
    Raises
    -------
    
    References
    -------
    [1] "철근콘크리트 건축구조물의 성능기반 내진설계 지침", 대한건축학회, p.103, 2021
    
    '''    
#%% Analysis Result 불러오기
    to_load_list = []
    file_names = os.listdir(result_path)
    for file_name in file_names:
        if (result_xlsx in file_name) and ('~$' not in file_name):
            to_load_list.append(file_name)
    
    # Gage data
    IDR_result_data = pd.DataFrame()
    for i in to_load_list:
        IDR_result_data_temp = pd.read_excel(result_path + '\\' + i,
                                    sheet_name='Drift Output', skiprows=[0, 2], header=0, usecols=[0, 1, 3, 5, 6]) # usecols로 원하는 열만 불러오기
        IDR_result_data = pd.concat([IDR_result_data, IDR_result_data_temp])        
    
    IDR_result_data = IDR_result_data.sort_values(by=['Load Case', 'Drift ID', 'Step Type']) # 지진파 순서가 섞여있을 때 sort
    
    # Story Info data
    story_info_xlsx_sheet = 'Story Data'
    story_info = pd.read_excel(input_path + '\\' + input_xlsx, sheet_name=story_info_xlsx_sheet, skiprows=3, usecols=[0, 1, 2], keep_default_na=False)
    story_info.columns = ['Index', 'Story Name', 'Height(mm)']
    story_name = story_info.loc[:, 'Story Name']
    
#%% Drift Name에서 story, direction 뽑아내기
    drift_name = IDR_result_data['Drift Name']
    
    story = []
    direction = []
    position = []
    for i in drift_name:
        i = i.strip()  # drift_name 앞뒤에 있는 blank 제거
    
        if i.count('_') == 2:
            story.append(i.split('_')[0])
            direction.append(i.split('_')[-1])
            position.append(i.split('_')[1].split('_')[0])
        else:
            story.append(None)
            direction.append(None)
    
    # Load Case에서 지진파 이름만 뽑아서 다시 naming
    load_striped = []        
    for i in IDR_result_data['Load Case']:
        load_striped.append(i.strip().split(' ')[-1])
        
    IDR_result_data['Load Case'] = load_striped
        
    
    IDR_result_data.reset_index(inplace=True, drop=True)
    IDR_result_data = pd.concat([pd.Series(story, name='Name'),\
                                 pd.Series(direction, name='Direction'),\
                                 pd.Series(position, name='Position'), IDR_result_data], axis=1)
        
#%% 지진파 이름 자동 생성

    load_name_list = IDR_result_data['Load Case'].drop_duplicates()
    seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]
    
    seismic_load_name_list.sort()
    
    DE_load_name_list = [x for x in load_name_list if 'DE' in x] # base shear로 사용할 지진파 개수 산정을 위함
    MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]

#%% IDR값(방향에 따른)
    ### 지진파별 평균
    
    # 각 지진파들로 변수 생성 후, 값 대입
    for load_name in seismic_load_name_list:
        globals()['IDR_x_max_{}_avg'.format(load_name)] = IDR_result_data[(IDR_result_data['Load Case'] == '{}'.format(load_name)) &\
                                                                      (IDR_result_data['Direction'] == 'X') &\
                                                                      (IDR_result_data['Step Type'] == 'Max')].groupby(['Name', 'Position'])['Drift']\
                                                                      .agg(**{'X Max avg':'mean'}).groupby('Name').max()
        
        globals()['IDR_x_min_{}_avg'.format(load_name)] = IDR_result_data[(IDR_result_data['Load Case'] == '{}'.format(load_name)) &\
                                                                      (IDR_result_data['Direction'] == 'X') &\
                                                                      (IDR_result_data['Step Type'] == 'Min')].groupby(['Name'])['Drift']\
                                                                      .agg(**{'X Min avg':'mean'}).groupby('Name').min()
            
        globals()['IDR_y_max_{}_avg'.format(load_name)] = IDR_result_data[(IDR_result_data['Load Case'] == '{}'.format(load_name)) &\
                                                                      (IDR_result_data['Direction'] == 'Y') &\
                                                                      (IDR_result_data['Step Type'] == 'Max')].groupby(['Name'])['Drift']\
                                                                      .agg(**{'X Max avg':'mean'}).groupby('Name').max()
        
        globals()['IDR_y_min_{}_avg'.format(load_name)] = IDR_result_data[(IDR_result_data['Load Case'] == '{}'.format(load_name)) &\
                                                                      (IDR_result_data['Direction'] == 'Y') &\
                                                                      (IDR_result_data['Step Type'] == 'Min')].groupby(['Name'])['Drift']\
                                                                      .agg(**{'X Min avg':'mean'}).groupby('Name').min()
            
        globals()['IDR_x_max_{}_avg'.format(load_name)].reset_index(inplace=True)
        globals()['IDR_x_min_{}_avg'.format(load_name)].reset_index(inplace=True)
        globals()['IDR_y_max_{}_avg'.format(load_name)].reset_index(inplace=True)
        globals()['IDR_y_min_{}_avg'.format(load_name)].reset_index(inplace=True)
        
    # Story 정렬하기
    story_name_window = IDR_result_data['Name'].drop_duplicates()
    story_name_window_reordered = [x for x in story_name[::-1].tolist() \
                                    if x in story_name_window.tolist()]  # story name를 reference로 해서 정렬
    
    # 정렬된 Story에 따라 IDR값도 정렬
    for load_name in seismic_load_name_list:   
        globals()['IDR_x_max_{}_avg'.format(load_name)]['Name'] = pd.Categorical(globals()['IDR_x_max_{}_avg'.format(load_name)]['Name'], story_name[::-1])
        globals()['IDR_x_max_{}_avg'.format(load_name)].sort_values('Name', inplace=True)
        
        globals()['IDR_x_min_{}_avg'.format(load_name)]['Name'] = pd.Categorical(globals()['IDR_x_min_{}_avg'.format(load_name)]['Name'], story_name[::-1])
        globals()['IDR_x_min_{}_avg'.format(load_name)].sort_values('Name', inplace=True)
        
        globals()['IDR_y_max_{}_avg'.format(load_name)]['Name'] = pd.Categorical(globals()['IDR_y_max_{}_avg'.format(load_name)]['Name'], story_name[::-1])
        globals()['IDR_y_max_{}_avg'.format(load_name)].sort_values('Name', inplace=True)
        
        globals()['IDR_y_min_{}_avg'.format(load_name)]['Name'] = pd.Categorical(globals()['IDR_y_min_{}_avg'.format(load_name)]['Name'], story_name[::-1])
        globals()['IDR_y_min_{}_avg'.format(load_name)].sort_values('Name', inplace=True)
        
#%% IDR값(방향에 따른) 전체 평균
    
    if len(DE_load_name_list) != 0:
    
        IDR_x_max_DE_avg = IDR_result_data[(IDR_result_data['Load Case'].str.contains('DE')) &\
                                                (IDR_result_data['Direction'] == 'X') &\
                                                (IDR_result_data['Step Type'] == 'Max')].groupby(['Name', 'Position'])['Drift'].agg(**{'X Max avg':'mean'}).groupby('Name').max()
        IDR_x_min_DE_avg = IDR_result_data[(IDR_result_data['Load Case'].str.contains('DE')) &\
                                                (IDR_result_data['Direction'] == 'X') &\
                                                (IDR_result_data['Step Type'] == 'Min')].groupby(['Name', 'Position'])['Drift'].agg(**{'X Max avg':'mean'}).groupby('Name').min()
        IDR_y_max_DE_avg = IDR_result_data[(IDR_result_data['Load Case'].str.contains('DE')) &\
                                                (IDR_result_data['Direction'] == 'Y') &\
                                                (IDR_result_data['Step Type'] == 'Max')].groupby(['Name', 'Position'])['Drift'].agg(**{'X Max avg':'mean'}).groupby('Name').max()
        IDR_y_min_DE_avg = IDR_result_data[(IDR_result_data['Load Case'].str.contains('DE')) &\
                                                (IDR_result_data['Direction'] == 'Y') &\
                                                (IDR_result_data['Step Type'] == 'Min')].groupby(['Name', 'Position'])['Drift'].agg(**{'X Max avg':'mean'}).groupby('Name').min()
 
        IDR_x_max_DE_avg.reset_index(inplace=True)
        IDR_x_min_DE_avg.reset_index(inplace=True)
        IDR_y_max_DE_avg.reset_index(inplace=True)
        IDR_y_min_DE_avg.reset_index(inplace=True)
    
    if len(MCE_load_name_list) != 0:
        
        IDR_x_max_MCE_avg = IDR_result_data[(IDR_result_data['Load Case'].str.contains('MCE')) &\
                                                (IDR_result_data['Direction'] == 'X') &\
                                                (IDR_result_data['Step Type'] == 'Max')].groupby(['Name', 'Position'])['Drift'].agg(**{'X Max avg':'mean'}).groupby('Name').max()
        IDR_x_min_MCE_avg = IDR_result_data[(IDR_result_data['Load Case'].str.contains('MCE')) &\
                                                (IDR_result_data['Direction'] == 'X') &\
                                                (IDR_result_data['Step Type'] == 'Min')].groupby(['Name', 'Position'])['Drift'].agg(**{'X Max avg':'mean'}).groupby('Name').min()
        IDR_y_max_MCE_avg = IDR_result_data[(IDR_result_data['Load Case'].str.contains('MCE')) &\
                                                (IDR_result_data['Direction'] == 'Y') &\
                                                (IDR_result_data['Step Type'] == 'Max')].groupby(['Name', 'Position'])['Drift'].agg(**{'X Max avg':'mean'}).groupby('Name').max()
        IDR_y_min_MCE_avg = IDR_result_data[(IDR_result_data['Load Case'].str.contains('MCE')) &\
                                                (IDR_result_data['Direction'] == 'Y') &\
                                                (IDR_result_data['Step Type'] == 'Min')].groupby(['Name', 'Position'])['Drift'].agg(**{'X Max avg':'mean'}).groupby('Name').min()
    
        IDR_x_max_MCE_avg.reset_index(inplace=True)
        IDR_x_min_MCE_avg.reset_index(inplace=True)
        IDR_y_max_MCE_avg.reset_index(inplace=True)
        IDR_y_min_MCE_avg.reset_index(inplace=True)
    
    # 정렬된 Story에 따라 IDR값도 정렬
    if len(DE_load_name_list) != 0:
    
        IDR_x_max_DE_avg['Name'] = pd.Categorical(IDR_x_max_DE_avg['Name'], story_name[::-1])
        IDR_x_min_DE_avg['Name'] = pd.Categorical(IDR_x_min_DE_avg['Name'], story_name[::-1])
        IDR_y_max_DE_avg['Name'] = pd.Categorical(IDR_y_max_DE_avg['Name'], story_name[::-1])
        IDR_y_min_DE_avg['Name'] = pd.Categorical(IDR_y_min_DE_avg['Name'], story_name[::-1])
    
        IDR_x_max_DE_avg.sort_values('Name', inplace=True)
        IDR_x_min_DE_avg.sort_values('Name', inplace=True)
        IDR_y_max_DE_avg.sort_values('Name', inplace=True)
        IDR_y_min_DE_avg.sort_values('Name', inplace=True)
    
    if len(MCE_load_name_list) != 0:
        IDR_x_max_MCE_avg['Name'] = pd.Categorical(IDR_x_max_MCE_avg['Name'], story_name[::-1])
        IDR_x_min_MCE_avg['Name'] = pd.Categorical(IDR_x_min_MCE_avg['Name'], story_name[::-1])
        IDR_y_max_MCE_avg['Name'] = pd.Categorical(IDR_y_max_MCE_avg['Name'], story_name[::-1])
        IDR_y_min_MCE_avg['Name'] = pd.Categorical(IDR_y_min_MCE_avg['Name'], story_name[::-1])
    
        IDR_x_max_MCE_avg.sort_values('Name', inplace=True)
        IDR_x_min_MCE_avg.sort_values('Name', inplace=True)
        IDR_y_max_MCE_avg.sort_values('Name', inplace=True)
        IDR_y_min_MCE_avg.sort_values('Name', inplace=True)
    
#%% 그래프 (방향에 따른)

    count = 1

    # DE Plot
    if len(DE_load_name_list) != 0:

        ### H1 DE 그래프
        fig1 = plt.figure(count, figsize=(5, 7), dpi=150)
        plt.xlim(-0.025, 0.025)
        
        # 지진파별 plot
        for load_name in DE_load_name_list:
            plt.plot(globals()['IDR_x_max_{}_avg'.format(load_name)].iloc[:,-1], IDR_x_max_DE_avg.iloc[:,0], label='{}'.format(load_name), linewidth=0.7)
            plt.plot(globals()['IDR_x_min_{}_avg'.format(load_name)].iloc[:,-1], IDR_x_max_DE_avg.iloc[:,0], linewidth=0.7)
                
        # 평균 plot      
        plt.plot(IDR_x_max_DE_avg.iloc[:,-1], IDR_x_max_DE_avg.iloc[:,0], color='k', label='Average', linewidth=2)
        plt.plot(IDR_x_min_DE_avg.iloc[:,-1], IDR_x_min_DE_avg.iloc[:,0], color='k', linewidth=2)
        
        # reference line 그려서 허용치 나타내기
        plt.axvline(x=-cri_DE, color='r', linestyle='--', label='LS')
        plt.axvline(x=cri_DE, color='r', linestyle='--')
        
        # 기타
        plt.yticks(story_name_window_reordered[::yticks], story_name_window_reordered[::yticks])
        plt.grid(linestyle='-.')
        plt.xlabel('Interstory Drift Ratios(m/m)')
        plt.ylabel('Story')
        plt.legend(loc=4, fontsize=8)
        plt.title('H1 DE')
        
        plt.tight_layout()
        # plt.savefig(result_path + '\\' + 'IDR_H1_DE')
        plt.close()
        count += 1
        
        yield fig1
        
        ### H2 DE 그래프
        fig2 = plt.figure(count, figsize=(5, 7), dpi=150)
        plt.xlim(-0.025, 0.025)
        
        # 지진파별 plot
        for load_name in DE_load_name_list:
            plt.plot(globals()['IDR_y_max_{}_avg'.format(load_name)].iloc[:,-1], IDR_y_max_DE_avg.iloc[:,0], label='{}'.format(load_name), linewidth=0.7)
            plt.plot(globals()['IDR_y_min_{}_avg'.format(load_name)].iloc[:,-1], IDR_y_max_DE_avg.iloc[:,0], linewidth=0.7)
        
        # 평균 plot
        plt.plot(IDR_y_max_DE_avg.iloc[:,-1], IDR_y_max_DE_avg.iloc[:,0], color='k', label='Average', linewidth=2)
        plt.plot(IDR_y_min_DE_avg.iloc[:,-1], IDR_y_min_DE_avg.iloc[:,0], color='k', linewidth=2)
        
        # reference line 그려서 허용치 나타내기
        plt.axvline(x=-cri_DE, color='r', linestyle='--', label='LS')
        plt.axvline(x=cri_DE, color='r', linestyle='--')
        
        # 기타
        plt.yticks(story_name_window_reordered[::yticks], story_name_window_reordered[::yticks])
        plt.grid(linestyle='-.')
        plt.xlabel('Interstory Drift Ratios(m/m)')
        plt.ylabel('Story')
        plt.legend(loc=4, fontsize=8)
        plt.title('H2 DE')
        
        plt.tight_layout()
        # plt.savefig(result_path + '\\' + 'IDR_H2_DE')
        plt.close()
        count += 1
        
        yield fig2

    # DE Plot
    if len(MCE_load_name_list) != 0:
    
        ### H1 MCE 그래프
        fig3 = plt.figure(count, figsize=(5, 7), dpi=150)
        plt.xlim(-0.025, 0.025)
        
        # 지진파별 plot
        for load_name in MCE_load_name_list:
            plt.plot(globals()['IDR_x_max_{}_avg'.format(load_name)].iloc[:,-1], IDR_x_max_MCE_avg.iloc[:,0], label='{}'.format(load_name), linewidth=0.7)
            plt.plot(globals()['IDR_x_min_{}_avg'.format(load_name)].iloc[:,-1], IDR_x_max_MCE_avg.iloc[:,0], linewidth=0.7)
                
        # 평균 plot      
        plt.plot(IDR_x_max_MCE_avg.iloc[:,-1], IDR_x_max_MCE_avg.iloc[:,0], color='k', label='Average', linewidth=2)
        plt.plot(IDR_x_min_MCE_avg.iloc[:,-1], IDR_x_max_MCE_avg.iloc[:,0], color='k', linewidth=2)
        
        # reference line 그려서 허용치 나타내기
        plt.axvline(x=-cri_MCE, color='r', linestyle='--', label='CP')
        plt.axvline(x=cri_MCE, color='r', linestyle='--')
        
        # 기타
        plt.yticks(story_name_window_reordered[::yticks], story_name_window_reordered[::yticks])
        plt.grid(linestyle='-.')
        plt.xlabel('Interstory Drift Ratios(m/m)')
        plt.ylabel('Story')
        plt.legend(loc=4, fontsize=8)
        plt.title('H1 MCE')
        
        plt.tight_layout()
        # plt.savefig(result_path + '\\' + 'IDR_H1_DE')
        plt.close()
        count += 1
        
        yield fig3
        
        # H2 MCE 그래프
        fig4 = plt.figure(count, figsize=(5, 7), dpi=150)
        plt.xlim(-0.025, 0.025)
        
        # 지진파별 plot
        for load_name in MCE_load_name_list:
            plt.plot(globals()['IDR_y_max_{}_avg'.format(load_name)].iloc[:,-1], IDR_y_max_MCE_avg.iloc[:,0], label='{}'.format(load_name), linewidth=0.7)
            plt.plot(globals()['IDR_y_min_{}_avg'.format(load_name)].iloc[:,-1], IDR_y_max_MCE_avg.iloc[:,0], linewidth=0.7)
        
        # 평균 plot
        plt.plot(IDR_y_max_MCE_avg.iloc[:,-1], IDR_y_max_MCE_avg.iloc[:,0], color='k', label='Average', linewidth=2)
        plt.plot(IDR_y_min_MCE_avg.iloc[:,-1], IDR_y_max_MCE_avg.iloc[:,0], color='k', linewidth=2)
        
        # reference line 그려서 허용치 나타내기
        plt.axvline(x=-cri_MCE, color='r', linestyle='--', label='CP')
        plt.axvline(x=cri_MCE, color='r', linestyle='--')
        
        # 기타
        plt.yticks(story_name_window_reordered[::yticks], story_name_window_reordered[::yticks])
        plt.grid(linestyle='-.')
        plt.xlabel('Interstory Drift Ratios(m/m)')
        plt.ylabel('Story')
        plt.legend(loc=4, fontsize=8)
        plt.title('H2 MCE')
        
        plt.tight_layout()
        # plt.savefig(result_path + '\\' + 'IDR_H2_DE')
        plt.close()
        count += 1
        
        yield fig4
        
#%% Pushover

def Pushover(result_path, x_result_txt, y_result_txt, base_SF_design=None, pp_x=None, pp_y=None):
    ''' 

    Perform-3D 해석 결과에서 각 지진파에 대한 Base층의 전단력을 막대그래프 형식으로 출력. (kN)
    
    Parameters
    ----------
    result_path : str
                  Perform-3D에서 나온 해석 파일의 경로.
                  
    result_xlsx : str, optional, default='Analysis Result'
                  Perform-3D에서 나온 해석 파일의 이름. 해당 파일 이름이 포함된 파일들을 모두 불러온다.
                  
    ylim : int, optional, default=70000
           그래프의 y축 limit 값. y축 limit 안의 값만 표기되므로, limit를 넘어가는 값을 확인하고 싶을 시에는 ylim 값을 더 크게 설정하면 된다.
    
    Returns
    -------
    '''
    
#%%

    # 설계밑면전단력 값 입력
    result_path = r'K:\2105-이형우\성능기반 내진설계\GM4R\Pushover_txt'
    x_result_txt = '111_PO_X.txt'
    y_result_txt = '111_PO_Y.txt'
    design_base_shear = 11291 # kN (GEN값 * 0.85)
    pp_DE_x = [0.001864, 7466]
    pp_MCE_x = [0.002373, 8265]
    pp_DE_y = [4.449e-4, 3585]
    pp_MCE_y = [5.403e-4, 3832]
###############################################################################    

    data_X = pd.read_csv(result_path+'\\'+x_result_txt, skiprows=8, header=None)
    data_Y = pd.read_csv(result_path+'\\'+y_result_txt, skiprows=8, header=None)
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
    if design_base_shear != None:
        plt.axhline(design_base_shear, 0, 1, color = 'royalblue', linestyle='--', linewidth = 1.5)
    
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
    
    if design_base_shear != None:
        plt.axhline(design_base_shear, 0, 1, color='royalblue', linestyle='--', linewidth=1.5)

    plt.plot(pp_DE_y[0], pp_DE_y[1], color='r', marker='o')
    plt.text(pp_DE_y[0]*1.3, pp_DE_y[1], 'Performance Point at DE \n ({},{})'.format(pp_DE_y[0], pp_DE_y[1])
             , verticalalignment='top')
    
    plt.plot(pp_MCE_y[0], pp_MCE_y[1], color='g', marker='o')
    plt.text(pp_MCE_y[0]*1.3, pp_MCE_y[1], 'Performance Point at MCE \n ({},{})'.format(pp_MCE_y[0], pp_MCE_y[1])
             , verticalalignment='bottom')
    
    plt.show()
    # yield fig2
    
    print(max(data_X['Base Shear']), design_base_shear, max(data_X['Base Shear'])/design_base_shear)
    print(max(data_Y['Base Shear']), design_base_shear, max(data_Y['Base Shear'])/design_base_shear)