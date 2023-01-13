import pandas as pd
import matplotlib.pyplot as plt


# 설계밑면전단력 값 입력
result_path = r'K:\2105-이형우\성능기반 내진설계\GM4R\Pushover_txt'
x_result_txt = '105_PO_X.txt'
y_result_txt = '105_PO_Y2.txt'
base_SF_design = 14015*0.85 # kN (GEN값 * 0.85)
pp_DE_x = [1.017e-4, 17390]
pp_MCE_x = [1.907e-4, 21900]
pp_DE_y = [.002068, 14130]
pp_MCE_y = [.002847, 16370]
###############################################################################

def pushover(result_path, x_result_txt, y_result_txt, base_SF_design=None
             , pp_DE_x=None, pp_DE_y=None, pp_MCE_x=None, pp_MCE_y=None):
        
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
    # plt.xlim([0, 0.016])
    plt.ylim([0, max(max(data_X['Base Shear']), base_SF_design)+3000])
    
    # 설계 밑면전단력 그리기
    if base_SF_design != None:
        plt.axhline(base_SF_design, 0, 1, color = 'royalblue', linestyle='--', linewidth = 1.5)
    
    # 성능점 그리기
    if pp_DE_x != None:
        plt.plot(pp_DE_x[0], pp_DE_x[1], color='r', marker='o')
        plt.text(pp_DE_x[0]*1.3, pp_DE_x[1], 'Performance Point at DE \n ({},{})'.format(pp_DE_x[0], pp_DE_x[1])
                 , verticalalignment='top')
    
    if pp_MCE_x != None:
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
    plt.ylim([0, max(max(data_Y['Base Shear']), base_SF_design)+3000])
    
    if base_SF_design != None:
        plt.axhline(base_SF_design, 0, 1, color='royalblue', linestyle='--', linewidth=1.5)
    
    if pp_DE_y != None:
        plt.plot(pp_DE_y[0], pp_DE_y[1], color='r', marker='o')
        plt.text(pp_DE_y[0]*1.3, pp_DE_y[1], 'Performance Point at DE \n ({},{})'.format(pp_DE_y[0], pp_DE_y[1])
                 , verticalalignment='top')
    
    if pp_MCE_y != None:        
        plt.plot(pp_MCE_y[0], pp_MCE_y[1], color='g', marker='o')
        plt.text(pp_MCE_y[0]*1.3, pp_MCE_y[1], 'Performance Point at MCE \n ({},{})'.format(pp_MCE_y[0], pp_MCE_y[1])
                 , verticalalignment='bottom')
    
    plt.show()
    # yield fig2
    
    print(max(data_X['Base Shear']), base_SF_design, max(data_X['Base Shear'])/base_SF_design)
    print(max(data_Y['Base Shear']), base_SF_design, max(data_Y['Base Shear'])/base_SF_design)
    
#%% Execute the code

pushover(result_path, x_result_txt, y_result_txt, base_SF_design=base_SF_design
         , pp_DE_x=pp_DE_x, pp_MCE_x=pp_MCE_x, pp_DE_y=pp_DE_y, pp_MCE_y=pp_MCE_y)