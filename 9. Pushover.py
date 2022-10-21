import pandas as pd
import matplotlib.pyplot as plt

data_path = r'C:\Users\khpark\Desktop\21-RM-157 시티오씨엘 6단지 성능기반내진설계\퍼폼 모델링\601\Output\data_X' # data_X 폴더 경로
data_X_name = '601_PO_X.txt'
data_Y_name = '601_PO_Y.txt'
fig_save_name_X = '601_PO_X'
fig_save_name_Y = '601_PO_Y'
data_X = pd.read_csv(data_path+'\\'+data_X_name, skiprows=8, header=None)
data_Y = pd.read_csv(data_path+'\\'+data_Y_name, skiprows=8, header=None)
data_X.columns = ['Drift', 'Base Shear']
data_Y.columns = ['Drift', 'Base Shear']

# 설계밑면전단력 값 입력
design_base_shear = 13936 * 0.85 # kN

### 성능곡선 그리기
### X Direction
plt.figure(1, figsize=(8,5))  # 그래프 사이즈
plt.grid()
plt.plot(data_X['Drift'], data_X['Base Shear'], color = 'k', linewidth = 1)
plt.title('Capacity Curve (X-dir)', pad = 10)
plt.xlabel('Reference Drift', labelpad= 10) # fontweight='bold'
plt.ylabel('Base Shear(kN)', labelpad= 10)
plt.xlim([0, max(data_X['Drift'])])
plt.ylim([0, max(data_X['Base Shear'])+3000])

# 설계 밑면전단력 그리기
plt.axhline(design_base_shear, 0, 1, color = 'royalblue', linestyle='--', linewidth = 1.5)

# 그래프 저장
plt.savefig(data_path + '\\' + fig_save_name_X)


### Y Direction
plt.figure(2, figsize=(8,5))  # 그래프 사이즈
plt.grid()
plt.plot(data_Y['Drift'], data_Y['Base Shear'], color = 'k', linewidth = 1)
plt.title('Capacity Curve (Y-dir)', pad = 10)
plt.xlabel('Reference Drift', labelpad= 10) # fontweight='bold'
plt.ylabel('Base Shear(kN)', labelpad= 10)
plt.xlim([0, max(data_Y['Drift'])])
plt.ylim([0, max(data_Y['Base Shear'])+3000])

# 설계 밑면전단력 그리기
plt.axhline(design_base_shear, 0, 1, color = 'royalblue', linestyle='--', linewidth = 1.5)

# 그래프 저장
plt.savefig(data_path + '\\' + fig_save_name_Y)

print(max(data_X['Base Shear']), design_base_shear, max(data_X['Base Shear'])/design_base_shear)
print(max(data_Y['Base Shear']), design_base_shear, max(data_Y['Base Shear'])/design_base_shear)






