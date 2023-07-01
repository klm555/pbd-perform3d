# -*- coding: utf-8 -*-
"""
Created on Wed Jun  7 19:12:21 2023

@author: hwlee
"""
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib import rc
import numpy as np

# 한국어 나오게하는 코드(이해는 못함)
matplotlib.rcParams['axes.unicode_minus'] = False
font_name = fm.FontProperties(fname='c:\\windows\\fonts\\malgun.ttf').get_name()
rc('font', family=font_name)

# User Input
period_list = [2.2015, 2.7458, 2.6636, 2.3839, 2.8638, 3.6431, 2.8940, 2.6017, 2.7316, 2.4188]
bldg_list = range(101, 111)
num_bldg = len(bldg_list)
min_factor = 0.22
max_factor = 1.5

period_range_list = [[min_factor*i, i, max_factor*i] for i in period_list]
x = np.arange(len(bldg_list))

# Plot
fig = plt.figure(1, dpi=150)

# 각 동의 T1 range plot
for period_range, bldg in zip(period_range_list, x):
    plt.plot([bldg, bldg, bldg], period_range, marker='o', lw=0.3)
    
# T1 max, T1, T1 min 연결
period_range_arr = np.array(period_range_list)
plt.plot(range(num_bldg), period_range_arr[:,0], label='$%s T_1$'%min_factor)
plt.plot(range(num_bldg), period_range_arr[:,1], label='$T_1$')
plt.plot(range(num_bldg), period_range_arr[:,2], label='$%s T_1$'%max_factor)

plt.fill_between(range(-1, num_bldg+1), 0.22, 6, color='red', alpha=0.1)
plt.axhline(y= 0.22, color='r', linestyle='-', lw=0.5)
plt.axhline(y= 6, color='r', linestyle='-', lw=0.5)

plt.xlim(-0.5, num_bldg-0.5)
plt.xticks(range(num_bldg), bldg_list)
plt.yticks([0.22, 1,2,3,4,5,6])

plt.grid(axis='y', linestyle='--', lw=0.5, alpha=0.7)
plt.xlabel('Building')
plt.ylabel('$T_1$ (sec)')
plt.legend()
plt.title('관심주기 영역')

plt.show()