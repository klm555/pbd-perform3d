# -*- coding: utf-8 -*-
"""
Created on Wed Nov 29 09:14:34 2023

@author: hwlee
"""

import pandas as pd
import matplotlib.pyplot as plt

fig, ax = plt.subplots()

load_case = ['MCE21', 'MCE22']

## Time History
# for i in load_case:
    # base_SF_sheet = 'base_SF_%s'%i
    # displacement_sheet = 'displacement_%s'%i
    # base_SF = pd.read_excel(r'K:\2108-REZA\Namsan\Displacement-Base_SF.xlsx', sheet_name=base_SF_sheet, skiprows=8, usecols=[2,3])
    # displ = pd.read_excel(r'K:\2108-REZA\Namsan\Displacement-Base_SF.xlsx', sheet_name=displacement_sheet, skiprows=11, usecols=[2,3])

    # # Delete rows from when nan appears
    # delete_idx_base_SF = base_SF.index[pd.isna(base_SF.iloc[:,0]) | (base_SF.iloc[:,0].str.contains('--'))]
    # delete_idx_displ = displ.index[pd.isna(displ.iloc[:,0]) | (displ.iloc[:,0].str.contains('--'))]

    # base_SF = base_SF.iloc[0:delete_idx_base_SF[0], :]
    # displ = displ.iloc[0:delete_idx_displ[0], :]
    
    # ### Do I have to average them?
    # # base_SF_avg = base_SF.mean(axis=1)
    # # displ_avg = displ.mean(axis=1)
    # # ax.plot(displ_avg, base_SF_avg, label='X direction', linewidth=0.5)
        
    # ax.plot(displ.iloc[0:500,0], base_SF.iloc[0:500,0], label='X direction', linewidth=0.5)
    # ax.plot(displ.iloc[0:500,1], base_SF.iloc[0:500,1], label='Y direction', linewidth=0.5)

## Pushover
# PO_X = pd.read_excel(r'K:\2108-REZA\Namsan\Displacement-Base_SF.xlsx', sheet_name='PO_X', skiprows=18, usecols=[2,3])
# delete_idx_PO = PO_X.index[(pd.isna(PO_X.iloc[:,0])) | (PO_X.iloc[:,0].str.contains('--'))]
# PO_X = PO_X.iloc[0:delete_idx_PO[0], :]

# ax.plot(PO_X.iloc[:,0], PO_X.iloc[:,1], label='Pushover')

## Elastic
# ax.axhline(y= 5427.47, label='Linear Elastic')

# Displacement Check
displ_X = pd.read_excel(r'K:\2108-REZA\1-이형우\231220Displacement-Base_SF.xlsx', sheet_name='overturning', skiprows=13, usecols=[15,17])
displ_Y = pd.read_excel(r'K:\2108-REZA\1-이형우\231220Displacement-Base_SF.xlsx', sheet_name='overturning', skiprows=13, usecols=[15,18])
# delete_idx_displ_X = displ_X.index[(pd.isna(displ_X.iloc[:,0])) | (displ_X.iloc[:,0].str.contains('--'))]
# delete_idx_displ_Y = displ_Y.index[(pd.isna(displ_Y.iloc[:,0])) | (displ_Y.iloc[:,0].str.contains('--'))]
# displ_X = displ_X.iloc[0:delete_idx_displ_X[0], :]
# displ_Y = displ_Y.iloc[0:delete_idx_displ_Y[0], :]

ax.plot(displ_X.iloc[:,0], displ_X.iloc[:,1], label='Displacement X')
ax.plot(displ_Y.iloc[:,0], displ_Y.iloc[:,1], label='Displacement Y')

## Decorate Graph
ax.set_xlabel('Time(sec)')
ax.set_ylabel('Story(mm)')
ax.set_title('Displacement')

#%%
## Displacement @ every Story
# displacement_sheet = 'displacement_MCE21_all'
# displacement_value = 2.49

# # Get Node Number and Displacement
# displ_raw = pd.read_excel(r'D:\이형우\남산뷰\Results\Displacement-Base_SF.xlsx', sheet_name=displacement_sheet, usecols=[1,2,3])
# node_num = displ_raw[displ_raw.iloc[:,0] == 'NODE']
# displ = displ_raw[displ_raw.iloc[:,0] == displacement_value]
# displ = pd.concat([node_num.reset_index(drop=True), displ.reset_index(drop=True)], axis=1)
# displ_sliced = displ.iloc[:,[2,4,5]]
# displ_sliced.columns = ['Node', 'D_x', 'D_y']

# # Get Z(mm) from Node Info
# node = pd.read_excel(r'D:\이형우\남산뷰\Results\Displacement-Base_SF.xlsx', sheet_name='Node', usecols=[0,1,2,3])

# # Merge Z(mm) + displacement
# displ_merged = pd.merge(displ_sliced, node)
# displ_merged.sort_values(by='Z(mm)', inplace=True)

# ax.plot(displ_merged['D_x'], displ_merged['Z(mm)'], label='Displacement by Story')

# # Decorate Graph
# ax.set_xlabel('Displacement(mm)')
# ax.set_ylabel('Z(mm)')
# ax.set_title('Displacement')

## Base Shear Force @ every Story
# base_SF_sheet = 'base_SF_MCE21_all'
# base_SF_value = 2.49

# # Get Node Number and Displacement
# base_SF_raw = pd.read_excel(r'D:\이형우\남산뷰\Results\Displacement-Base_SF.xlsx', sheet_name=base_SF_sheet, usecols=[0,1,2])
# story = base_SF_raw[base_SF_raw.iloc[:,0] == '**']
# story = [i.split('-')[0] for i in story.iloc[:,1]]
# base_SF = base_SF_raw[base_SF_raw.iloc[:,1] == base_SF_value]
# base_SF = pd.concat([pd.Series(story), base_SF.reset_index(drop=True)], axis=1)
# base_SF_sliced = base_SF.iloc[:,[0,3]]
# base_SF_sliced.columns = ['Story', 'Base_SF']

# ax.plot(base_SF_sliced['Base_SF'], base_SF_sliced['Story'], label='Base Shear Force by Story')

# Decorate Graph
# ax.set_xlabel('Base Shear Force(kN)')
# ax.set_ylabel('Story')
# ax.set_title('Shear Force')

#%%

ax.grid(True, which='both')
ax.axvline(x=0, color='k')
ax.axhline(y=0, color='k')

# ax.set_xlabel('Displacement(mm)')
# ax.set_ylabel('Base Shear Force (kN)')
# ax.set_title('Base Shear Force - Displacement')
ax.legend()

# Total
# ax2.plot(cell_pos1.iloc[:,2], cell_pos1.iloc[:,3], label='Position1', linewidth=0.5)

