# -*- coding: utf-8 -*-
"""
Created on Tue May  9 11:56:17 2023

@author: hwlee
"""
import numpy as np

a = np.array([
    
1920	,
5375	,
3875	,
5230	,
3875	,
3875	,
5230	,
4310	,
5375	,
1920	,
5230	,
4430	,
3315	,
2845	,
2845	,
3330	,
5390	,
6290	,
7845	,
4435	,
3395	,
3220	,
3700	,
2125	,
5065	,
5065	,
2125	,
5065	,
2075	,
5390	,
2345	,
4435	,
3255	,

])

a_tiled = np.tile(a, (29,1))
a_final = a_tiled.reshape(-1,1, order='F')
