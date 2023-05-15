# -*- coding: utf-8 -*-
"""
Created on Sat Mar 18 12:35:54 2023

@author: hwlee
"""

from screeninfo import get_monitors()

test = get_monitors()

import win32gui
current_win = win32gui.GetForegroundWindow()
current_win_name = win32gui.GetWindowText(current_win)
