import win32gui
import win32api
from operator import itemgetter

# 현재 모니터 정보
monitors = win32api.EnumDisplayMonitors()

# 좌상단 좌표가 (0,0)으로 시작하는 모니터를 주 모니터로 설정
monitor_map = []
for monitor in monitors:
    if monitor[2][0] == 0 & monitor[2][1] == 0:
        monitor_type = 0
    else:
        monitor_type = 1
        
    monitor_map.append({'type': monitor_type, 'handle': monitor[0]
                        , 'left': monitor[2][0], 'top': monitor[2][1]})
    

# 현재 창 정보
active_title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
title = win32gui.GetWindowText(hwnd)


def getActiveWindowHandle(self):
        def callback(hwnd, hwnd_list: list):
            activeTitle = win32gui.GetWindowText(win32gui.GetForegroundWindow())
            title = win32gui.GetWindowText(hwnd)
            if win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd) and title:
                if title == activeTitle:
                    rect = win32gui.GetWindowRect(hwnd)
                    hwnd_list.append((title, hwnd, rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1]))
        output = []
        win32gui.EnumWindows(callback, output)
        return output[0]

    
    


import pyautogui
active_win = pyautogui.getActiveWindow()
current_win = win32gui.GetForegroundWindow()
current_win_name = win32gui.GetWindowText(current_win)

import screeninfo

test = screeninfo.get_monitors()

print(current_win_name)

