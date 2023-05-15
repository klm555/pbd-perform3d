# -*- coding: utf-8 -*-
"""
Created on Fri Mar 17 16:59:25 2023

@author: hwlee
"""

import cv2
import numpy as np
import win32gui, win32ui, win32con


win_title = '원격 데스크톱 연결'

class WindowCapture():
    
    
    # properteis
    w = 0
    h = 0
    hwnd = None
    
    def __init__(self, window_name):
        
        self.hwnd = win32gui.FindWindow(None, window_name)
        # hwnd = None
        if not self.hwnd:
            raise Exception('Window not found: %s' % window_name)
        
        # define your monitor width and height
        self.w = 1920
        self.h = 1080
        
    def get_window(self):
        

    def get_screenshot(self):
        
        # get the window image data
        wDC = win32gui.GenWindowDC(self.hwnd)
        dcObj = win32ui.CreateDCFromHandle(wDC)
        cDC = dcObj.CreateCompatibleDC()
        dataBitMap = win32ui.CreateBitmap()
        dataBitMap.CreateCompatibleBitmap(dcObj, self.w, self.h)
        cDC.SelectObject(dataBitMap)
        cDC.BitBlt((0,0), (self.w,self.h), dcObj, (0,0), win32con.SRCCOPY)
        
        # save the screenshot
        # dataBitMap.SaveBitmapFile(cDC, 'debug.bmp')
        signedIntsArray = dataBitMap.GetBitmapBits(True)
        img = np.fromstring(signedIntsArray, dtype='uint8')
        img.shape(self.h, self.w, 4)
        
        # Free Resources
        dcObj.DeleteDC()
        cDC.DeleteDC()
        win32gui.ReleaseDC(self.hwnd, wDC)
        win32gui.DeleteObject(dataBitMap.GetHandle())
        
        # drop the alpha channel, or cv.matchTemplate() will throw an error like:
        # error: (-215:Assertion failed) (depth == CV_8U || depth == CV_32F) && type == _templ.type()
        # && _img.dims() <= 2 in function 'cv::matchTemplate'
        img = img[...,:3]
        
        # make image C_CONTIGUOUS to avoid erros that look like:
        # File ... in draw_rectangles
        # TypeError: an integer is required (got type tuple)
        # see the discussion here:
        # https://github.com/opencv/opencv/issues/14866#issuecomment-580207109
        img = np.ascontiguousarray(img)
        
        return img
    
    
while(True):    
    screenshot = get_screenshot()
    
    cv2.imshow('Real-Time Screenshot', screenshot)
    
    # press 'q' with the output window focused to exit
    # waits 1ms every loop to process key presses
    if cv2.waitKey(1) == ord('q'):
        cv2.destroyAllWindows()
        break

print('Done')

