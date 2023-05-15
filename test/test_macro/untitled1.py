
import pyautogui as pag
import pandas as pd
import os

import time
from threading import Thread
from pynput import keyboard

from PIL import ImageGrab
from functools import partial
ImageGrab.grab = partial(ImageGrab.grab, all_screens=True) # pag.locateOnScreen이 듀얼 모니터에서 안되는 문제 해결

import numpy as np
import cv2

# import argparse
import imutils
import mss
# import glob

#%% 

# 자동 좌표 획득

position_test = pag.locateCenterOnScreen('images/test.png'
                                                , confidence=0.8, grayscale=True) 
# test
pag.moveTo(position_test)

obj_target = cv2.imread('images/H1_big_blue.png')
obj_target = cv2.cvtColor(obj_target, cv2.COLOR_BGR2GRAY)
obj_target = cv2.Canny(obj_target, 20, 20)

template = cv2.imread('.png')
template = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
template = cv2.Canny(template, 20, 20)

# Loop over the scales of the image
for scale in np.linspace(0.2, 2, 40):
    # Resize the image according to the scale, and keep track
    # of the ratio of the resizing
    resized = imutils.resize(obj_target, width = int(obj_target.shape[1]*scale))
    r = obj_target.shape[1] / float(resized.shape[1])
    
    # if the resized image is smaller than the obj_target, then break
	# from the loop
    if (resized.shape[0] > obj_target.shape[0]) or (resized.shape[1] > obj_target.shape[1]):
	    break

# (tH, tW) = obj_target.shape[:2]
cv2.imshow("obj_target", obj_target)

# 입력값 받으면 창끄기
if cv2.waitKey(0) & 0xff ==27:  # esc 누르면 꺼짐(esc키 값=27)
    cv2.destroyAllWindows()

result = cv2.matchTemplate(template, obj_target, cv2.TM_SQDIFF_NORMED)

min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

print('Best match top left position: %s' %str(max_loc))
print('Best match confidence: %s' %max_val)

threshold = 0.17
if max_val >= threshold:
    print('Found the match')
else:
    print('Match not found')

locations = np.where(result <= threshold)
locations = list(zip(*locations[::-1]))

if locations:
    print('Found the match')