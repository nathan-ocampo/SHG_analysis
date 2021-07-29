# -*- coding: utf-8 -*-
"""
Created on Wed Jun 30 12:06:05 2021

@author: natha
"""
#import different packages
#you might have to download the packages

import pyautogui as pg
import time
from pynput.keyboard import Key, Controller


#make controller object
keyboard = Controller()


#gives 5 seconds to click on imagej window near heart
time.sleep(5)

#if want to stop auto clicking move cursor to top left of computer
pg.FAILSAFE = True

#auto clicks 150 times
for i in range(300):    
    pg.click(button='left')
    time.sleep(.05)
    keyboard.press('m')
    time.sleep(.05)
    keyboard.press('.')