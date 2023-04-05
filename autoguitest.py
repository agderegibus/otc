import pyautogui as pg
import time 

pg.click(183,742)
time.sleep(2)

pg.write("calcu")
time.sleep(3)

pg.press("enter")

time.sleep(5)

cinco = pg.locateCenterOnScreen("cinco.png")

for i in range(15):
    pg.click(cinco)
    time.sleep(0.5)