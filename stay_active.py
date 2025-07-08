import pyautogui, time, os


while True:
    pyautogui.moveTo(900,200,duration=1)
    time.sleep(2)
    pyautogui.moveTo(1600,500,duration=1)
    time.sleep(2)
    pyautogui.press("left")
    time.sleep(2)