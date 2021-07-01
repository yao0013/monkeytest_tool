import pyautogui
#实时获取屏幕中鼠标坐标
try:
    while True:
        x, y = pyautogui.position()
        print(x,y)
except KeyboardInterrupt:
    print('\nExit.')