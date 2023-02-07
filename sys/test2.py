import pyautogui as pg
import time





time.sleep(3)



返回 = (230, 231)
start = (742, 349)
gap = 25
滚动 = 100


下载 = (293,624)
# while(True):
#     M = pg.position()
#     print(M)


for i in range(20):
    current = (start[0] , start[1]+gap * i)
    pg.click(current)
    time.sleep(1)
    pg.scroll(-1000)
    pg.click(下载)
    time.sleep(2)
    pg.scroll(1000)
    pg.click(返回)

