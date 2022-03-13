import os
import sys
import time
from datetime import datetime
from threading import Thread

import win32gui
import cv2
import pyautogui
import numpy as np
import pydirectinput
import keyboard
from queue import Queue

if os.path.exists('./GenShin_AutoScrape_Artifact.log'):
    os.remove('./GenShin_AutoScrape_Artifact.log')


class Logger(object):
    def __init__(self, filename, stream=sys.stdout):
        self.terminal = stream
        self.log = open(filename, 'a')

    def write(self, message):
        self.terminal.write(message)
        self.terminal.flush()
        self.log.write(message)
        self.log.flush()

    def flush(self):
        pass


sys.stdout = Logger('./GenShin_AutoScrape_Artifact.log', sys.stdout)
sys.stderr = Logger('./GenShin_AutoScrape_Artifact.log', sys.stderr)
print(str(datetime.now()) + '||[INFO]: GenShin_AutoScrape_Artifact start')


def scrape(q):
    para_hld = win32gui.FindWindow(None, "原神")
    if para_hld == 0:
        para_hld = win32gui.FindWindow(None, "Genshin Impact")
    if para_hld is not 0:
        print(str(datetime.now()) + '||[INFO]: found {}'.format(para_hld))
        win32gui.SetForegroundWindow(para_hld)
        time.sleep(0.5)
        left, top, right, bottom = win32gui.GetWindowRect(para_hld)
        flag = False
        exit_flag = False
        cnt = 0
        while True:
            if cnt > 90:
                break
            pydirectinput.leftClick(left + scrape_pos[0], top + scrape_pos[1])
            print(str(datetime.now()) + '||[INFO]: move to scrape button ==> {}-{}, leftClick'.format(
                left + scrape_pos[0],
                top + scrape_pos[1]))
            time.sleep(0.5)
            for i in range(16):
                if q.get():
                    found = False
                    current_pos = (
                        fr[0] + fr[2] // 2 + int(w_interval * (i % 8)), fr[1] + fr[3] // 2 + int(h_interval * (i // 8)))
                    pydirectinput.leftClick(left + current_pos[0], top + current_pos[1])
                    print(str(datetime.now()) + '||[INFO]: move to current artifact ==> {}-{}, leftClick'.format(
                        current_pos[0], top + current_pos[1]))
                    time.sleep(0.2)
                    img = pyautogui.screenshot(
                        region=(left + color_rect[0], top + color_rect[1], color_rect[2], color_rect[3]))
                    time.sleep(0.1)
                    img = cv2.cvtColor(np.asarray(img), cv2.COLOR_RGB2BGR)
                    ic = img[color_rect[3] // 2, color_rect[2] // 2]
                    print(str(datetime.now()) + '||[INFO]: detect color in {}-{} '.format(
                        color_rect[3] // 2, color_rect[2] // 2))
                    if 190 < ic[0] < 230 and 190 < ic[1] < 230 and 190 < ic[2] < 230 and ic[0] < ic[1] < ic[2]:
                        found = True
                    print(str(datetime.now()) + '||[INFO]: can scrape? {} '.format(found))
                    if i == 0 and found is False:
                        flag = True
                        print(str(datetime.now()) + '||[INFO]: found none to scrape')
                        break
                    if found is False:
                        break
                    q.put(1)
                else:
                    exit_flag = True
            if exit_flag:
                break
            if flag:
                pydirectinput.leftClick(left + scrape_pos[0], top + scrape_pos[1])
                print(str(datetime.now()) + '||[INFO]: leftClick {}-{} to return '.format(left + scrape_pos[0],
                                                                                          top + scrape_pos[1]))
                break
            pydirectinput.leftClick(left + confirm_pos[0], top + confirm_pos[1])
            print(str(datetime.now()) + '||[INFO]: leftClick {}-{} to confirm '.format(left + confirm_pos[0],
                                                                                       top + confirm_pos[1]))

            time.sleep(1)
            pydirectinput.leftClick(left + reconfirm_pos[0], top + reconfirm_pos[1])
            print(str(datetime.now()) + '||[INFO]: leftClick {}-{} to reconfirm '.format(left + reconfirm_pos[0],
                                                                                         top + reconfirm_pos[1]))
            time.sleep(1.2)
            pydirectinput.leftClick()
            print(str(datetime.now()) + '||[INFO]: clear ')
            cnt += 1


if __name__ == '__main__':
    ctrl_q = Queue()
    ctrl_q.put(1)
    scrape_pos = (53, 716)  # x_start, y_start
    info_rect = (880, 120, 320, 560)  # x_start,y_start,w,h
    color_rect = (305, 638, 30, 4)  # x_start,y_start,w,h
    fr = (82, 117, 81, 101)  # first_rect: x_start,y_start,w,h
    confirm_pos = (1152, 716)  # confirm scrape position
    reconfirm_pos = (791, 589)  # reconfirm scrape position
    w_interval = 97.5
    h_interval = 116
    ans = str(input(
        '1.本程序适用于 win10，需要以右键管理员身份运行\n'
        '2.请将游戏窗口化并请将分辨率调整至 1280*720，保持游戏窗口在桌面上【完整】【无遮挡】\n'
        '3.打开圣遗物界面，将圣遗物按稀有度由低到高排序\n'
        '4.程序运行过程中，请勿触碰鼠标，可以按 ESC 键退出\n'
        '5.是否开始执行 [yes/no]:\n'
        '1.This program is suitable for win10 and needs to be run as administrator\n'
        '2.Please adjust the resolution to 1280*720, keep the game window on the desktop [full] and [unobstructed]\n'
        '3.Sort the artifacts in order from low star to high star\n'
        '4.When the program is running, please do not touch the mouse, you can press the "ESC" to exit\n'
        '5.ready to run? [yes/no]:\n')).lower()
    if ans == 'yes' or ans == 'y':
        th_scrape = Thread(target=scrape, args=(ctrl_q,), daemon=True)
        th_scrape.start()
        keyboard.wait('esc')
        ctrl_q.put(0)

    print(str(datetime.now()) + '||[INFO]: job done ')
