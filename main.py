import cv2
import numpy as np
import HandTrackingModule as htm
import time
import pyautogui
from pynput.keyboard import Controller, Key
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import math
import win32gui
import win32con

##########################
wCam, hCam = 640, 480
frameR = 100  # 框架缩减
smoothening = 7  # 平滑参数
#########################

pTime = 0
plocX, plocY = 0, 0  # 上一帧的位置
clocX, clocY = 0, 0  # 当前帧的位置
prev_ring_y = 0  # 上一帧无名指y坐标
click_executed = False  # 用于记录是否已经执行点击

# 初始化键盘控制
keyboard = Controller()

# 获取默认音频输出设备
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = interface.QueryInterface(IAudioEndpointVolume)
# 获取音量范围
volRange = volume.GetVolumeRange()
minVol = volRange[0]
maxVol = volRange[1]

# 设置摄像头
cap = cv2.VideoCapture(0)
cap.set(3, wCam)
cap.set(4, hCam)

# 检测器实例化
detector = htm.handDetector(maxHands=1)

# 获取屏幕尺寸
wScr, hScr = pyautogui.size()
print(f"Screen Size: {wScr} x {hScr}")

# 禁用 PyAutoGUI 的 Fail-Safe 功能
pyautogui.FAILSAFE = False

# 初始化控制模式状态
inVolumeControl = False
fiveFingersDetectedTime = 0  # 五指张开时间计时器
fiveFingersCloseDetectedTime = 0  # 五指并拢时间计时器

# 创建窗口并获取窗口句柄
cv2.namedWindow("Image")
hwnd = win32gui.FindWindow(None, "Image")

while True:
    # 1. 获取图像并找到手部
    success, img = cap.read()

    img = cv2.flip(img, 1)

    img = detector.findHands(img)
    lmList, bbox = detector.findPosition(img)

    # 2. 检查是否检测到手部关键点
    if len(lmList)!= 0:
        # 获取食指和大拇指的位置
        x1, y1 = lmList[4][1], lmList[4][2]
        x2, y2 = lmList[8][1], lmList[8][2]
        cx, cy = (x1 + x2) // 2, (y1 + y2) // 2

        # 计算食指和大拇指之间的距离
        length = math.hypot(x2 - x1, y2 - y1)
        length_index_middle, img, lineInfo_index_middle = detector.findDistance(8, 12, img)

        # 控制食指和大拇指的连线颜色变化
        if inVolumeControl:
            # 进入音量控制模式，食指和大拇指连线变绿
            cv2.line(img, (x1, y1), (x2, y2), (0, 255, 0), 3)
        else:
            # 否则是红色
            cv2.line(img, (x1, y1), (x2, y2), (0, 0, 255), 3)

        # 检测手指状态进行其他操作（鼠标控制、点击、滚动等）
        fingers = detector.fingersUp()
        print(f"Fingers Up: {fingers}")

        if fingers == [0, 0, 0, 0, 0]:  # 五指并拢
            if fiveFingersCloseDetectedTime == 0:
                fiveFingersCloseDetectedTime = time.time()  # 记录五指并拢的时间
            elif time.time() - fiveFingersCloseDetectedTime >= 1.0:  # 如果超过0.4秒
                print("Five fingers together detected for 0.7 seconds, pressing Alt + F4 to close window.")
                pyautogui.hotkey('alt', 'f4')  # 模拟按下 Alt + F4 键
                fiveFingersCloseDetectedTime = 0  # 重置计时器
        else:
            fiveFingersCloseDetectedTime = 0  # 五指没有并拢时，重置计时器

        # 判断是否进入音量控制模式
        if length < 50 and fingers[1]!=0:  # 食指和大拇指并拢
            if not inVolumeControl:
                print("Entering Volume Control Mode")
                inVolumeControl = True


        # 进入音量控制模式时，根据手势长度调整音量
        if inVolumeControl:
            # 映射手势长度到音量范围
            vol = np.interp(length, [10, 230], [minVol, maxVol])
            volume.SetMasterVolumeLevel(vol, None)
            print(f"Volume Level: {int(vol)}")

            # 绘制控制图标
            cv2.circle(img, (cx, cy), 10, (0, 255, 0), cv2.FILLED)  # 用绿色绘制圆形
        if length_index_middle < 40:
            if inVolumeControl:
                print("Exiting Volume Control Mode")
                inVolumeControl = False



        # 检测五指张开
        if fingers == [0, 1, 1, 1, 1]:  # 如果五指都张开
            if fiveFingersDetectedTime == 0:
                fiveFingersDetectedTime = time.time()  # 记录五指张开的时间
            elif time.time() - fiveFingersDetectedTime >= 1.0:  # 如果超过1.5秒
                # 模拟 Alt + Tab 切换窗口
                pyautogui.hotkey('win', 'tab')
                print("Win + Tab pressed - Switching Windows")
                fiveFingersDetectedTime = 0  # 重置计时器
        else:
            fiveFingersDetectedTime = 0  # 五指没有张开时，重置计时器



        # 三指滑动滚动模式（食指、中指和无名指）
        if fingers[1] == 1 and fingers[2] == 1 and fingers[3] == 1:
            # 滑动手势的检测
            x_ring, y_ring = lmList[16][1:]  # 获取无名指的y坐标
            if prev_ring_y != 0:
                if y_ring < prev_ring_y - 20:  # 向上滑动
                    # 使用鼠标滚轮向上滑动
                    pyautogui.scroll(-100)
                    print("Scrolling Up")
                elif y_ring > prev_ring_y + 20:  # 向下滑动
                    # 使用鼠标滚轮向下滑动
                    pyautogui.scroll(100)
                    print("Scrolling Down")
            # 更新上一帧的无名指y坐标
            prev_ring_y = y_ring


        # 双指点击模式
        elif fingers[1] == 1 and fingers[2] == 1 and fingers[3] != 1 and fingers[4] == 0:
            # 计算两指之间的距离
            length, img, lineInfo = detector.findDistance(8, 12, img)
            if length < 40:
                # 如果未执行点击，则进行点击
                if not click_executed:
                    cv2.circle(img, (lineInfo[4], lineInfo[5]), 15, (0, 255, 0), cv2.FILLED)
                    pyautogui.click()
                    print("Clicking")
                    click_executed = True  # 标记已点击
            else:
                click_executed = False  # 如果双指分开，重置点击状态

        # 单指移动模式
        elif fingers[1] == 1 :
            print("Moving Mode Activated")
            # 将坐标转换为屏幕坐标
            x1, y1 = lmList[8][1:]
            x3 = np.interp(x1, (frameR, wCam - frameR), (0, wScr))
            y3 = np.interp(y1, (frameR, hCam - frameR), (0, hScr))

            # 平滑移动
            clocX = plocX + (x3 - plocX) / smoothening
            clocY = plocY + (y3 - plocY) / smoothening

            # 移动鼠标
            pyautogui.moveTo(clocX, clocY)
            print(f"Moving Mouse to: ({clocX}, {clocY})")
            cv2.circle(img, (x1, y1), 15, (255, 0, 255), cv2.FILLED)
            plocX, plocY = clocX, clocY

        else:
            prev_ring_y = 0  # 重置上一帧y坐标以避免误触发

    # 计算并显示FPS
    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime
    cv2.putText(img, f'FPS: {int(fps)}', (20, 50), cv2.FONT_HERSHEY_PLAIN, 2, (255, 0, 0), 2)

    # 显示图像
    cv2.imshow("Image", img)

    # 设置窗口置顶
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

    cv2.waitKey(1)
