import pyautogui as ag
import cv2
import win32gui
import mss
import numpy as np
import win32com.client
import time


def get_win_list():
    def callback(hwnd, hwnd_list: list):
        title = win32gui.GetWindowText(hwnd)
        if win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd) and title:
            hwnd_list.append((title, hwnd))
        return True
    output = []
    win32gui.EnumWindows(callback, output)
    return output


def get_win_size(hwnd):
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    return left, top, right, bottom


def get_win_image(x1, y1, x2, y2):
    # img = cv2.cvtColor(np.array(ImageGrab.grab(bbox=(x1, y1, x2, y2))), cv2.COLOR_RGB2BGR) # PIL VER
    with mss.mss() as sct: # mss로 캡처 수정 - 2022.03.08
        pos = {"left":x1, "top":y1, "width":x2-x1, "height":y2-y1}
        img = np.array(sct.grab(pos))
    return img


def get_image_from_box(box):
    # img = cv2.cvtColor(np.array(ImageGrab.grab(bbox=(x1, y1, x2, y2))), cv2.COLOR_RGB2BGR) # PIL VER
    with mss.mss() as sct: # mss로 캡처 수정 - 2022.03.08
        pos = {"left":box[0], "top":box[1], "width":box[2], "height":box[3]}
        img = np.array(sct.grab(pos))
    return img





def set_foreground(hwnd):
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('%')
    win32gui.SetForegroundWindow(hwnd)


def game_state(pre_img):
    for win in get_win_list():
        if 'Betwiz' in win[0]:
            window_name = win[0]
            window_handle = win[1]
            break
    set_foreground(window_handle)
    x1, y1, x2, y2 = get_win_size(window_handle)
    img = get_win_image(x1, y1, x2, y2)
    if pre_img is not None:
        diff = img - pre_img
    pre_img = img
    cv2.imwrite("./capture/current.png",pre_img)
    case = 5
    if ag.locateOnScreen("./capture/selection.png", confidence=0.8) is not None:
        case = 3  # 숫자 선택 게임 창
    elif ag.locateOnScreen("./capture/choice.png", confidence=0.8) is not None:
        case = 2  # 추첨중
    elif ag.locateOnScreen("./capture/purchaced_log_odd.png", confidence=0.8) is not None:
        case = 1  # 구매완료
    elif ag.locateOnScreen("./capture/before_purchase.png", confidence=0.9) is not None:
        case = 0  # 미구매

    return case


def change_summation_game():
    button = ag.locateOnScreen("./capture/summation_game.png", confidence=0.9)
    ag.click(button, duration=0.25)


def purchase(commands: list):
    for cmd in commands:
        button = ag.locateOnScreen(f"./capture/{cmd}.png", confidence=0.9)
        ag.click(button, duration=0.25)

    button = ag.locateOnScreen("./capture/purchase.png", confidence=0.9)
    ag.click(button, duration=0.25)
    button = ag.locateOnScreen("./capture/purchase_check.png", confidence=0.9)
    ag.click(button, duration=0.25)


def get_my_money():
    button = ag.locateOnScreen("./capture/money.png", confidence=0.9)
    print(button)
    img = get_image_from_box(button)
    cv2.imshow("deposit", img)
    cv2.waitkey(0)
    return


# pre_img = None
# pre_state = None
# while True:
#     is_end = False
#     pre_img = state = game_state(pre_img)
#     money = get_my_money()
#     print(f"{state}", end="")
#
#     if state == 0:
#         bets = ["odd"]
#         purchase(bets)
#         print(f"\t {bets} 구매 완료")
#         on_the_table = 1000
#     elif state == 1:
#         print("추첨 대기 중", end="")
#     elif state == 2:
#         print("추첨 중", end="")
#     elif state == 3:
#         change_summation_game()
#         print("now we are in selection game")
#         if pre_state == 2:
#             print("win ", ag.locateOnScreen("./capture/win.png", confidence=0.9))
#             print("lost ", ag.locateOnScreen("./capture/lost.png", confidence=0.9))
#     elif state == 4:
#         pass
#
#     time.sleep(20)