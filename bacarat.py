import pyautogui as ag
import cv2
import win32gui
import mss
import numpy as np
import win32com.client
import time


def get_image_from_box(box):
    with mss.mss() as sct:
        pos = {"left": box[0], "top": box[1], "width": box[2], "height": box[3]}
        img = np.array(sct.grab(pos))
    return img


class BaccaratMaster:
    def __init__(self):
        pass


class Game:
    def __init__(self):
        self.site_name = 'Betwiz'
        self.window_name = None
        self.window_handle = None
        self.pre_img = None
        self.x1, self.y1, self.x2, self.y2 = None, None, None, None
        self.pos = None
        self.width, self.height = None, None

    def get_win_list(self):
        def callback(hwnd, hwnd_list):
            title = win32gui.GetWindowText(hwnd)
            if win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd) and title:
                hwnd_list.append((title, hwnd))
            return True

        output = []
        win32gui.EnumWindows(callback, output)
        return output

    def get_win_size(self):
        self.x1, self.y1, self.x2, self.y2 = win32gui.GetWindowRect(self.window_handle)
        self.pos = (self.x1, self.y1, self.x2, self.y2)
        self.width = self.x2 - self.x1
        self.height = self.y2 - self.y1

        return self.pos

    def get_win_image(self):
        with mss.mss() as sct:
            pos = {"left": self.x1, "top": self.y1, "width": self.x2 - self.x1, "height": self.y2 - self.y1}
            img = np.array(sct.grab(pos))
        return img

    def set_foreground(self):
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(self.window_handle)

    def get_handle(self):
        for win in self.get_win_list():
            if self.site_name in win[0]:
                self.window_name = win[0]
                self.window_handle = win[1]
                break

    def get_diff_img(self):
        self.get_win_size()
        img = self.get_win_image()

        if self.pre_img is not None:
            diff = img - self.pre_img
            cv2.imwrite("./capture/current.png", diff)
        else:
            print("call again")
        self.pre_img = img

    def resize_window(self, width=1547, height=1167):
        win32gui.MoveWindow(self.window_handle, 0, 0, width, height, True)

    def setting(self):
        self.get_handle()
        self.set_foreground()

        pos = self.get_win_size()
        self.resize_window(1547, 1167)
        img = self.get_win_image()
        cv2.imwrite("./image.png", img)


class Researcher:
    def __init__(self, table_pos, round_n):
        self.table_pos = table_pos
        self.table_width = table_pos[2] - table_pos[0]
        self.table_height = table_pos[3] - table_pos[1]
        self.table_box = (table_pos[0], table_pos[1], self.table_width, self.table_height)
        self.round_n_box = (table_pos[0], table_pos[1], round_n[0]-table_pos[0], round_n[1]-table_pos[1])

    def get_table_img(self):
        img = get_image_from_box(self.table_box)
        cv2.imwrite(f"./game_result/{time.time()}.png", img)
        print("save image")

    def is_table_done(self):
        img = get_image_from_box(self.round_n_box)
        pass


if __name__ == "__main__":
    game_window = Game()
    game_window.setting()
    me = BaccaratMaster()

    stat = Researcher(table_pos=(1035, 380, 1410, 624))
    ag.mouseInfo()
    1062, 399
    # for i in range(100):
    #     stat.get_table_img()
    #
    #     time.sleep(1)







