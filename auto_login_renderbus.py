#!/usr/bin/env python
# -*- coding:utf-8 -*-

"""
目标：自动登录RenderBus客户端（Windows）
步骤：
    1.打开客户端（固定位置、注册表、环境变量等获取方式）
    2.找到窗口句柄
    3.鼠标定位和点击删除原有信息
    4.输入新的信息
    5.点击登录

类或方法：
    1.获取客户端位置
    2.通过类名和标题查找窗口句柄，并获得窗口位置和大小、
    # 3.通过父句柄获取子句柄
    4.鼠标定位和点击
    5.键盘发送信息或回车键
"""

import os
import sys
import time
import win32gui
import win32api
import win32con


class SoftwareOperation(object):
    """
    对软件的一些操作
    如：获取软件位置、开启软件、关闭软件、检查软件进程
    """
    def __init__(self, software_name):
        """
        :param software_name: 软件名
        """
        self.software_name = software_name
        self.software_path = self.get_software_position()

    def get_software_position(self):
        """
        根据软件名获取软件路径
        :return: 软件路径
        :rtype: str
        """
        path = r'C:\Users\chensirui\AppData\Roaming\rayvision\RenderBus4.0\QRenderBus.exe'
        return path

    def check_software_process(self):
        """
        检查软件进程是否存在
        大小写敏感
        :return: True/False
        :rtype: bool
        """
        cmd = r'tasklist /FI "IMAGENAME eq {0}"'.format(self.software_name)
        content = os.popen(cmd).read()

        if content.find(self.software_name) > -1:
            result = True
        else:
            result = False

        return result

    def start_software(self):
        """
        判断软件进程是否已经存在，如果不存在则打开软件
        只负责开启软件，不确认是否开启成功
        """
        is_process_exist = self.check_software_process()
        if not is_process_exist:
            cmd = 'start "" "{}"'.format(self.software_path)
            os.system(cmd)

    def stop_software(self):
        """
        kill软件进程
        """
        cmd = r'taskkill /FI "IMAGENAME eq {0}"'.format(self.software_name)
        os.system(cmd)


class WindowOperation(object):
    """
    对窗口的一些操作
    如：查找窗口句柄、通过父句柄获取子句柄
    """
    def __init__(self, class_name, title_name):
        """
        :param class_name: 窗口类名，可用spy++获取
        :param title_name: 窗口标题，可用spy++获取
        """
        self.class_name = class_name
        self.title_name = title_name
        try:
            self.hwnd = self.get_window_handle()
        except:
            self.hwnd = 0

    def get_window_handle(self):
        """
        获取窗口句柄
        :return: 窗口句柄
        :rtype: long
        """
        self.hwnd = win32gui.FindWindow(self.class_name, self.title_name)
        return self.hwnd

    def get_child_window_handle(self):
        """
        获得parent的所有子窗口句柄
        :return: 子窗口句柄列表
        """
        hwnd_chile_list = []
        if self.hwnd > 0:
            try:
                win32gui.EnumChildWindows(self.hwnd, lambda hwnd, param: param.append(hwnd), hwnd_chile_list)
            except Exception as e:
                print e
        return hwnd_chile_list

    def get_windows_title(self):
        """
        获取窗口标题
        :return: 窗口标题
        """
        title_name = win32gui.GetWindowText(self.hwnd)
        return title_name

    def get_windows_class(self):
        """
        获取窗口类名
        :return: 窗口类名
        """
        class_name = win32gui.GetClassName(self.hwnd)
        return class_name

    def get_child_window_by_class_name(self, class_name):
        """
        获取父句柄hwnd类名为class_name的子句柄
        :param class_name: 子类名
        :return:
        """
        hwnd = win32gui.FindWindowEx(self.hwnd, None, class_name, None)
        return hwnd

    def get_window_rect(self):
        """
        获取窗口左上角和右下角坐标
        """
        left, top, right, bottom = win32gui.GetWindowRect(self.hwnd)
        return left, top, right, bottom


class Mouse(object):
    """
    鼠标的一些操作
    如：鼠标定位、左键单击、左键双击、右键单击、右键双击
    """
    def __init__(self):
        pass

    def set_cursor_position(self, width, height):
        """
        设置鼠标位置
        :param width: 宽
        :param height: 高
        :return:
        """
        win32api.SetCursorPos([width, height])

    def left_click(self):
        """
        执行左单键击
        """
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

    def right_click(self):
        """
        执行右单键击
        """
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP | win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)

    def left_double_click(self, interval_time=0.01):
        """
        左键双击
        """
        self.left_multi_click(interval_time=interval_time)

    def right_double_click(self, interval_time=0.01):
        """
        右键双击
        """
        self.right_multi_click(interval_time=interval_time)

    def left_multi_click(self, times=2, interval_time=0.01):
        """
        左键多击
        :param times: 点击次数，默认2次
        :param interval_time: 间隔时长，默认10ms
        """
        for i in range(times):
            self.left_click()
            time.sleep(interval_time)

    def right_multi_click(self, times=2, interval_time=0.01):
        """
        右键多击
        :param times: 点击次数，默认2次
        :param interval_time: 间隔时长，默认10ms
        """
        for i in range(times):
            self.right_click()
            time.sleep(interval_time)


class keyBoard(object):
    """
    键盘的一些操作
    如：输入字符串
    """
    def __init__(self):
        # 参考：https://docs.microsoft.com/en-us/windows/desktop/inputdev/virtual-key-codes
        self.VK_CODE = {
            'backspace': 0x08,
            'tab': 0x09,
            'clear': 0x0C,
            'enter': 0x0D,
            'shift': 0x10,
            'ctrl': 0x11,
            'alt': 0x12,
            'pause': 0x13,
            'caps_lock': 0x14,
            'esc': 0x1B,
            'spacebar': 0x20,
            'page_up': 0x21,
            'page_down': 0x22,
            'end': 0x23,
            'home': 0x24,
            'left_arrow': 0x25,
            'up_arrow': 0x26,
            'right_arrow': 0x27,
            'down_arrow': 0x28,
            'select': 0x29,
            'print': 0x2A,
            'execute': 0x2B,
            'print_screen': 0x2C,
            'ins': 0x2D,
            'del': 0x2E,
            'help': 0x2F,
            '0': 0x30,
            '1': 0x31,
            '2': 0x32,
            '3': 0x33,
            '4': 0x34,
            '5': 0x35,
            '6': 0x36,
            '7': 0x37,
            '8': 0x38,
            '9': 0x39,
            'a': 0x41,
            'b': 0x42,
            'c': 0x43,
            'd': 0x44,
            'e': 0x45,
            'f': 0x46,
            'g': 0x47,
            'h': 0x48,
            'i': 0x49,
            'j': 0x4A,
            'k': 0x4B,
            'l': 0x4C,
            'm': 0x4D,
            'n': 0x4E,
            'o': 0x4F,
            'p': 0x50,
            'q': 0x51,
            'r': 0x52,
            's': 0x53,
            't': 0x54,
            'u': 0x55,
            'v': 0x56,
            'w': 0x57,
            'x': 0x58,
            'y': 0x59,
            'z': 0x5A,
            'numpad_0': 0x60,
            'numpad_1': 0x61,
            'numpad_2': 0x62,
            'numpad_3': 0x63,
            'numpad_4': 0x64,
            'numpad_5': 0x65,
            'numpad_6': 0x66,
            'numpad_7': 0x67,
            'numpad_8': 0x68,
            'numpad_9': 0x69,
            'multiply_key': 0x6A,
            'add_key': 0x6B,
            'separator_key': 0x6C,
            'subtract_key': 0x6D,
            'decimal_key': 0x6E,
            'divide_key': 0x6F,
            'F1': 0x70,
            'F2': 0x71,
            'F3': 0x72,
            'F4': 0x73,
            'F5': 0x74,
            'F6': 0x75,
            'F7': 0x76,
            'F8': 0x77,
            'F9': 0x78,
            'F10': 0x79,
            'F11': 0x7A,
            'F12': 0x7B,
            'F13': 0x7C,
            'F14': 0x7D,
            'F15': 0x7E,
            'F16': 0x7F,
            'F17': 0x80,
            'F18': 0x81,
            'F19': 0x82,
            'F20': 0x83,
            'F21': 0x84,
            'F22': 0x85,
            'F23': 0x86,
            'F24': 0x87,
            'num_lock': 0x90,
            'scroll_lock': 0x91,
            'left_shift': 0xA0,
            'right_shift': 0xA1,
            'left_control': 0xA2,
            'right_control': 0xA3,
            'left_menu': 0xA4,
            'right_menu': 0xA5,
            'browser_back': 0xA6,
            'browser_forward': 0xA7,
            'browser_refresh': 0xA8,
            'browser_stop': 0xA9,
            'browser_search': 0xAA,
            'browser_favorites': 0xAB,
            'browser_start_and_home': 0xAC,
            'volume_mute': 0xAD,
            'volume_Down': 0xAE,
            'volume_up': 0xAF,
            'next_track': 0xB0,
            'previous_track': 0xB1,
            'stop_media': 0xB2,
            'play/pause_media': 0xB3,
            'start_mail': 0xB4,
            'select_media': 0xB5,
            'start_application_1': 0xB6,
            'start_application_2': 0xB7,
            'attn_key': 0xF6,
            'crsel_key': 0xF7,
            'exsel_key': 0xF8,
            'play_key': 0xFA,
            'zoom_key': 0xFB,
            'clear_key': 0xFE,
            '+': 0xBB,
            ',': 0xBC,
            '-': 0xBD,
            '.': 0xBE,
            '/': 0xBF,
            '`': 0xC0,
            ';': 0xBA,
            '[': 0xDB,
            '\\': 0xDC,
            ']': 0xDD,
            "'": 0xDE,
            '`': 0xC0
        }

        # 需要组合按键的字符
        self.VK_CODE2 = {
            '_': 'right_shift|-'  # 以|分隔
        }

    def write(self, write_str=""):
        """
        键盘输入字符串
        :param write_str: 字符串
        :return:
        """
        for c in write_str:
            if c in self.VK_CODE2:
                # 处理下划线_等需要组合按键的字符
                combination_str = self.VK_CODE2.get(c)
                combination_list = combination_str.split('|')
                length = len(combination_list)

                for i in range(length):
                    win32api.keybd_event(self.VK_CODE[combination_list[i]], 0, 0, 0)

                for j in range(length-1, -1, -1):
                    win32api.keybd_event(self.VK_CODE[combination_list[j]], 0, win32con.KEYEVENTF_KEYUP, 0)
            else:
                win32api.keybd_event(self.VK_CODE[c], 0, 0, 0)
                win32api.keybd_event(self.VK_CODE[c], 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
            time.sleep(0.01)


def auto_login_renderbus():
    """
    自动登录RenderBus
    """
    sys.path.insert(0, r'E:\github\account')
    from renderbus_account import USERNAME, PASSWORD

    software_name = 'QRenderBus.exe'
    software_operation_obj = SoftwareOperation(software_name)

    class_name = "Qt5QWindowIcon"
    title_name = "Renderbus Render"
    windows_operation_obj = WindowOperation(class_name, title_name)

    mouse_obj = Mouse()
    keyboard_obj = keyBoard()

    # 1.开启软件
    software_operation_obj.start_software()

    # 2.获取窗口坐标
    while True:
        try:
            # 获取句柄
            hwnd = windows_operation_obj.hwnd
            if hwnd == 0:
                hwnd = windows_operation_obj.get_window_handle()

            if hwnd != 0:
                print hwnd

            # 获取窗口左上角和右下角坐标
            left, top, right, bottom = windows_operation_obj.get_window_rect()
            break
        except Exception as e:
            # 窗口还没打开
            time.sleep(1)

    print left, top, right, bottom

    # 3.输入用户名
    ## 鼠标定位到用户名处
    set_cursor_width = left + 245
    set_cursor_height = top + 145
    mouse_obj.set_cursor_position(set_cursor_width, set_cursor_height)

    ## 执行左单键击
    mouse_obj.left_click()

    ## TODO: 全选（ctrl+A, 双击，End、Shift+HOME）删除（Backspace，Delete）
    # 双击+Delete
    mouse_obj.left_double_click()
    keyboard_obj.write(['del'])  # 可参考keyboard_obj.VK_CODE

    ## 输入用户名
    username = USERNAME
    keyboard_obj.write(username)

    # 4.输入密码
    ## 鼠标定位到密码处
    set_cursor_width2 = left + 245
    set_cursor_height2 = top + 195
    mouse_obj.set_cursor_position(set_cursor_width2, set_cursor_height2)

    ## 执行左单键击
    mouse_obj.left_click()

    ## TODO: 全选（ctrl+A, 双击，End、Shift+HOME）删除（Backspace，Delete）
    # 双击+Delete
    mouse_obj.left_double_click()
    keyboard_obj.write(['del'])  # 可参考keyboard_obj.VK_CODE

    ## 输入密码
    password = PASSWORD
    keyboard_obj.write(password)

    # 6.登录（点击或回车）
    ## 鼠标定位到登录处
    set_cursor_width3 = left + 245
    set_cursor_height3 = top + 275
    mouse_obj.set_cursor_position(set_cursor_width3, set_cursor_height3)

    ## 执行左单键击
    mouse_obj.left_click()


def main():
    auto_login_renderbus()


if __name__ == '__main__':
    main()
