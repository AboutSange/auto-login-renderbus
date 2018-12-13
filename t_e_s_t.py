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

VK_CODE = {
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
VK_CODE2 = {
    '_': 'right_shift|-'  # 以|分隔
}


def get_software_position(software_name):
    """
    根据软件名获取软件路径
    :param str software_name: 软件名
    :return: 软件路径
    :rtype: str
    """
    path = r'C:\Users\chensirui\AppData\Roaming\rayvision\RenderBus4.0\QRenderBus.exe'
    return path


def start_software(software_name):
    """
    判断软件进程是否已经存在，如果不存在则打开软件
    只负责开启软件，不确认是否开启成功
    :param software_name: 软件名
    """
    path = get_software_position(software_name)
    cmd = 'start "" "{}"'.format(path)
    os.system(cmd)


def check_software_process(software_name):
    """
    检查软件进程是否存在
    大小写敏感
    :param software_name: 软件名
    :return: True/False
    :rtype: bool
    """
    check_cmd = r'tasklist /FI "IMAGENAME eq {0}"'.format(software_name)
    content = os.popen(check_cmd).read()

    if content.find(software_name) > -1:
        result = True
    else:
        result = False

    return result


def key_input(input_str=''):
    """
    模拟键盘操作
    :param input_str:
    :return:
    """
    for c in input_str:
        if c in VK_CODE2:
            # 处理下划线_等需要组合按键的字符
            combination_str = VK_CODE2.get(c)
            combination_list = combination_str.split('|')
            length = len(combination_list)

            for i in range(length):
                win32api.keybd_event(VK_CODE[combination_list[i]], 0, 0, 0)

            for j in range(length-1, -1, -1):
                win32api.keybd_event(VK_CODE[combination_list[j]], 0, win32con.KEYEVENTF_KEYUP, 0)
        else:
            win32api.keybd_event(VK_CODE[c], 0, 0, 0)
            win32api.keybd_event(VK_CODE[c], 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
        time.sleep(0.01)


def main():
    # 1.启动软件
    software_name = 'QRenderBus.exe'
    # start_software(software_name)  # 调试时关闭

    # 2.查找句柄
    class_name = "Qt5QWindowIcon"
    title_name = "Renderbus Render"

    # 3.等软件完全启动
    while True:
        try:
            # 获取句柄
            hwnd = win32gui.FindWindow(class_name, title_name)
            print hwnd

            # 获取窗口左上角和右下角坐标
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            break
        except Exception as e:
            # 窗口还没打开
            time.sleep(1)

    print left, top, right, bottom

    # 4.输入用户名
    ## 鼠标定位到用户名处
    set_cursor_width = left + 245
    set_cursor_height = top + 145
    win32api.SetCursorPos([set_cursor_width, set_cursor_height])

    ## 执行左单键击，若需要双击则延时几毫秒再点击一次即可
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

    ## TODO: 全选（ctrl+A, 双击，End、Shift+HOME）删除（Backspace，Delete）
    ## 输入用户名
    username = 'xxx'
    key_input(username)

    # 5.输入密码
    ## 鼠标定位到密码处
    set_cursor_width2 = left + 245
    set_cursor_height2 = top + 195
    win32api.SetCursorPos([set_cursor_width2, set_cursor_height2])

    ## 执行左单键击，若需要双击则延时几毫秒再点击一次即可
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

    ## TODO: 全选（ctrl+A, 双击，End、Shift+HOME）删除（Backspace，Delete）
    ## 输入密码
    password = 'xxx'
    key_input(password)

    # 6.登录（点击或回车）
    ## 鼠标定位到登录处
    set_cursor_width3 = left + 245
    set_cursor_height3 = top + 275
    win32api.SetCursorPos([set_cursor_width3, set_cursor_height3])

    ## 执行左单键击，若需要双击则延时几毫秒再点击一次即可
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)


if __name__ == '__main__':
    main()
