# -*- coding: utf-8 -*-
"""
   File Name:  main.py
   Author :    liccoo
   Time:       2022/4/22 20:19
"""
import re
import time

import pyautogui
import pyperclip
import xlrd


# 对 command.xls 数据进行检查
def data_check(sheet):
    # cmdType.value  1.0 左键单击   2.0 左键双击   3.0 右键单击   4.0 输入   5.0 等待   6.0 滚轮
    # ctype     空：0
    #           字符串：1
    #           数字：2
    #           日期：3
    #           布尔：4
    #           error：5
    flag = True

    # 检查是否输入了操作数据 (行数 > 1)  ->  True
    if sheet.nrows < 2:
        raise ValueError("command.xls does not enter the operation command")

    # 对 sheet 的单元格数据进行检查  --- 按列进行
    i = 1
    while i < sheet.nrows:
        # 对指令类型进行检查 --- 第一列
        command_type = sheet.row(i)[0]
        if (command_type.value != 1 and command_type.value != 2 and command_type.value != 3 and command_type.value != 4
                and command_type.value != 5 and command_type.value != 6):
            raise ValueError(f"The data in column 1 of row {i} is not of the specified type")

        # 对操作内容进行检查 --- 第二列
        command_vaule = sheet.row(i)[1]
        # 1.若为读图点击类型指令，内容一定为字符串类型且以.png结尾  -> ctype == 1 and re.search('.png$', command_vaule.vaule)
        if command_type.value == 1 or command_type.value == 2 or command_type.value == 3:
            search_result = re.search('.png$', command_vaule.value, flags=re.IGNORECASE)
            if not search_result:
                raise ValueError(f"The data in column 2 of row {i+1} must be the image name suffixed with .png")
        # 2.若为输入类型 ，则输入类容不能为空
        elif command_type.value == 4 and command_vaule.ctype == 0:
            raise ValueError(f"The data in column 2 of row {i+1} cannot be empty")
        # 3.若为等待类型，内容必须为数字
        elif command_type.value == 5 and command_vaule.ctype != 2:
            raise ValueError(f"The data in column 2 of row {i+1} must be a number")
        # 4.若为滚轮类型，则内容必须为数字
        elif command_type.value == 6 and command_vaule.ctype != 2:
            raise ValueError(f"The data in column 2 of row {i+1} must be a number")

        # 对操作类型重复次数进行检查 --- 第三列
        repeat_time = sheet.row(i)[2]  # 只能为 > 0的自然数或-1，-1 代表一直重复
        if repeat_time.value == '':  # 防止为空不能比较，'' 代表重复一次
            repeat_time.value = 1
        if not (repeat_time.ctype == 0 or repeat_time.value > 0 or repeat_time.value == -1):
            raise ValueError(f"The data in column 3 of row {i+1} must be a number less than -1 or empty")

        # 对是否是否需要等待上一个任务进行检查 --- 第四列
        wait = sheet.row(i)[3]  # 只能为 '1' or ''，默认''
        if wait.value != 1 and wait.value != '':
            raise ValueError(f"The data in column 4 of row {i+1} must be a number 1 or empty")

        i += 1

    return flag


# 按照 command.xls 运行指令
def run_command(sheet):
    # cmdType.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮
    i = 1
    while i < sheet.nrows:
        command_type = sheet.row(i)[0]  # 获取第 i 行操作指令的类型
        repeat_time = sheet.row(i)[2].value  # 获取操作类型的重复次数
        wait = sheet.row(i)[3]  # 获取是否是否需要等待上一个任务

        # 根据不同的操作指令类型去执行相关命令
        # 1.单击左键 -> 1
        if command_type.value == 1:
            image = './image/' + sheet.row(i)[1].value   # 获取操作图片名称
            mouse_click(1, 'left', image, repeat_time, wait)
            time.sleep(0.5)  # 操作间隔时间

        # 2.双击左键 -> 2
        elif command_type.value == 2:
            image = './image/' + sheet.row(i)[1].value  # 获取操作图片名称
            mouse_click(2, 'left', image, repeat_time, wait)
            time.sleep(0.5)  # 操作间隔时间

        # 3.单击右键 -> 3
        elif command_type.value == 3:
            image = './image/' + sheet.row(i)[1].value  # 获取操作图片名称
            mouse_click(1, 'right', image, repeat_time, wait)
            time.sleep(0.5)  # 操作间隔时间

        # 4.输入内容 -> 4
        elif command_type.value == 4:
            input_value = sheet.row(i)[1].value
            # 调用键盘输入函数  --- keyboard_input()
            keyboard_input(input_value)
            time.sleep(0.5)  # 操作间隔时间

        # 5.等待 -> 5
        elif command_type.value == 5:
            # 获取等待时间
            wait_time = sheet.row(i)[1].value
            time.sleep(wait_time)

        # 6.鼠标滚轮 -> 6
        elif command_type.value == 6:
            pyautogui.scroll(int(sheet.row(i)[1].value), )

        # 用户干预暂停运行，确认后继续
        current_x_1, current_y_1 = pyautogui.position()
        time.sleep(0.1)
        current_x_2, current_y_2 = pyautogui.position()
        if current_x_1 != current_x_2 and current_y_1 != current_y_2:
            # 检测到鼠标移动则暂停程序，确认后继续运行
            pyautogui.alert(text='用户干预已暂停，确认后继续运行，5s内保持暂停前的应用获得光标', title='', button='OK')
            time.sleep(5)  # 为暂停后扑获光标留取时间

        i += 1


def mouse_click(click_time, mouse_button, image, repeat_time, wait):
    """
    All mouse-like tasks

    Args:
        click_time (int): The number of clicks on the specified key
        mouse_button (str): The left button, right button or middle button of the mouse,
          the corresponding value is 'left', 'middle 'or 'right'
        image (str): Name of the screenshot of icon or area
        repeat_time (int): The number of times the task is repeated.
          The value must be -1 or a natural number greater than 0
        wait (str): Whether to wait for the previous step to complete.Because the time spent
          in the previous step cannot be estimated.The value can only be 'true', the default is 'false'

    Returns:
        None

    """
    secs_between_clicks = 0.3  # 单位 s
    duration_time = 0.3  # 单位 s
    if repeat_time == '':  # 空代表循环一次
        repeat_time = 1
    if wait.value == '':  # 默认不等待上一次任务，不然找不到图片不报错，会一直找
        wait.value = 0
    # 获取截图的中心对应坐标
    location = pyautogui.locateCenterOnScreen(image, confidence=0.9)
    # 根据任务所需的次数进行分类
    if repeat_time > 0:  # 重复次数 >=1 次
        while repeat_time > 0:
            if location is not None:
                # 执行鼠标点击操作
                pyautogui.click(location.x, location.y, click_time,
                                interval=secs_between_clicks, duration=duration_time, button=mouse_button)
            else:
                if wait.value == 1:  # 需要等待上一次任务执行完成
                    while True:
                        if location is not None:
                            # 执行鼠标点击操作
                            pyautogui.click(location.x, location.y, click_time,
                                            interval=secs_between_clicks, duration=duration_time, button=mouse_button)
                            break
                        time.sleep(1)
                else:
                    # 窗口未匹配到图片异常弹框
                    pyautogui.alert(
                        text=f"The window does not find the specified {image}, please check the window or {image}",
                        title='Error', button='OK')
                    # 控制行 Error
                    raise OSError(f"The window does not find the specified {image}, please check the window or {image}")
            repeat_time -= 1
    else:
        while True:  # 无限重复
            if location is not None:
                # 执行鼠标点击操作
                pyautogui.click(location.x, location.y, click_time,
                                interval=secs_between_clicks, duration=duration_time, button=mouse_button)
            else:
                if wait.value == 1:  # 需要等待上一次任务执行完成
                    while True:
                        if location is not None:
                            # 执行鼠标点击操作
                            pyautogui.click(location.x, location.y, click_time,
                                            interval=secs_between_clicks, duration=duration_time, button=mouse_button)
                            break
                        time.sleep(1)
                else:
                    # 窗口未匹配到图片异常弹框
                    pyautogui.alert(
                        text=f"The window does not find the specified {image}, please check the window or {image}",
                        title='Error', button='OK')
                    # 控制行 Error
                    raise OSError(f"The window does not find the specified {image}, please check the window or {image}")


def keyboard_input(input_value):
    # 所有的功能键
    keys = [r'\t', r'\n', r'\r', ' ', '!', '"', '#', '$', '%', '&', "'", '(', ')', '*', '+',
            ',', '-', '.', '/', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', ':', ';',
            '<', '=', '>', '?', '@', '[', r'\\', ']', '^', '_', '`', 'a', 'b', 'c', 'd', 'e',
            'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u',
            'v', 'w', 'x', 'y', 'z', '{', '|', '}', '~', 'accept', 'add', 'alt', 'altleft',
            'altright', 'apps', 'backspace', 'browserback', 'browserfavorites', 'browserforward',
            'browserhome', 'browserrefresh', 'browsersearch', 'browserstop', 'capslock', 'clear',
            'convert', 'ctrl', 'ctrlleft', 'ctrlright', 'decimal', 'del', 'delete', 'divide', 'down',
            'end', 'enter', 'esc', 'escape', 'execute', 'f1', 'f10', 'f11', 'f12', 'f13', 'f14', 'f15',
            'f16', 'f17', 'f18', 'f19', 'f2', 'f20', 'f21', 'f22', 'f23', 'f24', 'f3', 'f4', 'f5', 'f6',
            'f7', 'f8', 'f9', 'final', 'fn', 'hanguel', 'hangul', 'hanja', 'help', 'home', 'insert',
            'junja', 'kana', 'kanji', 'launchapp1', 'launchapp2', 'launchmail', 'launchmediaselect',
            'left', 'modechange', 'multiply', 'nexttrack', 'nonconvert', 'num0', 'num1', 'num2', 'num3',
            'num4', 'num5', 'num6', 'num7', 'num8', 'num9', 'numlock', 'pagedown', 'pageup', 'pause',
            'pgdn', 'pgup', 'playpause', 'prevtrack', 'print', 'printscreen', 'prntscrn', 'prtsc', 'prtscr',
            'return', 'right', 'scrolllock', 'select', 'separator', 'shift', 'shiftleft', 'shiftright', 'sleep',
            'stop', 'subtract', 'tab', 'up', 'volumedown', 'volumemute', 'volumeup', 'win', 'winleft',
            'winright', 'yen', 'command', 'option', 'optionleft', 'optionright']

    # 根据键入的内容是普通的文本还是功能键进行相应的划分
    if input_value in keys:
        pyautogui.press(input_value)  # 所有键盘的功能键
    else:  # 普通文本
        pyperclip.copy(input_value)
        pyautogui.hotkey('ctrl', 'v')
        # interval_time = 0.3  # 输入字符时每个字符之间的间隔时间
        # pyautogui.typewrite(input_value, interval=interval_time)  # 键入 input_value


if __name__ == '__main__':
    file_name = 'command.xls'
    # 用 xlrd 打开 command.xls 文件
    workbook = xlrd.open_workbook(filename=file_name)

    # 通过索引获取 command.xls 文件相应的 sheet页 --- 默认Sheet1
    sheet_1 = workbook.sheet_by_index(0)  # 获取 Sheet1

    # pyautogui 保护措施  ---  光标位于屏幕左上角时就会退出
    pyautogui.FAILSAFE = True
    # 为所有 PyAutoGUI 函数增加0.5秒的延迟
    pyautogui.PAUSE = 0.5

    # 是否运行程序的消息弹框
    conform = pyautogui.confirm(text='开始后，鼠标光标移至屏幕最左上角可终止程序', title='', buttons=['OK', 'Cancel'])
    if conform == 'OK':
        # 检查 command.xls 文件的输入数据是否有误
        check_command = data_check(sheet_1)

        # print('command.xls 文件的输入数据无误')  # 调试专用
        if check_command:
            run_command(sheet_1)  # 按照 'command.xls' 文件执行命令
        else:
            raise ValueError("The data of command.xls does not meet the specified requirements")

    # 程序运行结束弹窗
    if conform == 'OK':
        pyautogui.alert(text='运行完毕', title='', button='OK')
