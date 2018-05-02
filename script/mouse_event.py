# encoding=utf8
import sys
import win32gui

import pyHook


def BlockInput():
    # 自定义监听鼠标事件
    def onMouseEvent(event):
        # 监听鼠标事件
        # print("MessageName%s:" % event.MessageName)
        print("Message%s   :" % event.Message)
        # print("Time%s:" % event.Time)
        print("Window%x    :" % event.Window)

        # print("WindowName%s:" % event.WindowName)
        # print("Position%s:" % event.Position)
        # print("Wheel%s:" % event.Wheel)
        # print("Injected%s:" % event.Injected)
        print("---")
        return True

    # 自定义监听键盘事件
    def onKeyboardEvent(event):
        flag = False  # 保留一键退出功能
        if event.Key == "Escape":  # 若按下esc键则退出程序
            flag = True
            print("esc")
            sys.exit()
        return flag

    hm = pyHook.HookManager()   # 实例化管理对象
    hm.KeyDown = onKeyboardEvent  # 将键盘按下事件改为自定义键盘事件
    hm.MouseLeftDown = onMouseEvent  # 将鼠标事件改为自定义鼠标事件
    hm.HookMouse()  # 生成鼠标钩子
    hm.HookKeyboard()  # 生成键盘钩子

    win32gui.PumpMessages()  # 开始监听


if __name__ == "__main__":
    BlockInput()