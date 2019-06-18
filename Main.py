# coding=utf-8
from pykeyboard import PyKeyboard
from pymouse import PyMouse
import time
import pyHook
import pythoncom

import  pyperclip
from pynput import mouse,keyboard

def copy():

    k.press_key(k.control_l_key)
    k.tap_key("C")
    k.release_key(k.control_l_key)


# def onMouseEvent(event):
#     if(event.MessageName!="mouse move"):# 因为鼠标一动就会有很多mouse move，所以把这个过滤下
#         print(copy())
#     return True # 为True才会正常调用，如果为False的话，此次事件被拦截
#
# hm = pyHook.HookManager()
#     # 监听鼠标
# hm.MouseAll = onMouseEvent
# hm.HookMouse()
#     # 循环监听
# pythoncom.PumpMessages()


def onpressed(key):
    if(key==keyboard.Key.caps_lock):
        last=pyperclip.paste()
        maxTime=5
        while(pyperclip.paste()==last and maxTime>0):
            maxTime=maxTime-1
            time.sleep(0.5)
            print('doing')
            copy()
        if(maxTime<=0):
            return
        print(pyperclip.paste())

k=PyKeyboard()
m=PyMouse()
with keyboard.Listener(on_press=onpressed) as listener:
    listener.join()


# def on_click(x,y,button,pressed):
#     if(button==mouse.Button.left or button==mouse.Button.right):
#         last = pyperclip.paste()
#         maxTime = 5
#         while (pyperclip.paste() == last and maxTime > 0):
#             maxTime = maxTime - 1
#             time.sleep(0.5)
#             print('doing')
#             copy()
#         print(pyperclip.paste())
#
# with mouse.Listener(on_click=on_click) as listener:
#     listener.join()

