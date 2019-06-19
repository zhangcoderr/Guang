# coding=utf-8
from pykeyboard import PyKeyboard
from pymouse import PyMouse
import time
import pyHook
import pythoncom
import xlrd
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
class ExcelData:
    def __init__(self,keyArray,result,type):
        self.keyArray=keyArray
        self.result=result
        self.type=type

def getExcelData():
    datas=[]
    excel = xlrd.open_workbook(r"C:\Users\Administrator\Desktop\guang.xlsx")
    table = excel.sheets()[0]
    rowCount = table.nrows
    colCount = table.ncols

    for i in range(rowCount):
        if(i==0): continue
        keyArray=str(table.cell_value(i,0)).split('$')
        result=str(table.cell_value(i,1))
        type=str(table.cell_value(i,4))
        data=ExcelData(keyArray,result,type)
        datas.append(data)

    return datas

def getresult(str):
    print(str)
    datas= getExcelData()

    result=''
    contains_key=False
    for data in datas:

        for key in data.keyArray:
            if(key in str):
                contains_key=True
                continue
            else:
                contains_key=False
                break
        if(contains_key):
            result=data.result

    return result

def onpressed(key):
    if(key==keyboard.Key.caps_lock):
        last=pyperclip.paste()
        maxTime=5
        while(pyperclip.paste()==last and maxTime>0):
            maxTime=maxTime-1
            time.sleep(0.5)
            print('doing')
            copy()
        if(maxTime>0):
            r= getresult(pyperclip.paste())
            if(r==''):
                print('没有这个:'+pyperclip.paste()+' ，需更新表格')
                return
            k.tap_key(k.escape_key)
            k.tap_key(k.left_key)
            k.tap_key(k.left_key)
            k.tap_key(k.down_key)
            k.type_string(r)
            k.tap_key(k.enter_key)
k=PyKeyboard()
m=PyMouse()
print('start')
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

