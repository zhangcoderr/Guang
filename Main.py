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
    def __init__(self,keyArray,results,args,type):
        self.keyArray=keyArray
        self.results=results
        self.args=args
        self.type=type

def getExcelData():
    datas=[]
    excel = xlrd.open_workbook(r"C:\Users\123\Desktop\广联达\安装\guang.xlsx")
    table = excel.sheets()[0]
    rowCount = table.nrows
    colCount = table.ncols

    for i in range(rowCount):
        if(i==0): continue
        keyArray=str(table.cell_value(i,0)).split('$')
        #result=str(table.cell_value(i,1))
        results=str(table.cell_value(i,1)).split('$')
        args=str(table.cell_value(i,2)).split('$')
        type=str(table.cell_value(i,5))
        data=ExcelData(keyArray,results,args,type)
        datas.append(data)

    return datas

def getresult(str):
    print(str)
    datas= getExcelData()

    result=ExcelData('','','','')
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
            result=data

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
            result_data= getresult(pyperclip.paste())
            if(result_data.results==''):
                print('没有这个:'+pyperclip.paste()+' ，需更新表格')
                return
            k.tap_key(k.escape_key)
            k.tap_key(k.left_key)
            k.tap_key(k.left_key)
            k.tap_key(k.down_key)
            for i in range(len(result_data.results)):

                k.type_string(result_data.results[i])
                k.tap_key(k.enter_key)
                k.tap_key(k.enter_key)

                if(len(result_data.args)>0):
                    for j in range(int(result_data.args[i])):
                        print(1111)
                        k.tap_key(k.enter_key)
                        #time.sleep(1)


                    k.tap_key(k.escape_key)
                    k.tap_key(k.left_key)
                    k.tap_key(k.left_key)
                    k.tap_key(k.left_key)
                    k.tap_key(k.left_key)




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

