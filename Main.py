# coding=utf-8
from pykeyboard import PyKeyboard
from pymouse import PyMouse
import time
import pyHook
import pythoncom
import xlrd
import  pyperclip
from pynput import mouse,keyboard
import re
import xml.dom.minidom
def copy():

    k.press_key(k.control_l_key)
    k.tap_key("c")
    k.release_key(k.control_l_key)

def tapkey(key,count=1):
    for i in range(0,count):
        k.tap_key(key)
        time.sleep(0.05)

class ExcelData:
    def __init__(self,keyArray,results,args,max_args,max_number):
        self.keyArray=keyArray
        self.results=results
        self.args=args
        self.max_args=max_args
        self.max_number=max_number

def getExcelData():
    datas=[]
    #excel = xlrd.open_workbook(r"C:\Users\123\Desktop\广联达\安装\guang.xlsx")
    excel = xlrd.open_workbook(r"C:\Users\Administrator\Desktop\guang.xlsx")  # ----------------------------
    table = excel.sheets()[0]
    rowCount = table.nrows
    colCount = table.ncols

    for i in range(rowCount):
        if(i==0): continue
        keyArray=str(table.cell_value(i,0)).split('$')
        #result=str(table.cell_value(i,1))
        results=str(table.cell_value(i,1)).split('$')
        args=str(table.cell_value(i,2)).split('$')
        max_args=str(table.cell_value(i,5)).split('$')
        max_number=str(table.cell_value(i,6))
        data=ExcelData(keyArray,results,args,max_args,max_number)
        datas.append(data)

    return datas

def getresult(string):
    print(string)
    datas= getExcelData()

    result=ExcelData('','','','','')
    contains_key=False
    last_result=''
    for data in datas:
        if data.keyArray==['']:
            continue
        for key in data.keyArray:
            if(key in string):
                contains_key=True
                continue
            else:
                contains_key=False
                break
        if (contains_key):
            if(data.max_args==[''] or data.max_args==[]):
                result = data
            if(data.max_args!=[''] and data.max_args!=[] and contains_key):
                #str = '1、名称：止回阀 2、型号：400*250 3、设单独支吊架'
                regex_compile=data.max_args[0]
                if(regex_compile=='*'):
                    compile='\d+\\'+regex_compile+'\d+'
                else:
                    compile='\d+'+regex_compile+'\d+'
                regex = re.compile(compile)
                if(regex.search(string)!=None):
                    value_re = str(regex.search(string).group())
                    left=int(value_re.split(regex_compile)[0])
                    right = int(value_re.split(regex_compile)[1])
                    max=int(data.max_args[1])
                    if((left+right)*2<=max):
                        result = data

            if(data.max_number!=''):
                result=last_result
                re_compile_list=re.findall(re.compile('\d+'),string)
                max=0
                for number in re_compile_list:
                    if(int(number)>max):
                        max=int(number)
                if (max <= float(data.max_number)):
                    result = data
                    last_result = data



    return result

def Huan(data):

    if(data.keyArray==['矿物','电缆','-']):
        tapkey(k.function_keys[2])

        # time.sleep(2)
        # print(m.position())
        # time.sleep(10)

        time.sleep(1)
        m.press(505, 393)
        time.sleep(0.5)
        tapkey(k.enter_key)
    if(data.keyArray[0]=='电力电缆' or data.keyArray==['电缆', '头', '户内干包', '10'] or data.keyArray==['电缆', '头', '户内干包', '25']):
        if(data.max_number=='10.0' or data.max_number == '10'):
            #k.tap_key(k.function_keys[5])  # Tap F5
            k.tap_key(k.function_keys[2])
            time.sleep(1)
            m.press(505, 471)
            time.sleep(0.5)
            k.tap_key(k.enter_key)
            print('huan 10')
        if (data.max_number == '25.0'or data.max_number == '25'):
            k.tap_key(k.function_keys[2])
            time.sleep(1)
            m.press(507, 442)
            time.sleep(0.5)
            k.tap_key(k.enter_key)
            print('huan 20')

def onpressed(key):
    if(key==keyboard.Key.caps_lock):
        last=pyperclip.paste()
        maxTime=3
        mouse_position=m.position()
        while(pyperclip.paste()==last and maxTime>0):
            maxTime=maxTime-1
            time.sleep(0.5)
            print('doing')
            copy()
        if(maxTime>0):
            result_data= getresult(pyperclip.paste())
            if(result_data==''or result_data.results==''or result_data.results==['']):
                print('没有这个:\n'+pyperclip.paste()+' \n，需更新表格')
                return
            k.tap_key(k.escape_key)
            tapkey(k.left_key,6)
            tapkey(k.right_key,3)
            k.tap_key(k.down_key)
            for i in range(len(result_data.results)):

                k.type_string(result_data.results[i])
                k.tap_key(k.enter_key)
                Huan(result_data)
                k.tap_key(k.enter_key)

                if(len(result_data.args)>0 and result_data.args!=['']):
                    for j in range(int(result_data.args[i])):
                        k.tap_key(k.enter_key)
                        #time.sleep(1)


                    k.tap_key(k.escape_key)
                    k.tap_key(k.left_key)
                    k.tap_key(k.left_key)
                    k.tap_key(k.left_key)
                    k.tap_key(k.left_key)
        m.move(mouse_position[0],mouse_position[1])

k=PyKeyboard()
m=PyMouse()
datas = getExcelData()
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

