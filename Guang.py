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

    def __init__(self,keyArray,results,args,max_args,max_number,excelIndex):
        self.keyArray=keyArray
        self.results=results
        self.args=args
        self.max_args=max_args
        self.max_number=max_number
        self.excelIndex=excelIndex

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
        data=ExcelData(keyArray,results,args,max_args,max_number,i+1)
        datas.append(data)

    return datas

def getresult(string):
    print(string)
    datas= getExcelData()

    result=ExcelData('','','','','',0)
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
                if(data.excelIndex==839):
                    print(1)
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
    mouse_position = m.position()

    #print(data.keyArray)
    if(data.keyArray==['矿物','电缆','-']):

        tapkey(k.function_keys[2])

        # time.sleep(2)
        # print(m.position())
        # time.sleep(10)

        time.sleep(1)
        #m.press(505, 393)
        m.press(290, 251)
        time.sleep(0.5)
        tapkey(k.enter_key)
        print('矿物')
    if(data.keyArray[0]=='电力电缆'):
        if(data.max_number=='10.0' or data.max_number == '10'):
            #k.tap_key(k.function_keys[5])  # Tap F5
            k.tap_key(k.function_keys[2])
            time.sleep(1)
            m.press(291, 287)
            #m.press(503, 443)
            time.sleep(0.5)
            k.tap_key(k.enter_key)
            print('换 10')
        if (data.max_number == '25.0'or data.max_number == '25'):
            k.tap_key(k.function_keys[2])
            time.sleep(1)
            m.press(293, 263)
            #m.press(503, 414)
            time.sleep(0.5)
            k.tap_key(k.enter_key)
            print('换 25')
    #if(data.keyArray==['电缆', '头', '户内干包', '10']):
    if ('户内干包' in data.keyArray and '10' in data.keyArray):
        k.tap_key(k.function_keys[2])
        time.sleep(1)
        #m.press(501, 393)
        m.press(294, 251)
        time.sleep(0.5)
        k.tap_key(k.enter_key)
        print('huan 10')
    #if (data.keyArray == ['电缆', '头', '户内干包', '25']):
    if ('户内干包' in data.keyArray and '25' in data.keyArray):
        k.tap_key(k.function_keys[2])
        time.sleep(1)
        #m.press(505, 336)
        m.press(291, 204)
        time.sleep(0.5)
        k.tap_key(k.enter_key)
        print('huan 25')
    #m.move(mouse_position[0],mouse_position[1])


def onpressed(key):
    if(key==keyboard.Key.caps_lock):
        last=pyperclip.paste()
        maxTime=3
        mouse_position=m.position()
        while(pyperclip.paste()==last and maxTime>0):
            maxTime=maxTime-1
            time.sleep(0.5)
            #print('doing')
            copy()
        if(maxTime>0):
            result_data= getresult(pyperclip.paste())
            if(result_data==''or result_data.results==''or result_data.results==['']):
                print('没有这个:\n'+pyperclip.paste()+' \n，需更新表格')
                return
            # print('表格：')
            # print(result_data.keyArray)
            # print('Index:'+str(result_data.excelIndex))
            # print('表格end')
            k.tap_key(k.escape_key)
            tapkey(k.left_key,6)
            tapkey(k.right_key,3)
            k.tap_key(k.down_key)
            for i in range(len(result_data.results)):

                k.type_string(result_data.results[i])
                k.tap_key(k.enter_key)
                Huan(result_data)#----------------------------------------------------------------------------
                k.tap_key(k.enter_key)

                if(len(result_data.args)>0 and result_data.args!=['']):
                    for j in range(int(result_data.args[i])):
                        k.tap_key(k.enter_key)
                        #time.sleep(1)


                    k.tap_key(k.escape_key)#！！！！/有的文件回车后在工程量计算式这列所以需要把注释去掉！！
                    k.tap_key(k.left_key)  #自己创的文件需要去注释  别人的文件要测试下回车后在哪
                    k.tap_key(k.left_key)
                    k.tap_key(k.left_key)
                    k.tap_key(k.left_key)
        m.move(mouse_position[0],mouse_position[1])
        #print('move x:'+str(mouse_position[0])+'move y:'+str(mouse_position[1]))

k=PyKeyboard()
m=PyMouse()

# time.sleep(4)
# print(m.position())
# p=m.position()
# m.move(p[0],p[1])
# time.sleep(10)

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

