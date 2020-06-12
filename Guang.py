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

        
def getCopy(maxTime=2):
    #maxTime = 3  # 3秒复制 调用copy() 不管结果对错
    while (maxTime > 0):
        maxTime = maxTime - 0.5
        time.sleep(0.5)
        # print('doing')
        copy()

    result = pyperclip.paste()
    return result

class ExcelData:
    
    def __init__(self,keyArray,results,args,max_args,max_number,excelIndex,argType,argValues,argNames):
        self.keyArray=keyArray
        self.results=results
        self.args=args
        self.max_args=max_args
        self.max_number=max_number
        self.excelIndex=excelIndex
        self.argType=argType
        self.argValues=argValues
        if(argValues!=[]):
            try:
                self.firstArgValue=argValues[0]
            except:
                print(self.argValues)

        self.argNames=argNames
        if (argValues != []):
            try:
                self.firstArgName = argValues[0]
            except:
                print(self.argNames)


def getExcelData():
    datas=[]
    #excel = xlrd.open_workbook(r"C:\Users\123\Desktop\广联达\安装\guang.xlsx")
    excel = xlrd.open_workbook(r"C:\Users\Administrator\Desktop\guang - New.xlsx")  # ----------------------------
    table = excel.sheets()[0]
    rowCount = table.nrows
    colCount = table.ncols

    for i in range(rowCount):
       
        #6.10
        if(i==0): continue
        keyArray=str(table.cell_value(i,0)).split('$')
        #result=str(table.cell_value(i,1))
        results=str(table.cell_value(i,1)).split('$')
        args=str(table.cell_value(i,2)).split('$')
        max_args=str(table.cell_value(i,5+2)).split('$')
        max_number=str(table.cell_value(i,6+2))
        
        argArray=str(table.cell_value(i,3)).split('$')
        argType=argArray[0]
        argValues=[]
        if(argArray!=['']):
            argValues=argArray[1].split('/')


        argNames=str(table.cell_value(i,4)).split('$')


        data=ExcelData(keyArray,results,args,max_args,max_number,i+1,argType,argValues,argNames)
        datas.append(data)
        

    return datas

def getresult(targetFeature,targetName):
    #print(targetFeature)
    datas= getExcelData()

    result=ExcelData('','','','','',0,'','','')
    contains_key=False
    last_result=''
    for data in datas:
        #if(data.excelIndex==1469):
            #print('debug')
        if data.keyArray==['']:
            continue
        for key in data.keyArray:
            if(key in targetFeature):
                contains_key=True
                continue
            else:
                contains_key=False
                break
        
        if(contains_key):
            if(data.argNames!=['']):
                contains_key=targetName in data.argNames


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
                if(regex.search(targetFeature)!=None):
                    value_re = str(regex.search(targetFeature).group())
                    left=int(value_re.split(regex_compile)[0])
                    right = int(value_re.split(regex_compile)[1])
                    max=int(data.max_args[1])
                    if((left+right)*2<=max):
                        result = data

            if(data.max_number!=''):
                result=last_result
                re_compile_list=re.findall(re.compile('\d+'),targetFeature)
                max=0
                for number in re_compile_list:
                    if(int(number)>max):
                        max=int(number)
                if (max <= float(data.max_number)):
                    result = data
                    last_result = data



    return result

def Huan(data):
    #old
    mouse_position = m.position()
    #6.10
    
    x=int(data.firstArgValue.split(',')[0])
    y=int(data.firstArgValue.split(',')[1])
    k.tap_key(k.function_keys[2])
    time.sleep(1)
    #m.press(505, 336)
    m.press(x, y)
    time.sleep(0.5)
    k.tap_key(k.enter_key)

    print('huan')


    
def onpressed(key):
    if(key==keyboard.Key.caps_lock):
        print('doing')
        last = pyperclip.paste()
        maxTime = 3
        mouse_position = m.position()
        while (pyperclip.paste() == last and maxTime > 0):
            maxTime = maxTime - 1
            time.sleep(0.5)
            # print('doing')
            copy()


        if(maxTime>0):
            targetFeature=getCopy()
            tapkey(k.escape_key)
            tapkey(k.left_key)
            tapkey(k.down_key)
            tapkey(k.up_key)
            targetName = getCopy()

            isSameTarget = targetName+targetFeature in sameTargets

            if(isSameTarget):
                print('--相同项--')
                k.type_string('SAME!!')
                #可f4 删除，todo
                return

            result_data= getresult(targetFeature,targetName)
            if(result_data==''or result_data.results==''or result_data.results==['']):#无结果
                print('--无匹配--')
                return
           
            # print('表格：')
            # print(result_data.keyArray)
            sameTargets.append(targetName+targetFeature)
            print('Index: '+str(result_data.excelIndex))
            # print('表格end')
            k.tap_key(k.escape_key)




            tapkey(k.left_key,6)
            tapkey(k.right_key,3)
            k.tap_key(k.down_key)
            for i in range(len(result_data.results)):

                k.type_string(result_data.results[i])
                tapkey(k.enter_key)

                if('1' in result_data.argType):
                     Huan(result_data)#----------------------------------------------------------------------------
                tapkey(k.enter_key)
                #去主材 只计安装费
                if('2' in result_data.argType):
                    delete=False
                    for value in result_data.argValues:
                        if(value in targetFeature):
                            delete=True

                    if(delete):
                        tapkey(k.function_keys[4])
                        tapkey(k.enter_key)
                else:
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

# time.sleep(5)
# print(m.position())
# p=m.position()
# m.move(p[0],p[1])
# time.sleep(10)

sameTargets=[]
datas = getExcelData()
print('start')

with keyboard.Listener(on_press=onpressed) as listener:
    listener.join()


