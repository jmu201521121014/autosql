import xlrd
import  os
import time
from datetime import date,datetime
import re
import pyperclip

#判断是否为数值类型 https://www.jb51.net/article/145009.htm  http://www.zzvips.com/article/150434.html
def is_number(numStr):
    if(not numStr):
        return False
    
    pattern = re.compile(r'^[+-]?[0-9]+\.?[0-9]*$')
    result = pattern.match(numStr)
    if result:
        return True
    else:
        return False

#判断是否为浮点数(检查小数点后是否全为0)
def numIsFloat(num):
    numStr = str(num)
    if(not is_number(numStr)):
        print("判断类型错误" + numStr + "不是数值类型")
        return NAN

    pointIndex = numStr.find('.')   ##查找小数点出现下标
    if(pointIndex <= -1 or pointIndex >= len(numStr) - 1):  ##没找到小数点 或 小数点在最后 不是浮点型
        return False

    for i in range(pointIndex + 1, len(numStr)):
        num = num * 10;             #先将num*10
        remainder = num % 10;       #再对num求余数，取到原小数点后一个数字
        if(remainder != 0):          #若小数点后数字不为0
            return True             #那么是浮点型

    return False        #小数点后数字全部为0, 不是浮点型

def analysisCell(data, paramList, row):
            if (not row or not data):
                return;
            
            #若直接使用col[i]读取时间类型会直接输出float类型（https://www.jb51.net/article/60510.htm）
            if(row.ctype == 3):        #类型3是time类型
                #time_str_tuple = xlrd.xldate_as_tuple(row,data.datemode)      #获取字符串元组时间 (1992, 2, 22, 0, 0, 0)
                #timeStr = date(*time_str_tuple[:3]).strftime("%Y/%m/%d %H:%M:%S")       #转化成字符串 1992/02/22 00:00:00
                
                # https://blog.csdn.net/weixin_30908941/article/details/94787990
                #struct_time_tuple = time.strptime(timeStr,"%Y/%m/%d %H:%M:%S")      #转化struct_time元组时间 (tm_year=1992, tm_mon=2, tm_mday=22, tm_hour=0, tm_min=0, tm_sec=0, tm_wday=3, tm_yday=91, tm_isdst=-1)
                #timestamp = int(time.mktime(struct_time_tuple))                          #转化为时间戳
                #sqlStr = sqlStr + str(timestamp) + ' '

                # https://www.jb51.net/article/65081.htm     https://www.5axxw.com/questions/content/ayt3ha   https://blog.csdn.net/orangleliu/article/details/38476881
                date_time = xlrd.xldate_as_datetime(row.value,data.datemode)      #转化为datetime类型
                timestamp = int(date_time.timestamp())                      #转化为时间戳
                print("参数类型:time " + str(date_time) + " timestamp " + str(timestamp))
                
                paramList.append(str(timestamp))
                
            elif(row.ctype == 1):        #类型1是str类型
                row_strip = row.value.strip()
                if not row_strip:
                    return;

                if not row_strip.startswith("\'"):
                    row_strip = "\'" + row_strip
                if not row_strip.endswith("\'"):
                    row_strip = row_strip + "\'"
                    
                print("参数类型:string " + row_strip)
                paramList.append(row_strip)

            elif(row.ctype == 2):   #类型2是数值类型
                if(not is_number(str(row.value))):     #若取出的值不为纯数字
                    print("python获取数据类型为数值型, 但是实际上不是由纯数字组成: " + str(row.value))
                    return
                else:           #若是数值类型
                    if(not numIsFloat(row.value)):
                        print("参数类型:int " + str(row.value))
                        paramList.append(int(row.value))
                    else:
                        print("参数类型:float " + str(row.value))
                        paramList.append(float(row.value))

            elif(row.ctype == 4):   #4是bool
                if(row.value == 0):
                    print("参数类型:bool false")
                else:
                    print("参数类型:bool true")
                    
                paramList.append(bool(row.value))

            elif(row.ctype == 0 or row.ctype == 5):              #0是空empty 5是错误error
                return

#写入剪切板  https://cloud.tencent.com/developer/ask/31734
def addToClipBoardCmd(text):
    command = 'echo ' + text.strip() + '| clip'
    print("\n" + command)
    os.system(command)

def addToClipBoard(text):
    r = Tk()
    r.withdraw()
    r.clipboard_clear()
    r.clipboard_append(text)
    r.update() # now it stays on the clipboard after the window is closed
    r.destroy()

#https://www.jb51.net/article/141263.htm
def addToClipBoard2(text):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_TEXT, text)
    w.CloseClipboard()

#https://blog.csdn.net/ppdyhappy/article/details/80216959
def addToClipBoard3(text):
    pyperclip.copy(text)

#写入文件
def writeToFile(text):
    fo = open("历史记录.txt", "a")
    fo.write(text + '\n')
    fo.close()

def analysisExec():
    try:
        data = xlrd.open_workbook("autocreatesql.xlsx")
    except FileNotFoundError:
        print("error: 同目录不存在文件 autocreatesql.xlsx")
        return
    
    if not data:
        print("error: 同目录不存在文件 autocreatesql.xlsx")
        return

    if (len(data.sheets()) <= 0):
        print("error: autocreatesql.xlsx 无页签")
        return
    table = data.sheets()[0]             #通过索引顺序获取

    nrow = table.nrows                  #获取行数
    ncols = table.ncols                  #获取列数

    sqlStr = 'SELECT * FROM '

    if ncols <= 0:
        print("error: 第一列应该为表名")
        return

    tableName = table.col_values(0)
    if (not tableName or len(tableName) < 2 or not tableName[1]):
        print("error: 第一列应该为表名")
        return

    print("表名: " + tableName[1] + "\n")
    
    sqlStr = sqlStr + tableName[1] + " WHERE "

    for iCol in range(1, ncols):    #遍历列
        col = table.col_values(iCol)
        if (not col or len(col) < 3):
            continue;

        colName = col[0].strip()    # 列名
        colSqlStr = col[1].strip()  # sql指令
        sqlParam = []               # 指令参数

        rowLen = len(col);
        for iRow in range(2, rowLen):   #遍历列中的行
            row = table.cell(iRow, iCol)
            analysisCell(data,sqlParam,row)

        sqlParamLen = len(sqlParam)
        if(sqlParamLen <= 0): #解析sql语句参数小于等于0, 跳过列
            continue
        
        sqlEndWithBrackets = colSqlStr.endswith('(')
        for i in range(0, sqlParamLen):
            colSqlStr += str(sqlParam[i])
            if(i < sqlParamLen - 1):
                colSqlStr += ','

        if sqlEndWithBrackets:
            colSqlStr += ') '
        else:
            colSqlStr += ' '
            
        print("列:\'{0}\'生成的sql语句为:{1}\n".format(colName, colSqlStr))
        sqlStr += colSqlStr

    sqlStr += ";"
    print("\n生成sql语句:\n" + sqlStr)
    addToClipBoard3(sqlStr)
    writeToFile(sqlStr)
    return

def main():
    analysisExec()
    os.system( 'pause' )

main()
