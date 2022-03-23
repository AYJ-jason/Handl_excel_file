# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
# import time
# import traceback

# import numpy as np
import xlrd
import xlwt
# import numpy
# import pandas as pd
# import datetime
from xlutils.copy import copy
# from xlrd import xldate_as_tuple
# import traceback
# import sys

def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.



class ExcelAction:
    '''
    只支持xls格式
    '''

    def transCode(self, filename, sheetname):
        try:
            filename = filename.decode('utf-8')
            sheetname = sheetname.decode('utf-8')
        except Exception:
            # print traceback.print_exc()
            return  filename, sheetname

    def read_excel(self, filename, sheetname):
        '''
        读取excel
        '''
        filename, sheetname = self.transCode(filename, sheetname)
        workbook = xlrd.open_workbook(filename)  # 获得工作薄
        sheet = workbook.sheet_by_name(sheetname)  # 获得sheet
        rows = sheet.nrows  # 文件总行数
        list = []
        print (u'-------文件内容-------')
        for i in range(0, rows):
            line = sheet.row_values(i)  # 获得一行的值，返回列表
            list.append(line)
            # 避免打印包含中文的列表时变成unicode
            print ('[' + ','.join("'" + str(element) + "'" for element in line) + ']')
        print (u'-----------------------')
        return list

    def write_excel(self, filename, sheetname, row, col, value, type=0):
        '''
        修改excel
        '''
        filename, sheetname = self.transCode(filename, sheetname)
        # 转成整形是因为要在ride中使用，ride把参数传过来默认是字符串，除非这样传${1}
        row = int(row)
        col = int(col)
        type = int(type)
        # formatting_info=True保存之前数据的格式
        rb = xlrd.open_workbook(filename, formatting_info=True)
        wb = copy(rb)
        sheet = wb.get_sheet(sheetname)
        # 设置样式，写入的文字为红色加粗
        style = xlwt.easyxf('font: bold 1, color red;')
        if type == 1:
            sheet.write(row, col, value, style)
        else:
            sheet.write(row, col, value)
        wb.save(filename)

    def addSheet(self, filename, sheetname, row, valueList):
        '''
        写入excel,一次写一行
        '''
        filename, sheetname = self.transCode(filename, sheetname)

        wb = xlwt.Workbook(filename)
        # 其实会覆盖第一个sheet页
        sheet = wb.add_sheet(sheetname)
        for i in range(len(valueList)):
            # 需要转码
            # valueList[i] = str(valueList[i]).deCode('utf-8')#这一行编译不过去
            sheet.write(row, i, valueList[i])
        wb.save(filename)

    # 写母件和子件信息
    def writePartInfo(self, filename, sheetname, row, col, value):
        rb = xlrd.open_workbook(filename, formatting_info=True)
        wb = copy(rb)
        sheet = wb.get_sheet(sheetname)
        sheet.write(row, col, value)
        wb.save(filename)

    # 检查目标文件的合法性
    def checkFile(self, filename):
        data = xlrd.open_workbook(filename)
        table = data.sheets()[0]
        nrows = table.nrows  # 行数
        ncols = table.ncols  # 列数
        if nrows < 3 and ncols <2:
            result: bool = False
        else:
            result: bool = True

        return result, nrows, ncols

    def handleMP(self, list):
        # 定义母件编码字符串内容
        gMPCodeSTR = '母件编码'

        # 定义"母件编码"字符串行位置
        gMPCodeROW = 1
        # 定义"母件编码"字符串列位置
        gMPCodeCOL = 0

        # 读 母件编码行位置
        gRMPCodeNumROW = 2
        # 读 母件编码列位置
        gRMPCodeNumCOL = 0
        # 读 母件名称行位置
        gRMPCodeNameROW = 2
        # 读 母件名称列位置
        gRMPCodeNameCOL = 1

        # 根据目标表格内容获取母件名称'母件编码'
        if list[gMPCodeROW][gMPCodeCOL] == '母件编码':
            return True, list[gRMPCodeNumROW][gRMPCodeNumCOL], list[gRMPCodeNameROW][gRMPCodeNameCOL]
        else:
        # 根据目标表格内容获取并返回母件 编码 内容 取母件 名称 内容
            return False, 0, 0

    # 写母件编码到指定位置
    def WriteMPCodeval(self, writeFilename, MPCodeROW, MPCodeCOL, MPcodeVal,):
        ea.writePartInfo(writeFilename, 'Sheet0', MPCodeROW, MPCodeCOL, MPcodeVal)
        return

    # 写母件名称到指定位置
    def WriteMPNameval(self, writeFilename, MPNameROW, MPnameCOL, MPnameVal):
        ea.writePartInfo(writeFilename, 'Sheet0', MPNameROW, MPnameCOL,MPnameVal)
        return

    def checkSPL1(self, SPL1SP, list):
        if list[SPL1SP + 4][0] == '+':
            return True
        else:
            return False

    def  writeSPL1val(self, sp_SPL1, list, writeFilename, writeFilesheet):
        SP1CodeRow = sp_SPL1 + 4
        SP1CodeCol = 4
        SP1NameRow = sp_SPL1 + 4
        SP1NameCol = 5
        # 写子件编码
        ea.writePartInfo(writeFilename, writeFilesheet, sp_SPL1+1, 2, list[SP1CodeRow][SP1CodeCol])
        # 写母件编码
        ea.writePartInfo(writeFilename, writeFilesheet, sp_SPL1+1, 3, list[SP1NameRow][SP1NameCol])

if __name__ == '__main__':


    # 写 母件编码行位置
    gWMPCodeROWsta = 1
    # 写 母件编码列位置
    gWMPCodeCOLsta = 0
    # 写 母件名称行位置
    gWMPNameROWsta = 1
    # 写 母件名称列位置
    gWMPNameCOLsta = 1

    # 子件的指针
    sp_SPL1 = 0
    sp_SPL2 = 0
    sp_SPL3 = 0
    sp_SPL4 = 0
    sp_SPL5 = 0


    ea = ExcelAction()
    gsourceFilename = r'/Users/yangjun/Desktop/母件结构表2.xlsx'
    gwrietFilename = r'/Users/yangjun/Desktop/BOMblank.xls'
    gsheetname:str = 'Sheet0'



    # 获取源表格行列数量核对结果，获取总行数和总列数
    checkRsl, fileTotalRow, fileTotalCol = ea.checkFile(gsourceFilename)
    if checkRsl == 0:
        print('file is Non-compliance')
    else:
    # 读取目标表格内容
        list = ea.read_excel(gsourceFilename, gsheetname)


    # 处理母件 2022
    handStatus, gMPCode, gMPName = ea.handleMP(list)
    # 写母件
    if  handStatus == True:
        ea.WriteMPCodeval(gwrietFilename, gWMPCodeROWsta, gWMPCodeCOLsta, gMPCode);
        ea.WriteMPNameval(gwrietFilename, gWMPNameROWsta, gWMPNameCOLsta, gMPName);
    else:
        print('母件编码处理异常')

    # 处理1级子件
    while(True):
        if ea.checkSPL1(sp_SPL1, list) == True:
            # 执行写1级子件信息
            ea.writeSPL1val(sp_SPL1, list, gwrietFilename, gsheetname)
            # 1级子件位置指针++
            sp_SPL1 += 1
        else:
            sp_SPL1 = 0
            break;

    # 处理2级子件
    while (True):
        break;

    # 处理3级子件
    while (True):
        break;

    # 处理4级子件
    while (True):
        break;

    # 处理5级子件
    while (True):
        break;






    # list = ea.read_excel(r'/Users/yangjun/Desktop/BOMblank.xls', sheetname)
    # # 创建一个新的文件，并写入一行数据
    # valueList = ['阿杜 - 烂好人', '阿杜 - 一诺千年', 'Coldplay - Hypnotised', 'Ruth B. - Superficial Love', '杨宗纬、张碧晨 - 凉凉']
    # ea.addSheet(filename, 'Sheet1', 0, valueList)

    # eh = excelHandle()
    # filename = r'/Users/yangjun/Desktop/母件结构表3.xlsx'
    # sheetname = 'Sheet0'
    # eh.read_excel(filename, sheetname)
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
