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
        if list[SPL1SP][0] == '+':
            return True
        else:
            return False

    def checkSPL2(self, SPL2SP, list):
        if list[SPL2SP][0] == '\xa0++':
             return True
        else:
             return False

    def checkSPL3(self, SPL3SP, list):
        if list[SPL3SP][0] == '\xa0\xa0+++':
             return True
        else:
             return False

    def checkSPL4(self, SPL4SP, list):
        if list[SPL4SP][0] == '\xa0\xa0\xa0++++':
             return True
        else:
             return False

    def checkSPL5(self, SPL5SP, list):
        if list[SPL5SP][0] == '\xa0\xa0\xa0\xa0+++++':
             return True
        else:
             return False

    def  writeSPL1val(self, sp_readSPL1, sp_writeSPL1, list, writeFilename, writeFilesheet):
        SP1CodeRow = sp_readSPL1
        SP1CodeCol = 4
        SP1NameRow = sp_readSPL1
        SP1NameCol = 5
        # 写子件编码
        ea.writePartInfo(writeFilename, writeFilesheet, sp_writeSPL1, 2, list[SP1CodeRow][SP1CodeCol])
        # 写子件编码
        ea.writePartInfo(writeFilename, writeFilesheet, sp_writeSPL1, 3, list[SP1NameRow][SP1NameCol])

    def  writeSPL2val(self, sp_readSPL2, sp_writeSPL2, list, writeFilename, writeFilesheet):
        SP2CodeRow = sp_readSPL2
        SP2CodeCol = 4
        SP2NameRow = sp_readSPL2
        SP2NameCol = 5
        # 写子件编码
        ea.writePartInfo(writeFilename, writeFilesheet, sp_writeSPL2, 4, list[SP2CodeRow][SP2CodeCol])
        # 写母件编码
        ea.writePartInfo(writeFilename, writeFilesheet, sp_writeSPL2, 5, list[SP2NameRow][SP2NameCol])

    def  writeSPL3val(self, sp_readSPL3, sp_writeSPL3, list, writeFilename, writeFilesheet):
        SP3CodeRow = sp_readSPL3
        SP3CodeCol = 4
        SP3NameRow = sp_readSPL3
        SP3NameCol = 5
        # 写子件编码
        ea.writePartInfo(writeFilename, writeFilesheet, sp_writeSPL3, 6, list[SP3CodeRow][SP3CodeCol])
        # 写母件编码
        ea.writePartInfo(writeFilename, writeFilesheet, sp_writeSPL3, 7, list[SP3NameRow][SP3NameCol])

    def writeSPL4val(self, sp_readSPL4, sp_writeSPL4, list, writeFilename, writeFilesheet):
        SP4CodeRow = sp_readSPL4
        SP4CodeCol = 4
        SP4NameRow = sp_readSPL4
        SP4NameCol = 5
        # 写子件编码
        ea.writePartInfo(writeFilename, writeFilesheet, sp_writeSPL4, 6, list[SP4CodeRow][SP4CodeCol])
        # 写母件编码
        ea.writePartInfo(writeFilename, writeFilesheet, sp_writeSPL4, 7, list[SP4NameRow][SP4NameCol])

    def writeSPL5val(self, sp_readSPL5, sp_writeSPL5, list, writeFilename, writeFilesheet):
        SP5CodeRow = sp_readSPL5
        SP5CodeCol = 4
        SP5NameRow = sp_readSPL5
        SP5NameCol = 5
        # 写子件编码
        ea.writePartInfo(writeFilename, writeFilesheet, sp_writeSPL5, 6, list[SP5CodeRow][SP5CodeCol])
        # 写母件编码
        ea.writePartInfo(writeFilename, writeFilesheet, sp_writeSPL5, 7, list[SP5NameRow][SP5NameCol])

    # 重写母件信息
    def reWriteMPval(self, sp_writeMP, rwMPfilename, CODE, NAME):
        ea.WriteMPCodeval(rwMPfilename, sp_writeMP, 0, CODE);
        ea.WriteMPNameval(rwMPfilename, sp_writeMP, 1, NAME);

    # 重写1级子件件信息
    def reWriteSPL1val(self, rewriteFilename, rewriteFilesheet, sp_writeSPL1, list):
        # 写子件编码
        ea.writePartInfo(rewriteFilename, rewriteFilesheet, sp_writeSPL1, 2, list[4][4])
        # 写母件编码
        ea.writePartInfo(rewriteFilename, rewriteFilesheet, sp_writeSPL1, 3, list[4][5])

    # 重写2级子件件信息
    def reWriteSPL2val(self, rewriteFilename, rewriteFilesheet, sp_writeSPL2, list):
        # 写子件编码
        ea.writePartInfo(rewriteFilename, rewriteFilesheet, sp_writeSPL2, 4, list[5][4])
        # 写母件编码
        ea.writePartInfo(rewriteFilename, rewriteFilesheet, sp_writeSPL2, 5, list[5][5])

    # 重写3级子件件信息
    def reWriteSPL3val(self, rewriteFilename, rewriteFilesheet, sp_writeSPL3, list):
        # 写子件编码
        ea.writePartInfo(rewriteFilename, rewriteFilesheet, sp_writeSPL3, 6, list[6][4])
        # 写母件编码
        ea.writePartInfo(rewriteFilename, rewriteFilesheet, sp_writeSPL3, 7, list[6][5])

    # 重写4级子件件信息
    def reWriteSPL4val(self, rewriteFilename, rewriteFilesheet, sp_writeSPL4, list):
        # 写子件编码
        ea.writePartInfo(rewriteFilename, rewriteFilesheet, sp_writeSPL4, 8, list[7][4])
        # 写母件编码
        ea.writePartInfo(rewriteFilename, rewriteFilesheet, sp_writeSPL4, 9, list[7][5])



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
    indexReadSubPart = 0
    indexWriteSubPart = 0
    sp_SPL1 = 0
    sp_SPL2 = 0
    sp_SPL3 = 0
    sp_SPL4 = 0
    sp_SPL5 = 0


    # 记录前次读取的层级
    Readedlevel = 0

    # 每个子件层级的值 常量
    SPL1 = 1
    SPL2 = 2
    SPL3 = 3
    SPL4 = 4
    SPL5 = 5

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


    # 处理母件
    handStatus, gMPCode, gMPName = ea.handleMP(list)
    # 写母件
    if  handStatus == True:
        ea.WriteMPCodeval(gwrietFilename, gWMPCodeROWsta, gWMPCodeCOLsta, gMPCode);
        ea.WriteMPNameval(gwrietFilename, gWMPNameROWsta, gWMPNameCOLsta, gMPName);
        # 读取源数据表的指针位置
        indexReadSubPart = 4
        # 写目标表的指针位置
        indexWriteSubPart = 1
    else:
        print('母件编码处理异常')


    # 处理1级子件
    while(True):
        if (ea.checkSPL1(indexReadSubPart, list) == True):
            # 读取到了1级子件标识符
            # 执行写1级子件信息
            ea.writeSPL1val(indexReadSubPart, indexWriteSubPart, list, gwrietFilename, gsheetname)
            # 记录当前操作的子件层级
            Readedlevel = SPL1

        elif (ea.checkSPL2(indexReadSubPart, list) == True):
            # 读取到了2级子件标识符
            if (Readedlevel == SPL2):
                indexWriteSubPart += 1
                ea.reWriteMPval(indexWriteSubPart, gwrietFilename, gMPCode, gMPName)
                ea.reWriteSPL1val(gwrietFilename, gsheetname, indexWriteSubPart, list)
            # 执行写2级子件信息
            ea.writeSPL2val(indexReadSubPart, indexWriteSubPart, list, gwrietFilename, gsheetname)
            # 记录当前操作的子件层级
            Readedlevel = SPL2

        elif (ea.checkSPL3(indexReadSubPart, list) == True):
            # 读取到了3级子件标识符
            if (Readedlevel == SPL3):
                indexWriteSubPart += 1
                ea.reWriteMPval(indexWriteSubPart, gwrietFilename, gMPCode, gMPName)
                ea.reWriteSPL1val(gwrietFilename, gsheetname, indexWriteSubPart, list)
                ea.reWriteSPL2val(gwrietFilename, gsheetname, indexWriteSubPart, list)
            # 执行写3级子件信息
            ea.writeSPL3val(indexReadSubPart, indexWriteSubPart, list, gwrietFilename, gsheetname)
            # 记录当前操作的子件层级
            Readedlevel = SPL3

        elif (ea.checkSPL4(indexReadSubPart, list) == True):
            # 读取到了4级子件标识符
            if (Readedlevel == SPL4):
                indexWriteSubPart += 1
                ea.reWriteMPval(indexWriteSubPart, gwrietFilename, gMPCode, gMPName)
                ea.reWriteSPL1val(gwrietFilename, gsheetname, indexWriteSubPart, list)
                ea.reWriteSPL2val(gwrietFilename, gsheetname, indexWriteSubPart, list)
                ea.reWriteSPL3val(gwrietFilename, gsheetname, indexWriteSubPart, list)
            # 执行写4级子件信息
            ea.writeSPL4val(indexReadSubPart, indexWriteSubPart, list, gwrietFilename, gsheetname)
            # 记录当前操作的子件层级
            Readedlevel = SPL4

        elif (ea.checkSPL5(indexReadSubPart, list) == True):
            # 读取到了5级子件标识符
            if (Readedlevel == SPL5):
                indexWriteSubPart += 1
                ea.reWriteMPval(indexWriteSubPart, gwrietFilename, gMPCode, gMPName)
                ea.reWriteSPL1val(gwrietFilename, gsheetname, indexWriteSubPart, list)
                ea.reWriteSPL2val(gwrietFilename, gsheetname, indexWriteSubPart, list)
                ea.reWriteSPL3val(gwrietFilename, gsheetname, indexWriteSubPart, list)
                ea.reWriteSPL4val(gwrietFilename, gsheetname, indexWriteSubPart, list)
            # 执行写5级子件信息
            ea.writeSPL5val(indexWriteSubPart, list, gwrietFilename, gsheetname)
            # 记录当前操作的子件层级
            Readedlevel = SPL5

        else:
            print('编码数据处理完成')
            break;

        # 读取数据的指针向下移动一行
        indexReadSubPart += 1





    # # 处理1级子件
    # while(True):
    #     if ea.checkSPL1(sp_SPL1, list) == True:
    #         # 执行写1级子件信息
    #         ea.writeSPL1val(sp_SPL1, list, gwrietFilename, gsheetname)
    #         # 1级子件位置指针++
    #         sp_SPL1 += 1
    #     else:
    #         sp_SPL1 = 0
    #         break;
    #
    # # 处理2级子件
    # while (True):
    #     if ea.checkSPL2(sp_SPL2, list) == True:
    #         # 执行写2级子件信息
    #         ea.writeSPL2val(sp_SPL2, list, gwrietFilename, gsheetname)
    #         if sp_SPL2 != 0:
    #             ea.reWriteSPL1val(gwrietFilename, gsheetname, sp_SPL2, list)
    #
    #         # 3级子件位置指针++
    #         sp_SPL2 += 1
    #     else:
    #         sp_SPL2 = 0
    #         break;
    #
    # # 处理3级子件
    # while (True):
    #     if ea.checkSPL3(sp_SPL3, list) == True:
    #         # 执行写3级子件信息
    #         ea.writeSPL3val(sp_SPL3, list, gwrietFilename, gsheetname)
    #         if sp_SPL3 != 0:
    #             # 写母件 写1级 写2级
    #             ea.reWriteMPval(sp_SPL3, gwrietFilename, gMPCode, gMPName)
    #             ea.reWriteSPL1val(gwrietFilename, gsheetname, sp_SPL3, list)
    #             ea.reWriteSPL2val(gwrietFilename, gsheetname, sp_SPL3, list)
    #
    #         # 3级子件位置指针++
    #         sp_SPL3 += 1
    #     else:
    #         sp_SPL3 = 0
    #         break;




    # list = ea.read_excel(r'/Users/yangjun/Desktop/BOMblank.xls', sheetname)
    # # 创建一个新的文件，并写入一行数据
    # valueList = ['阿杜 - 烂好人', '阿杜 - 一诺千年', 'Coldplay - Hypnotised', 'Ruth B. - Superficial Love', '杨宗纬、张碧晨 - 凉凉']
    # ea.addSheet(filename, 'Sheet1', 0, valueList)

    # eh = excelHandle()
    # filename = r'/Users/yangjun/Desktop/母件结构表3.xlsx'
    # sheetname = 'Sheet0'
    # eh.read_excel(filename, sheetname)
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
