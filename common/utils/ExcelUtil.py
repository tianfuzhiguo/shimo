import openpyxl
import xlrd
from xlwt import Pattern
from xlwt import XFStyle

'''
Created on 2022年1月6日
@author: dujianxiao
'''


class ExcelUtil():

    def readExcel(self, file):
        """
        读取用例文件,支持.xls和.xlsx格式
        param file:用例文件
        """
        book = ''
        if file.endswith('xls'):
            book = xlrd.open_workbook(file)
        elif file.endswith('xlsx'):
            book = openpyxl.load_workbook(file)
        return book

    def getValue(self, file, sheet, row, columnNum):
        """
        获取某行某列的值
        param file:用例文件
        param sheet:
        param row:行号
        param columnNum:列号
        """
        if file.endswith('xls'):
            ctype = sheet.cell(row, columnNum).ctype  # 表格的数据类型
            flag = sheet.cell_value(row, columnNum)
            if ctype == 2 and flag % 1 == 0:  # 如果是整形
                flag = int(flag)
            else:
                flag = flag
        elif file.endswith('xlsx'):
            flag = sheet.cell(row=row + 1, column=columnNum).value
            if flag == None:
                flag = ''
        return flag

    def findStr(self, file, sheet, field):
        """
        用于查找特定字符串所在的列号
        param file:用例文件
        param sheet:
        param field:关键字
        """
        try:
            if file.endswith('xls'):
                for i in range(1, self.ncols):
                    re = sheet.cell(1, i).value
                    if re == field:
                        return i
            elif file.endswith('xlsx'):
                for i in range(1, self.ncols):
                    re = sheet.cell(row=2, column=i).value
                    if re == field:
                        return i
        except Exception as e:
            print(e)
            return '未查找到字符串：' + field

    def getSheetNames(self, file):
        """
        获取用例文件中的全部页签名称
        param file:用例文件
        """
        sheetNames = ''
        book = self.readExcel(file)
        if file.endswith('xls'):
            sheetNames = book.sheet_names()
        elif file.endswith('xlsx'):
            sheetNames = book.get_sheet_names()
        return sheetNames

    def filterArr(self, arr, word):
        """
        去除Arr中含有word的字符串
        param arr:数组
        param word:关键字
        """
        return [item for item in arr if str(word) not in str(item)]

    def getArray(self, file, sheet, row, start, end):
        """
        获取某行start到end之间的数组－－原值
        param file:用例文件
        param sheet:
        param row:行号
        param start:
        param end:
        """
        result = []
        if file.endswith('xls'):
            for column in range(start, end):
                ctype = sheet.cell(row, column).ctype  # 表格的数据类型
                cell = sheet.cell_value(row, column)
                if ctype == 2 and cell % 1 == 0:  # 如果是整形
                    cell = int(cell)
                result.append(str(cell))
        elif file.endswith('xlsx'):
            for column in range(start, end):
                value = sheet.cell(row=row + 1, column=column).value
                if value is None:
                    value = ''
                result.append(str(value))
        return result

    def getInitArray(self, file, sheet, row, field, msg):
        """
        取异常数据中的真实值
        param file:用例文件
        param sheet:
        param row:行号
        param field:关键字
        param msg:异常数组
        """
        return [self.getValue(file, sheet, row, int(item)) for item in msg[1:]] if field in str(msg) else ''

    def setCellStyle(self, n):
        """
        设置单元格格式
        """
        pattern = Pattern()
        pattern.pattern = Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = n
        style = XFStyle()
        style.pattern = pattern
        return style
