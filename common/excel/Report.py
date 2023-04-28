import openpyxl
import shutil
import xlrd
from xlutils.copy import copy
from common.utils.ExcelUtil import ExcelUtil

'''
@excel结果文件
@author: dujianxiao
'''


class Report(ExcelUtil):

    def createReport(self, reportDate, path, file, sheetNames):
        """
        生成excel结果文件
        @param reportDate: 报告日期
        @param path: 文件路径
        @param file: 用例文件
        @param sheetNames: 用例文件中的全部页签名
        """
        try:
            sheetRes = []
            fileSrc = f'{path}'.replace('/', '\\') + '\\'
            fileRes = f'{fileSrc}result\\{file[:-4]}-{reportDate}-report.xls'
            if file.endswith('xls'):
                shutil.copyfile(fileSrc + file, fileRes)
                book = xlrd.open_workbook(fileRes, formatting_info=True)
                bookRes = copy(book)
                [sheetRes.append(bookRes.get_sheet(item)) for item in sheetNames]
            elif file.endswith('xlsx'):
                fileRes = f'{fileRes}x'
                shutil.copyfile(fileSrc + file, fileRes)
                book = openpyxl.load_workbook(fileRes)
                bookRes = book
                [sheetRes.append(bookRes.get_sheet_by_name(item)) for item in sheetNames]
            return bookRes, sheetRes, fileRes
        except Exception as e:
            print(e)
            fileCheck = f"文件：{fileRes} 正在被其他程序使用"
            print(fileCheck)
            self.setFonts('red', fileCheck)
