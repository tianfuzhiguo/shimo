import openpyxl
import shutil
import xlrd
import os
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
            
    def createHTMLReport(self, js, date, path, file):
        """
        创建html测试报告
        :param path: 文件路径
        :param file: 文件名
        :param js: json格式的测试结果
        @param date: 报告日期
        """
        try:
            html_template = self.resource_path("source/template")
            with open(html_template, "r", encoding="utf-8") as f:
                html_data = f.read()
            html = html_data.replace('${resultData}', str(js))
            file_name = file[:file.index('.xls')]
            html_file = f"{path}/result/{file_name}-{date}-report.html"
            if os.path.exists(html_file):
                os.remove(html_file)
            with open(html_file, 'w', encoding='utf-8') as f:
                f.write(html)
            return html_file
        except Exception as e:
            print(e)
