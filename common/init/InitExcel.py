import os
import shutil

from common.utils.ExcelUtil import ExcelUtil

'''                                                                                                                                                                                                                                                                                                         
@author: dujianxiao                                                                                                                                                                                                                                                                                         
'''


class InitExcel(ExcelUtil):

    def getBook(self, path, file):
        """
        读用例
        """
        book = ''
        try:
            book = self.readExcel(path + '/' + file)
            return book
        except Exception as e:
            print(e)
            self.consoleFunc('red', str(e))
            return book

    def getSheet(self, date, path, sheetName, file, book):
        """
        获取页签及其行数、列数
        """
        try:
            book = self.readExcel(path + '/' + file)
        except Exception as e:
            self.consoleFunc('red', str(e))
        try:
            # 生测试报告，历史报告移动到history中
            if file.endswith('xls'):
                sheet = book.sheet_by_name(sheetName)
                nrows = sheet.nrows
                ncols = sheet.ncols
            elif file.endswith('xlsx'):
                sheet = book.get_sheet_by_name(sheetName)
                nrows = sheet.max_row
                ncols = sheet.max_column
            try:
                isExists = os.path.exists(path + '/result/history')
                if not isExists:
                    os.makedirs(path + '/result/history')
                fileList = os.listdir(path + '/result/')
                for i in range(len(fileList)):
                    if str(date) not in str(fileList[i]) and str('report') in str(fileList[i]):
                        try:
                            shutil.move(path + '/result/' + str(fileList[i]), path + '/result/history')
                        except Exception as e:
                            print(e)
                return sheet, nrows, ncols
            except Exception as e:
                print(e)
        except Exception as e:
            print(e)

    def getColumn(self, file, sheet):
        """
        获取sheet页各标志位的列号
        param file:用例文件
        param sheet:
        """
        column = [self.findStr(file, sheet, 'name'), self.findStr(file, sheet, 'url'),
                  self.findStr(file, sheet, 'method'), self.findStr(file, sheet, 'param'),
                  self.findStr(file, sheet, 'file'), self.findStr(file, sheet, 'header'),
                  self.findStr(file, sheet, 'part101'), self.findStr(file, sheet, 'part201'),
                  self.findStr(file, sheet, 'part301'), self.findStr(file, sheet, 'section101'),
                  self.findStr(file, sheet, 'section201'), self.findStr(file, sheet, 'section301'),
                  self.findStr(file, sheet, 'resText'), self.findStr(file, sheet, 'resHeader'),
                  self.findStr(file, sheet, 'statusCode'), self.findStr(file, sheet, 'expression'),
                  self.findStr(file, sheet, 'status'), self.findStr(file, sheet, 'time'),
                  self.findStr(file, sheet, 'init001'), self.findStr(file, sheet, 'restore001'),
                  self.findStr(file, sheet, 'dyparam001'), self.findStr(file, sheet, 'key001'),
                  self.findStr(file, sheet, 'value001'), self.findStr(file, sheet, 'headerManager'),
                  self.findStr(file, sheet, '数据库'), self.findStr(file, sheet, 'Iteration')]
        return column
