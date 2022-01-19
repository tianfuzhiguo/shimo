from common.utils.ExcelUtil import ExcelUtil
from xlutils.copy import copy
import xlrd,openpyxl,shutil
'''
@excel结果文件
@author: dujianxiao
'''
class Report(ExcelUtil):
    
    '''
    @生成excel结果文件
    @param reportDate: 
    @param path:文件路径
    @param file:用例文件
    @param sheetNames:用例文件中的全部页签名    
    '''
    def createReport(self,reportDate,path,file,sheetNames):
        try:
            sheetRes=[]
            fileSrc = str(path).replace('/', '\\')+'\\'
            fileRes = fileSrc+'result\\'+file[:-4]+'-'+str(reportDate)+'-report.xls'
            book=self.readExcel(path+'/'+file)
            if file.endswith('xls'):
                shutil.copyfile(fileSrc+file,fileRes)
                bookRes = copy(book)
                book = xlrd.open_workbook(fileRes,formatting_info=True)
                bookRes = copy(book)
                [sheetRes.append(bookRes.get_sheet(item)) for item in sheetNames]
            elif file.endswith('xlsx'):
                fileRes = fileRes+'x'
                shutil.copyfile(fileSrc+file,fileRes)
                book = openpyxl.load_workbook(fileRes)
                bookRes = book
                [sheetRes.append(bookRes.get_sheet_by_name(item)) for item in sheetNames]
            return bookRes,sheetRes,fileRes
        except Exception as e:
            print(e)
            fileCheck='文件：'+fileRes+' 正在被其他程序使用'
            print(fileCheck)
            self.consoleFunc('red',str(fileCheck))
            