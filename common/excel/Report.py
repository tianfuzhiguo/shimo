from common.utils.Util import findStr
from xlutils.copy import copy
import xlrd
import time
import openpyxl
import shutil
from xlwt import Borders
import xlwt
from common.init.InitConfig import InitConfig

'''
@excel结果文件
@author: dujianxiao
'''
class Report(InitConfig):
    
    '''
    @生成excel结果文件
    @param path:文件路径
    @param file:用例文件
    @param data:
    @param sheetNames:用例文件中的全部页签名    
    '''
    def createReport(self,reportDate,path,file,data,sheetNames):
        try:
            sheet1=[]
            fileSrc = str(path).replace('/', '\\')+'\\'
            fileRes = fileSrc+'result\\'+file[:-4]+'-'+str(reportDate)+'-report.xls'
            if file.endswith('xls'):
                shutil.copyfile(fileSrc+file,fileRes)
                book = copy(data)
                data = xlrd.open_workbook(fileRes,formatting_info=True)
                book = copy(data)
                [sheet1.append(book.get_sheet(item)) for item in sheetNames]
            elif file.endswith('xlsx'):
                fileRes = fileRes+'x'
                shutil.copyfile(fileSrc+file,fileRes)
                data = openpyxl.load_workbook(fileRes)
                book = data
                [sheet1.append(book.get_sheet_by_name(item)) for item in sheetNames]
            return book,sheet1,fileRes
        except Exception as e:
            print(e)
            fileCheck='文件：'+fileRes+' 正在被其他程序使用'
            print(fileCheck)
            self.console.append("<font color=\"#000000\"></font> ")
            self.console.append("<font color=\"#FF0000\">"+str(fileCheck)+"</font> ")
            self.console.append("<font color=\"#000000\"></font> ")
            