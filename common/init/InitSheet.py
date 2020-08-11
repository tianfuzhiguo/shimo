from common.init.InitFile import InitFile
from common.init.InitConfig import InitConfig
import shutil
from common.utils.Util import readExcel,getValue,findStr
import os
'''
Created on 2020年4月24日
@author: dujianxiao
@deprecated: 获取用例文件；新建报告文件
'''
class InitSheet(InitFile,InitConfig):
    def getSheet(self,reportDate,path,sheetName,file,data):
        try:
            data = readExcel(path+'/'+file)
        except Exception as e:
            self.console.append("<font color=\"#FF0000\">"+str(e)+"</font> ")
            self.console.append("<font color=\"#000000\"></font> ")
        try:
            '''
            #生测试报告，历史报告移动到history中
            '''
            if file[-4:]=='.xls':
                sheet = data.sheet_by_name(sheetName)
                nrows = sheet.nrows
                ncols = sheet.ncols
            elif file[-5:]=='.xlsx':
                sheet = data.get_sheet_by_name(sheetName)
                nrows = sheet.max_row
                ncols = sheet.max_column
            try:
                isExists=os.path.exists(path+'/result/history')
                if not isExists:
                    os.makedirs(path+'/result/history') 
                fileList=os.listdir(path+'/result/')
                for i in range(0,len(fileList)):
                    if str(reportDate) not in str(fileList[i]) and str('report') in str(fileList[i]):
                        try:                    
                            shutil.move(path+'/result/'+str(fileList[i]), path+'/result/history')
                        except Exception as e:
                            print(e)
                return sheet,nrows,ncols
            except Exception as e:
                print(e)
                    
        except Exception as e:
            print(e)
