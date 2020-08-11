import os
from common.init.InitFile import InitFile
from common.init.InitSheet import InitSheet
from common.init.InitColumn import InitColumn
import time

'''
@author: dujianxiao
'''
class Init(InitColumn,InitSheet,InitFile):
    def init(self,reportDate,path,file,sheetName):
        fileData,email,userParams,userParamsValue=self.initConfig(path)
        data=self.getData(path,file)
        sheet,nrows,ncols=self.getSheet(reportDate,path,sheetName,file,data)
        column=self.getColumn(file,sheet,ncols)
        return fileData,email,userParams,userParamsValue,data,sheet,nrows,ncols,column