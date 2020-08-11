import xlrd
from common.init.InitConfig import InitConfig
from common.utils.Util import readExcel

'''
@author: dujianxiao
'''
class InitFile(InitConfig):
    def getData(self,path,file):
        data=''
        try:
            data = readExcel(path+'/'+file)
            return data
        except Exception as e:
            print(e)
            self.console.append("<font color=\"#FF0000\">"+str(e)+"</font> ")
            self.console.append("<font color=\"#000000\"></font> ")
            return data
            
    
    