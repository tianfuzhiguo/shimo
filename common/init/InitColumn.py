from common.init.InitSheet import InitSheet
from common.utils.Util import findStr
'''
Created on 2020年4月24日
@author: dujianxiao
'''

class InitColumn(InitSheet):
        
    '''
    @获取sheet页各标志位的列号
    @param file:用例文件
    @param sheet:
    @param ncols:列数   
    '''    
    def getColumn(self,file,sheet,ncols): 
        column=[]   
        nameCol=findStr(file,sheet,ncols,'name')
        urlCol=findStr(file,sheet,ncols,'url')
        methodCol=findStr(file,sheet,ncols,'method')
        paramCol=findStr(file,sheet,ncols,'param')
        fileCol=findStr(file,sheet,ncols,'file')
        headerCol=findStr(file,sheet,ncols,'header')
        part101Col=findStr(file,sheet,ncols,'part101')
        part201Col=findStr(file,sheet,ncols,'part201')
        part301Col=findStr(file,sheet,ncols,'part301')
        section101Col=findStr(file,sheet,ncols,'section101')
        section201Col=findStr(file,sheet,ncols,'section201')
        section301Col=findStr(file,sheet,ncols,'section301')
        resTextCol=findStr(file,sheet,ncols,'resText')
        resHeaderCol=findStr(file,sheet,ncols,'resHeader')
        statusCodeCol=findStr(file,sheet,ncols,'statusCode')
        expressionCol=findStr(file,sheet,ncols,'expression')
        statusCol=findStr(file,sheet,ncols,'status')
        timeCol=findStr(file,sheet,ncols,'time')
        init001Col=findStr(file,sheet,ncols,'init001')
        restore001Col=findStr(file,sheet,ncols,'restore001')
        dyparam001Col=findStr(file,sheet,ncols,'dyparam001')
        key001Col=findStr(file,sheet,ncols,'key001')
        value001Col=findStr(file,sheet,ncols,'value001')
        headerManagerCol=findStr(file,sheet,ncols,'headerManager')
        DBCol=findStr(file,sheet,ncols,'数据库')
        IterationCol=findStr(file,sheet,ncols,'Iteration')
        
        column.append(urlCol)
        column.append(methodCol)
        column.append(paramCol)
        column.append(fileCol)
        column.append(headerCol)
        column.append(part101Col)
        column.append(part201Col)
        column.append(part301Col)
        column.append(section101Col)
        column.append(section201Col)
        column.append(section301Col)
        column.append(resTextCol)
        column.append(resHeaderCol)
        column.append(statusCodeCol)
        column.append(statusCol)
        column.append(timeCol)
        column.append(init001Col)
        column.append(restore001Col)
        column.append(dyparam001Col)
        column.append(key001Col)
        column.append(value001Col)
        column.append(headerManagerCol)
        column.append(DBCol)
        column.append(nameCol)
        column.append(IterationCol)
        column.append(expressionCol)
        return column
    