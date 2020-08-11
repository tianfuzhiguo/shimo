import json
import re
import demjson
import time
import os
import xmltodict     
from common.utils.Util import *      
from common.init.Init import *
                                                                                                  
'''   
@获取校验字段和预期结果的原始值和结果值                                                                                                       
@author: dujianxiao          
'''                                                                                                          
class Array(): 
    '''
    @校验字段数组－－原值
    @param file:用例文件
    @param sheet:
    @param row:行号
    @param column:列号    
    '''
    
    def check(self,file,sheet,row,column,userParams,userParamsValue,conn,userVar,userVarValue):
        arr = getArray(file,sheet,row,column[5],column[8])
        return [repAll(str(item),file,sheet,row,conn,column,userParams,userParamsValue,userVar,userVarValue) for item in arr]
    
    '''
    @预期结果数组－－原值
    @param file:用例文件
    @param sheet:
    @param row:行号
    @param column:列号
    '''
    def expResultInit(self,file,sheet,row,column,userParams,userParamsValue,conn,userVar,userVarValue):
        arr = getArray(file,sheet,row,column[8],column[11])
        return [repAll(str(item),file,sheet,row,conn,column,userParams,userParamsValue,userVar,userVarValue) for item in arr]
                                                                                                                
    '''                                                                                                          
    @校验字段结果数组                                                                                            
    @param file:用例文件                                                                                         
    @param sheet:                                                                                                
    @param row:行号                                                                                              
    @param conn:数据库连接对象                                                                                   
    @param jsonValue:json数组                                                                                    
    @param regValue:正则数组                                                                                     
    @param column:列号                                                                                           
    '''                                                                                                          
    def checkRes(self,r,file,sheet,row,conn,column,userParams,userParamsValue,userVar,userVarValue):  
        jss=self.getResType(r)                                                                                        
        '''                                                                                                      
        @固定值数组                                                                                              
        '''      
        js=getArray(file,sheet,row,column[5],column[7])                                                          
        jsonValue=[]                                                                                              
        for item in js:
            jsonValue.append('' if item=='' else eval(jss+item))                                                            
        '''                                                                                                      
        @SQL数组                                                                                                 
        '''    
        sqlArr = getSqlResultArray(file,sheet,row,conn,column[7],column[8])
        arr=jsonValue + sqlArr                                                                                   
        arr=[repAll(item,file,sheet,row,conn,column,userParams,userParamsValue,userVar,userVarValue) for item in arr]
        getToLog('校验字段：'+str(arr))     
        return arr                                                                                               
                                                                                                                 
    '''                                                                                                          
    @预期结果值数组                                                                                              
    @param file:用例文件                                                                                         
    @param sheet:                                                                                                
    @param row:行号                                                                                              
    @param conn:数据库连接对象                                                                                   
    @param column:列号                                                                                           
    '''                                                                                                          
    def expResult(self,file,sheet,row,conn,column,userParams,userParamsValue,userVar,userVarValue):                      
        arr1 = getArray(file,sheet,row,column[8],column[9])                                                                                                                                                              
        arr2 = getSqlResultArray(file,sheet,row,conn,column[9],column[10])                                                                                                                                    
        arr3 = getArray(file,sheet,row,column[10],column[11])                                                                                                                                                                  
        arr = arr1 + arr2 + arr3                                                                                                                                                                                        
        arr=[repAll(item,file,sheet,row,conn,column,userParams,userParamsValue,userVar,userVarValue) for item in arr]  
        getToLog('预期结果：'+str(arr))   
        return arr                                                                                               