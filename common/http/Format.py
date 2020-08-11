from common.http.Http import Http
from common.excel.Array import *
from common.utils.Util import *
from common.utils.analy import analy
import json
import re
import demjson
import os
import chardet
import time

'''
@author: dujianxiao
'''

class Format(Http,analy):
    '''
    @合法性校验
    @param file: 用例文件
    @param sheetName: 页签名
    @param userParams: 用户变量
    @param userParamsValue: 用户变量值
    @param sheet: 
    @param row:行号
    @param conn:数据库连接对象
    @param column: 
    @return: 返回3个值，分别为：http响应、响应时间、异常信息
    '''
    def checkFormat(self,file,sheetName,userParams,userParamsValue,sheet,row,conn,column):
        '''
        @数据库Ip、用户名、密码错误等引起的异常
        '''
        if '数据库异常' in str(conn):
            return '','---',conn
        '''
        @有SQL无数据库连接引起的异常
        '''
        DBMsg=DBExists(file,sheet,row,column,conn)
        if DBMsg!=[]:
            return '','---',DBMsg
        '''
        @查询语句错误引起的异常
        '''
        DBMsg=sqlExcept(file,sheet,row,column,userParams,userParamsValue,self.userVar,self.userVarValue,conn)
        if '数据库异常' in str(DBMsg):
            return '','---',DBMsg
        initMsg=init(file,sheet,row,conn,column,userParams,userParamsValue,self.userVar,self.userVarValue)
        '''
        @数据库初始化语句异常
        '''
        if initMsg != []:
            return '','---',initMsg
        r,duration,msg=self.httpRequest(file,sheetName,userParams,userParamsValue,sheet,row,conn,column)
        '''
        @不直接返回，列出所有可能的异常
        '''
        if '参数异常' in msg:
            return r,duration,msg
        elif '请求头异常' in msg:
            return r,duration,msg
        elif 'url异常' in msg:
            return r,duration,msg
        elif '请求方式异常' in msg:
            return r,duration,msg
        elif '接口请求异常' in msg:
            return r,duration,msg
        elif 'json异常' in msg:
            return r,duration,msg
        elif '信息头管理器异常' in msg:
            return r,duration,msg
        else:
            '''
            @表达式校验放在Http类中更好一点
            '''
            expressionMsg=['表达式异常']
            column1 = column[25]
            js = self.getResType(r)
            expression = getArray(file,sheet,row,column[25],column[14])
            for i in range(0,len(expression)):
                expression[i]=repAll(str(expression[i]),file,sheet,row,conn,column,userParams,userParamsValue,self.userVar,self.userVarValue)
                '''
                @表达式可能涉及到接口响应
                '''
                expression[i] = str(expression[i]).replace("r.json()", js)
                if expression[i] != '':
                    try:
                        eval(expression[i])
                    except Exception as e:
                        getError(str(expression[i])+":"+str(e))
                        expressionMsg.append(str(column1))
                column1 = column1+1
            if len(expressionMsg)>1:
                return r,duration,expressionMsg
            else:        
                return r,duration,msg
    
    '''
    @解析JSON
    @param file: 用例文件
    @param sheetName: 页签名
    @param userParams: 用户变量
    @param userParamsValue: 用户变量值
    @param sheet: 
    @param row:行号
    @param conn:数据库连接对象
    @param column: 
    @return: JSON对象中全部字段的路径和值
    '''    
    def jsonFormat(self,file,sheetName,userParams,userParamsValue,sheet,row,conn,column):
        try:
            init(file,sheet,row,conn,column,userParams,userParamsValue,self.userVar,self.userVarValue)
            r,duration,msg=self.httpRequest(file,sheetName,userParams,userParamsValue,sheet,row,conn,column)
            restore(file,sheet,row,conn,column,userParams,userParamsValue,self.userVar,self.userVarValue)
            '''
            @处理字符集
            '''
            encoding = chardet.detect(r.content).get('encoding')
            if '8859' in str(encoding):
                r.encoding='utf-8'
            elif '2312' in str(encoding) or 'gbk' in str(encoding).lower() or 'gb18130' in str(encoding).lower():
                r.encoding='gbk'
            else:
                r.encoding='utf-8'
            s1,s2 = self.analy(eval(self.getResType(r)))
            return s1,s2
        except Exception as e:
            getError(msg)
            return e,'解析失败'
        
            
        