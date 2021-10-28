from time import sleep
from common.utils.Util import *
import requests
from requests.adapters import HTTPAdapter
from common.init.Init import Init
from common.excel.Array import *
import re
import chardet
import os
import json
import datetime
import time
import demjson
import xmltodict
from common.utils.Log import *
from common.utils.analy import analy
from jsonschema import validate

'''
@author: dujianxiao
'''
class Http(Init,analy):
    res=requests.session()
    '''
    @接口请求超时时间目前配的是30S,是否还有必要失败重试？
    '''
#     res.mount('http://',HTTPAdapter(max_retries=3))#重试3次
#     res.mount('https://',HTTPAdapter(max_retries=3))
    userVar=[]#接口变量数组
    userVarValue=[]#接口变量值数组
    headerManager=''
    
    '''
    @http请求
    @param file:用例文件
    @param sheetName:页签名 
    @param userParams:用户变量
    @param userParamsValue:用户变量值
    @param sheet:页签
    @param row:行号
    @param conn:数据库连接对象
    @param column:InitColumn--column
    '''
    def httpRequest(self,file,sheetName,userParams,userParamsValue,sheet,row,conn,column):
        body1=''
        interface=getArray(file,sheet,row,column[19],column[20])#接口变量数组
        url,method,pay,files,header=[repAll(str(getValue(file,sheet,row,column[i])),\
                                            file,sheet,row,conn,column,userParams,userParamsValue,\
                                            self.userVar,self.userVarValue) for i in range(5)]
        '''
        @如果请求头为空且信息头管理器不为空，则使用信息头管理器覆盖
        '''
        if header=='' and self.headerManager!='':
            header=self.headerManager
        getToLog("url："+url)
        getToLog("method："+method)
        getToLog("params："+pay)
        getToLog("files："+files)
        getToLog("headers："+header)
        
        '''
        @校验参数格式，需要是JSON格式
        '''
        try:
            payload = str(pay).encode('utf-8')
            body = payload.decode('utf-8')
            body1='' if body=='' else json.loads(body)
        except Exception as e:
            print(e)
            msg=['参数异常',column[2]]
            return '','---',msg
        '''
        @校验请求头格式，需要是JSON格式
        '''
        try:
            if header=='':
                self.headerManager
            else:
                json.loads(header)
        except Exception as e:
            msg=['请求头异常',column[4]]
            return '','---',msg
        '''
        @暂时只支持四种请求方式
        '''
        if str(method).upper() not in ['GET','POST','PUT','DELETE']:
            msg=['请求方式异常',column[1]]
            return '','---',msg
        r,duration,message=self.sendHttp(url,method,body,body1,header,files,column)
        try:
            '''
            @字符集处理
            '''
            encoding = chardet.detect(r.content).get('encoding')
            '''
            @接口请求异常返回的可能不是response,为了不影响逻辑try一下
            '''
            if '8859' in str(encoding):
                r.encoding='utf-8'
            elif '2312' in str(encoding) or 'gbk' in str(encoding).lower() or 'gb18130' in str(encoding).lower():
                r.encoding='gbk'
            else:
                r.encoding='utf-8'
        except Exception as e:
            print(e)
            
        if '异常' in str(message):
            return r,duration,message
        else:
            getToLog(str(r))
            try:
                '''
                @某些接口返回值是html格式,会出现大量的转义字符,使用loads进行反序列化
                '''
                getToLog("接口响应:"+str(json.loads(r.text)))
            except:
                '''
                @又由于有些接口返回值不是json格式,不能loads,所以如果反序列化失败即不再进行反序列化
                '''
                getToLog("接口响应:"+str(r.text))
            getToLog("响应头："+str(r.headers))
            msg3=Http.analyJSON(self,file,sheet,row,r,column,userParams,userParamsValue)
            '''
            @校验字段或接口变量字段等JSON字段错误
            '''
            if 'json异常' in msg3:
                return r,duration,msg3
            else:
                message=msg3
                
            self.setParams(r,interface,file,sheet,row,column)
            message=self.setHeaderParams(file, sheet, row, column, conn, userParams, userParamsValue)
            if len(message)>1:
                return r,duration,message
        expressionMsg=self.validateExp(r,file,sheet,row,column,conn,userParams,userParamsValue)
        if len(expressionMsg)>1:
            return r,duration,expressionMsg
        return r,duration,''
    
    '''
    @设置信息头管理器
    @供后续接口隐式调用，如果接口设置了信息头，则信息头管理器在此接口中无效
    '''
    def setHeaderParams(self,file,sheet,row,column,conn,userParams,userParamsValue):
        message=['信息头管理器异常']
        headerM=getValue(file,sheet,row,column[21])
        try:
            if headerM != '':
                json.loads(headerM)
                self.headerManager=repAll(str(headerM),file,sheet,row,conn,column,userParams,userParamsValue,self.userVar,self.userVarValue)  
        except Exception as e:
            print(e)
            message.append(column[21])
        return message
    
    '''
    @接口请求成功后存接口变量
    @如果已经存在同名的变量则覆盖，否则新建一个变量
    '''
    def setParams(self,r,interface,file,sheet,row,column):  
        js=self.getResType(r)
        num=len(interface)
        for i in range(len(interface)):
            vue=getValue(file,sheet,row,column[19]+i+num)
            if interface[i]!='':
                if vue not in self.userVar:
                    self.userVarValue.append(eval(js+interface[i]))
                    self.userVar.append(vue)
                else:
                    for j in range(len(self.userVar)):
                        if vue==self.userVar[j]:
                            self.userVarValue[j]=eval(js+interface[i])
    
    '''
    @表达式异常校验
    '''
    def validateExp(self,r,file,sheet,row,column,conn,userParams,userParamsValue):
        expressionMsg=['表达式异常']
        column1 = column[25]
        js = self.getResType(r)
        expression = getArray(file,sheet,row,column[25],column[14])
        for i in range(len(expression)):
            expression[i]=repAll(str(expression[i]),file,sheet,row,conn,column,userParams,userParamsValue,self.userVar,self.userVarValue)
            '''
            @表达式可能涉及到接口响应
            '''
            expression[i] = str(expression[i]).replace("r.json()", js)
            if expression[i] != '':
                try:
                    eval(expression[i])
                except Exception as e:
                    print(e)
                    getError(str(expression[i])+":"+str(e))
                    expressionMsg.append(str(column1))
            column1 = column1+1
        return expressionMsg
    
    '''
    @解析接口响应
    @param file:用例文件
    @param sheet:  
    @param row:行号 
    @param r: 接口请求返回对象
    @param column:列号 
    @param userParams:用户变量
    @param userParamsValue:用户变量值
    '''                      
    def analyJSON(self,file,sheet,row,r,column,userParams,userParamsValue):
        '''
        @取出用例中所能的JSON字段
        '''
        check=getArray(file,sheet,row,column[5],column[7])+getArray(file,sheet,row,column[19],column[20])
        msg = ['json异常']
        res = []
        col=column[5]
        js=self.getResType(r)
        for item in check:
            if(item == ''):
                res.append('')
            else:
                try:
                    item = repVar(str(item),userParams,userParamsValue)
                    res.append(eval(js + item))#eval("r.json()item")
                except Exception as e:
                    print(e)
                    msg.append(str(col))
            if col<column[7]-1:
                col=col+1
            elif col==column[7]-1:
                col=col+(column[19]-column[7])+1
        '''
        @返回异常信息或JSON值数组
        '''
        return msg if len(msg)>1 else res
        
        
    '''
    @return 响应类型,json.jsonp,xml
    '''
    def getResType(self,r):
        js=''
        if str(r)=='':
            return ['']
        else:
            js1='demjson.decode(r.text)'
            js2='json.dumps(xmltodict.parse(r.text))'
            js3='json.loads(re.search("^[^(]*?\((.*)\)[^)]*$",r.text,re.S).group(1))'
        try:
            '''       
            @返回类型为json
            '''
            eval(js1)
            js=js1
        except Exception:
            '''
            @返回类型为xml
            '''
            try:
                js=eval(js2)
            except:
                '''
                @返回类型为jsonp
                '''
                try:
                    eval(js3)
                    js=js3
                except:
                    js=str(r.text)
        return js
    
    '''
    @param url: 
    @param method1:请求方式 
    @param body: 参数
    @param body1: 字典化参数
    @param header1: 请求头
    @param files:上传文件
    @param column:列号  
    @return: r,响应时间,异常信息
    '''
    def sendHttp(self,url,method1,body,body1,header1,files,column):   
        r=''
        body=body.replace('&', '%26').replace('=','%3D')#对&和=进行转义
        post=''
        delete=''
        put=''
        GET="self.get(url,body,body1,header1)"
        POST="self.post(url,body,body1,header1,files)"
        DELETE="self.delete(url,body,body1,header1)"
        PUT="self.put(url,body,body1,header1,files)"
        arrMethod=['GET','POST','PUT','DELETE']
        arr=[str(method1).upper()]+filterArr(arrMethod,str(method1).upper())#把传入的method放到第一个，提高效率
        msg = ['请求方式异常', column[1]]
        try:
            r1,duration=eval(eval(arr[0]))
        except Exception as e:
            getError(str(e))
            if 'Invalid URL' in str(e):
                msg=['url异常',column[0]]
                return str(e),'---',msg
            elif 'Failed to parse: '+str(url)==str(e):
                msg=['url异常',column[0]]
                return str(e),'---',msg
            else:
                msg=['接口请求异常',column[0],column[1],column[2],column[3],column[4],]
                return str(e),'---',msg
        '''
        @如果请求失败，则调用其他请求方式，其他方式有200说明请求方式写错了
        '''
        if '200' in str(r1):
            return r1,duration,['']
        elif '200' in str(eval(eval(arr[1]))):
            return r1,duration,msg
        elif '200' in str(eval(eval(arr[2]))):
            return r1,duration,msg
        elif '200' in str(eval(eval(arr[3]))):
            return r1,duration,msg
        else:
            return r1,duration,[''] 
    
    '''
    @get请求
    @param url: 
    @param body: 参数
    @param body1: 字典化参数
    @param header1: 请求头
    @return: r,响应时间
    '''        
    def get(self,url,body,body1,header1):
        if header1=='':
            if body=='':
                r = self.res.get(url,timeout=30)
            else:
                r = self.res.get(url,params=body1,timeout=30)
        else:
            if body=='':
                r = self.res.get(url, headers=eval(header1),timeout=30)
            else:
                r = self.res.get(url, params=body1, headers=eval(header1),timeout=30)
        du=str(r.elapsed.total_seconds())
        return r,du[:-3]
    
    '''
    @post请求
    @param url: 
    @param body: 参数
    @param body1: 字典化参数
    @param header1: 请求头
    @return: r,响应时间
    ''' 
    def post(self,url,body,body1,header1,files):
        if body1 !='' and 'application/json' in str(header1).lower():
            body1=json.dumps(body1)
        if files=='':
            if header1=='':
                if body=='':
                    r = self.res.post(url,timeout=30)
                else:
                    r = self.res.post(url,data=body1,timeout=30)
            else:
                if body=='':
                    r = self.res.post(url, headers=eval(header1),timeout=30)
                else:
                    r = self.res.post(url, data=body1,headers=eval(header1),timeout=30)
        else:
            if header1=='':
                if body=='':
                    r = self.res.post(url, files=eval(files),timeout=30)
                else:
                    r = self.res.post(url, data=body1,files=files,timeout=30)
            else:
                if body=='':
                    r = self.res.post(url, headers=eval(header1),files=eval(files),timeout=30)
                else:
                    r = self.res.post(url, data=body1,headers=eval(header1),files=files,timeout=30)
        du=str(r.elapsed.total_seconds())
        return r,du[:-3]    
    
    '''
    @delete请求
    @param url: 
    @param body: 参数
    @param body1: 字典化参数
    @param header1: 请求头
    @return: r,响应时间
    '''            
    def delete(self,url,body,body1,header1):
        if header1=='':
            if body=='':
                r = self.res.delete(url,timeout=30)
            else:
                '''
                @不确定这里是json还是data
                '''
                r = self.res.delete(url,json=body1,timeout=30)
        else:
            if body=='':
                r = self.res.delete(url,headers=eval(header1),timeout=30)
            else:
                r = self.res.delete(url,json=body1,headers=eval(header1),timeout=30)
        du=str(r.elapsed.total_seconds())
        return r,du[:-3]
    
    '''
    @put请求
    @param url: 
    @param body: 参数
    @param body1: 字典化参数
    @param header1: 请求头
    @return: r,响应时间
    '''            
    def put(self,url,body,body1,header1,files):
        if body1 !='' and 'application/json' in str(header1).lower():
            body1=json.dumps(body1)
        if files=='':
            if header1=='':
                if body=='':
                    r = self.res.put(url,timeout=30)
                else:
                    r = self.res.put(url,data=body1,timeout=30)
            else:
                if body=='':
                    r = self.res.put(url,headers=eval(header1),timeout=30)
                else:
                    r = self.res.put(url,data=body1,headers=eval(header1),timeout=30)
        else:
            if header1=='':
                if body=='':
                    r = self.res.put(url,files=eval(files),timeout=30)
                else:
                    r = self.res.put(url,data=body1,files=files,timeout=30)
            else:
                if body=='':
                    r = self.res.put(url,headers=eval(header1),files=eval(files),timeout=30)
                else:
                    r = self.res.put(url,data=body1,headers=eval(header1),files=files,timeout=30)
        du=str(r.elapsed.total_seconds())
        return r,du[:-3]
    
    
