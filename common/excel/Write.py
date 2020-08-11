from common.excel.Array import *
from common.excel.Array import Array
from common.excel.Report import Report
from common.http.Format import Format
from common.excel.Template import Template
import json
import re
import demjson
import time
import os
import xmltodict  
from openpyxl.styles import PatternFill
'''
@有些看起来没有用到的库是为了表达式准备的－－其他类也一样
'''


'''
@author: dujianxiao
'''
class Write(Array,Format,Template,Report,Init):
        
    '''
    @deprecated: 写入接口请求结果
    @param file:用例文件
    @param model:模式(普通,简洁)
    @param row:行号
    @param sheetName:页签名
    @param userParams:用户变量数组
    @param userParamsValue:用户变量值数组
    @param sheet:
    @param book:
    @param sheet1:
    @param fileRes:用例结果文件
    @param column:列号
    @param itera:第itera次迭代，从0计数 
    '''
    def write(self,file,model,row,sheetName,userParams,userParamsValue,sheet,book,sheet1,fileRes,column,currentItera,Iteration):
        resp=[]
        skipDict=[]
        resultDict=[]
        status='成功'
        dict={}
        DBExc=[]
        iteraValue=getValue(fileRes,sheet,row,column[24])
        url=str(getValue(file,sheet,row,column[0]))
        url=repRel(row,self.userVar,self.userVarValue,url)
        url=repVar(str(url),userParams,userParamsValue)
        className=str(getValue(file,sheet,row,column[23]))
        className=repRel(row,self.userVar,self.userVarValue,className)
        className=repVar(str(className),userParams,userParamsValue)
        if isinstance(iteraValue, int) == False and iteraValue != '':
            skipDict=self.setSkip(sheet,row,book,sheet1,fileRes,'迭代异常',column,currentItera,Iteration,userParams,userParamsValue,'')
            print('迭代次数异常')
            status='异常'
            duration='--'
            resultDict=[]
            DBExc=[]
        else:
            conn=getConn(file,sheet,row,column,userParams,userParamsValue)
            r,duration,msg=self.checkFormat(file,sheetName,userParams,userParamsValue,sheet,row,conn,column)
            print(r,duration,msg)
            try:
                resp.append(str(r.text))
                '''
                @普通模式下打印接口响应
                @如果接口返回中含有html元素,使用append方式显示的是渲染后的结果,使用insertPlainText显示的是html原文
                @使用insertPlainText性能会很差,很容易导致页面卡死进而导致程序崩溃
                '''
                if model=='普通':
                    form,ss = self.getType(r)
                    num=len(ss)//1000+1
                    for i in range(0,num):
                        if form=='xml':
                            self.console.append("<font color=\"#000000\"></font>")
                            self.console.insertPlainText(ss[i*1000:(i+1)*1000])
                        
                        elif form=='json' or form=='jsonp':
                            self.console.append("<font color=\"#000000\">"+ss[i*1000:(i+1)*1000]+"</font>")
                        '''
                        @由于qt性能的原因，每1000个字符暂停100毫秒，100毫秒是一个经验值
                        '''
                        time.sleep(0.1)
                    '''
                    @解决普通模式下客户文字错位和文字颜色与预期不问题
                    '''
                    time.sleep(0.1)
            except:
                pass
            if '异常' in str(msg):
                if '数据库异常' not in str(msg):
                    url=rep(file,sheet,row,conn,url,column,userParams,userParamsValue,self.userVar,self.userVarValue)
                    className=rep(file,sheet,row,conn,className,column,userParams,userParamsValue,self.userVar,self.userVarValue) 
                skipDict=self.setSkip(sheet,row,book,sheet1,fileRes,msg,column,currentItera,Iteration,userParams,userParamsValue,conn)
                status='异常'
            else:
                url=rep(file,sheet,row,conn,url,column,userParams,userParamsValue,self.userVar,self.userVarValue)
                '''
                @取校验数据、预期结果的原始值和结果值
                '''
                checkRes1=self.checkRes(r,file,sheet,row,conn,column,userParams,userParamsValue,self.userVar,self.userVarValue)
                resInit=self.expResultInit(file,sheet,row,column,userParams,userParamsValue,conn,self.userVar,self.userVarValue)
                check1=self.check(file,sheet,row,column,userParams,userParamsValue,conn,self.userVar,self.userVarValue)
                result1=self.expResult(file,sheet,row,conn,column,userParams,userParamsValue,self.userVar,self.userVarValue)
                '''
                @在数据恢复之前进行三者替换
                '''
                statusCode=getArray(file,sheet,row,column[13],column[25])
                resHeader=getArray(file,sheet,row,column[12],column[13])
                res=getArray(file,sheet,row,column[11],column[12])
                expression = getArray(file,sheet,row,column[25],column[14])
                [repAll(str(item),file,sheet,row,conn,column,userParams,userParamsValue,self.userVar,self.userVarValue) for item in statusCode]
                [repAll(str(item),file,sheet,row,conn,column,userParams,userParamsValue,self.userVar,self.userVarValue) for item in resHeader]
                [repAll(str(item),file,sheet,row,conn,column,userParams,userParamsValue,self.userVar,self.userVarValue) for item in res]
                [repAll(str(item),file,sheet,row,conn,column,userParams,userParamsValue,self.userVar,self.userVarValue) for item in expression]
                '''
                @数据恢复之前把所有数据库相关的操作处理完
                '''    
                resMsg=restore(file,sheet,row,conn,column,userParams,userParamsValue,self.userVar,self.userVarValue)
                '''
                @数据库恢复部分的SQL异常
                '''
                if resMsg != []:
                    skipDict=self.setSkip(sheet,row,book,sheet1,fileRes,resMsg,column,currentItera,Iteration,userParams,userParamsValue,conn)
                    status='异常'
                else:
                    '''
                     @完全没有异常了再执行setResult
                    '''
                    resultDict=self.setResult(file,row,book,sheet,sheet1,fileRes,checkRes1,check1,result1,resInit,r,duration,column,userParams,userParamsValue,self.userVar,self.userVarValue,res,resHeader,statusCode,expression,currentItera,Iteration)
                    if len(resultDict)>0:
                        status='失败'
            try:
                if conn==[[]]:
                    pass
                else:
                    if '数据库异常' not in str(conn):
                        conn[0][0].close()
            except Exception as e:
                print(e)
        '''
        @只统计最后一次的结果
        '''
        if currentItera==Iteration-1:    
            '''    
            @信息存入字典，用于html测试报告
            '''
            dict['className']=className
            dict['url']=url
            dict['method']=getValue(file,sheet,row,column[1])
            dict['param']=getValue(file,sheet,row,column[2])
            dict['header']=getValue(file,sheet,row,column[4])
            dict['duration']=duration
            dict['resp']=resp
            dict['status']=status
            dict['log']=skipDict+resultDict+DBExc        
            return dict
        else:
            return []

    '''
    @获取接口响应类型：xml,json,jsonp
    '''
    def getType(self,r):
        form = ''
        ss = ''
        try:
            try:
                '''
                @某些接口返回值是html格式,会出现大量的转义字符,使用loads进行反序列化
                '''
                ss=str(json.loads(r.text))
                form='json'
            except Exception:
                try:
                    '''
                    @又由于有些接口返回值不是json格式,不能loads,所以如果反序列化失败即不再进行反序列化
                    '''
                    ss=str(r.text)
                    form='xml'
                except Exception:
                    try:
                        '''
                        @格式为jsonp
                        '''
                        ss=str(json.loads(re.match(".*?({.*}).*",r.text,re.S).group(1)))
                        form='jsonp'
                    except Exception:
                        pass
        except:
            if str(r)!='':
                self.console.append("<font color=\"#000000\">"+str(r)+"</font>")
        return form,ss
        
    '''
    @deprecated: 解析JSON
    @param file:用例文件
    @param row:行号
    @param sheetName:页签名
    @param userParams:用户变量数组
    @param userParamsValue:用户变量值数组
    @param sheet:
    @param fileRes:用例结果文件
    @param column:列号
    '''
    def analyFunc(self,file,row,sheetName,userParams,userParamsValue,sheet,fileRes,column):
        '''
        @JSON解析中不对异常情况进行处理，如有异常直接解析失败
        '''
        getToLog("☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆【"+str(sheetName)+"】第"+str(row+1)+"个接口解析开始☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆")
        self.console.append("<font color=green>"+str(row+1)+' '+str(getValue(file,sheet,row,column[23]))+"</font>")
        try:
            conn=getConn(file,sheet,row,column,userParams,userParamsValue)
            s1,s2=self.jsonFormat(file,sheetName,userParams,userParamsValue,sheet,row,conn,column)
            if s2=='解析失败':
                self.console.append("<font color=\"#FF0000\">"+'解析失败.'+"</font>")
            else:
                for i in range(0,len(s1)):
                    self.console.append("<font color=\"#000000\">"+str(s1[i])+':'+str(s2[i])+"</font>")
                    time.sleep(0.01)
        except Exception as e:
            print(e)  
            self.console.append("<font color=\"#FF0000\">"+'解析失败.'+"</font>")
        getToLog("☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆【"+str(sheetName)+"】第"+str(row+1)+"个接口解析结束☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆")
        
        
    '''
    @deprecated: 如果数据合法性校验不通过则调用此方法
    @param file:用例文件
    @param sheet:
    @param row:行号
    @param book:
    @param sheet1:
    @param fileRes:用例结果文件
    @param msg: 接口返回的异常信息
    @param column:列号
    '''
    def setSkip(self,sheet,row,book,sheet1,fileRes,msg,column,currentItera,Iteration,userParams,userParamsValue,conn):
        arr=[]
        skipDict=[]
        blue=setStyle(7)
        if '迭代异常' in str(msg):
            self.console.append("<font color=\"#FF0000\">迭代次数只能为空或非负整数</font>")
            self.status3=self.status3+1
            iteraValue=getValue(fileRes,sheet,row,column[24])
            skipDict.append("迭代次数异常:"+str(iteraValue))
            '''
            @标识结果为：skip，并设背景为蓝色
            '''
            if fileRes[-4:]=='.xls':
                sheet1.write(row,column[24],iteraValue,blue)
                sheet1.write(row,column[14],'skip',blue)
            elif fileRes[-5:]=='.xlsx':
                self.setValueColor(sheet1,row+1,column[24],iteraValue,"blue")
                self.setValueColor(sheet1,row+1,column[14],'skip',"blue") 
        else:
            if '数据库异常' in str(msg):
                if msg[0][1]==column[22] and msg[1]==[]:
                    '''
                    @有sql未选择数据库
                    '''
                    self.console.append("<font color=\"#FF0000\">"+str(msg[0])+"</font>")
                    getError(str(msg[0]))
                elif msg[0][1]==column[22] and msg[1]!=[]:
                    '''
                    @数据库连接类的异常
                    '''
                    err=msg[1][0]
                    self.console.append("<font color=\"#FF0000\">"+err+"</font>")
                    self.console.append("<font color=\"#FF0000\">"+str(msg[0])+"</font>")
                else:
                    '''
                    @sql执行异常
                    '''
                    for i in range(1,len(msg[0])):
                        exceValue=getValue(fileRes,sheet,row,int(msg[0][i]))
                        self.console.append("<font color=\'#FF0000\'>"+str(msg[1][i-1])+"</font>")
                    self.console.append("<font color=\"#FF0000\">"+str(msg[0])+"</font>")
                '''
                @异常信息存skipDict用于html测试报告
                '''
                skipDict.append(str(msg[0]))
                if msg[1]!=[]:
                    skipDict.append(str(msg[1]))
            else:    
                
                skipDict.append(str(msg))
                getToLog(str(msg))
                for i in range(1,len(msg)):
                    exceValue=getValue(fileRes,sheet,row,int(msg[i]))
                    exceValue=repAll(str(exceValue),fileRes,sheet,row,conn,column,userParams,userParamsValue,self.userVar,self.userVarValue)
                    self.console.append("<font color=\'#FF0000\'>"+exceValue+"</font>")
                    skipDict.append(exceValue)
                    getToLog(exceValue)
                self.console.append("<font color=\"#FF0000\">"+str(msg)+"</font>")
            if currentItera==Iteration-1:
                self.status3=self.status3+1
            '''
            @标识结果为：skip，并设背景为蓝色
            ''' 
            if fileRes[-4:]=='.xls':
                sheet1.write(row,column[14],'skip',blue)
            elif fileRes[-5:]=='.xlsx':
                self.setValueColor(sheet1,row+1,column[14],'skip',"blue") 
            
            '''
            @去掉异常信息的数组
            ''' 
            if '数据库异常' in str(msg):
                newArr=filterArr(msg[0],'异常')
            else:
                newArr=filterArr(msg,'异常')
            '''
            @标识合法性校验不通过的单元格为蓝色
            '''
            for item in newArr:
                if fileRes[-4:]=='.xls':
                    sheet1.write(row,int(item),getValue(fileRes,sheet,row,int(item)),blue)
                elif fileRes[-5:]=='.xlsx':
                    self.setValueColor(sheet1,row+1,int(item),getValue(fileRes,sheet,row,int(item)),"blue")
        book.save(fileRes)
        return skipDict
        
    '''
    @deprecated: 数据合法性校验通过后调用此方法，校验各字段的值是否正确
    @param file:用例文件
    @param row:行号
    @param book:
    @param sheet:
    @param sheet1:
    @param fileRes:用例结果文件    
    @param checkRes1:校验字段结果数组+文件数组
    @param check1: 校验字段数组－－原值
    @param result1: 预期结果值数组
    @param resInit: 预期结果数组－－原值
    @param r:接口响应
    @param duration:接口响应时间
    @param column:列号    
    '''
    def setResult(self,file,row,book,sheet,sheet1,fileRes,checkRes1,check1,result1,resInit,r,duration,column,userParams,userParamsValue,userVar,userVarValue,res,resHeader,statusCode,expression,currentItera,Iteration):
        resultDict=[]
        red=setStyle(2)
        green=setStyle(3)
        status=0
        '''
        @写入接口响应时间
        @预置结果为 true
        '''
        if fileRes[-4:]=='.xls':
            sheet1.write(row,column[15],duration)
            sheet1.write(row,column[14],'true',green)
        elif fileRes[-5:]=='.xlsx':
            self.setValueColor(sheet1,row+1,column[15],duration,"")
            self.setValueColor(sheet1,row+1,column[14],'true',"green")
            
        '''
        @校验预期结果，精确匹配
        '''
        for j in range(0,len(check1)):
            if str(checkRes1[j])!=str(result1[j]):#
                if fileRes[-4:]=='.xls':
                    sheet1.write(row,column[5]+j,str(check1[j])+'-->'+str(checkRes1[j])+':'+str(result1[j]),red)
                    sheet1.write(row,column[8]+j,str(resInit[j]),red)
                    sheet1.write(row,column[14],'false',red)
                elif fileRes[-5:]=='.xlsx':
                    self.setValueColor(sheet1,row+1,column[5]+j,str(check1[j])+'-->'+str(checkRes1[j])+':'+str(result1[j]),"red")
                    self.setValueColor(sheet1,row+1,column[8]+j,str(resInit[j]),"red")
                    self.setValueColor(sheet1,row+1,column[14],'false',"red")
                self.console.append("<font color=\"#FF0000\">"+str(check1[j])+':实际结果:'+str(checkRes1[j])+'-->预期结果:'+str(result1[j])+"</font>")
                resultDict.append(str(check1[j])+':实际结果:'+str(checkRes1[j])+'-->预期结果:'+str(result1[j]))
                status=1

        '''
        @响应断言
        '''
        for i in range(0,len(res)):
            if res[i] in r.text:
                pass
            else:
                if fileRes[-4:]=='.xls':
                    sheet1.write(row,column[14],'false',red)
                    sheet1.write(row,column[11]+i,res[i],red)
                elif fileRes[-5:]=='.xlsx':
                    self.setValueColor(sheet1,row+1,column[14],'false',"red")
                    self.setValueColor(sheet1,row+1,column[11]+i,res[i],"red")
                self.console.append("<font color=\"#FF0000\">"+'响应断言失败:'+str(res[i])+"</font>")
                resultDict.append('响应断言失败:'+str(res[i]))
                status=1
                
        '''
        @校验响应头，模糊匹配
        '''
        for i in range(0,len(resHeader)):
            if str(resHeader[i])=='':
                pass
            elif str(resHeader[i]) not in str(r.headers):
                if fileRes[-4:]=='.xls':
                    sheet1.write(row,column[12]+i,str(resHeader[i]),red)
                    sheet1.write(row,column[14],'false',red)
                elif fileRes[-5:]=='.xlsx':
                    self.setValueColor(sheet1,row+1,column[12]+i,str(resHeader[i]),"red")
                    self.setValueColor(sheet1,row+1,column[14],'false',"red")
                self.console.append("<font color=\"#FF0000\">"+'响应头断言失败:'+str(resHeader[i])+"</font>")
                resultDict.append('响应头断言失败:'+str(resHeader[i]))
                status=1
        
            
        '''
        @校验响应码，精确匹配
        '''
        for i in range(0,len(statusCode)):
            if str(statusCode[i])=='':
                pass
            elif str(statusCode[i]) !=str(r.status_code):
                if fileRes[-4:]=='.xls':
                    sheet1.write(row,column[13]+i,str(statusCode[i])+'-->'+str(r.status_code)+':'+str(statusCode[i]),red)
                    sheet1.write(row,column[14],'false',red)
                elif fileRes[-5:]=='.xlsx':
                    self.setValueColor(sheet1,row+1,column[13]+i,str(statusCode[i])+'-->'+str(r.status_code)+':'+str(statusCode[i]),"red")
                    self.setValueColor(sheet1,row+1,column[14],'false',"red")
                self.console.append("<font color=\"#FF0000\">"+'响应码断言失败:'+'实际结果:'+str(r.status_code)+'-->预期结果:'+str(statusCode[i])+"</font>")
                resultDict.append('响应码断言失败:'+'实际结果:'+str(r.status_code)+'-->预期结果:'+str(statusCode[i]))
                status=1
            
        '''
        @校验表达式
        '''
        js = self.getResType(r)
        for i in range(0,len(expression)):
            expreFlag = True
            expression[i] = str(expression[i]).replace("r.json()", js)
            if expression[i]=='':
                expreFlag = True
            else:
                expreFlag = eval(expression[i])
            if expreFlag != False:
                pass
            else:
                if fileRes[-4:]=='.xls':
                    sheet1.write(row,column[25]+i,str(expression[i]),red)
                    sheet1.write(row,column[14],'false',red)
                elif fileRes[-5:]=='.xlsx':
                    self.setValueColor(sheet1,row+1,column[25]+i,str(expression[i]),"red")
                    self.setValueColor(sheet1,row+1,column[14],'false',"red")
                self.console.append("<font color=\"#FF0000\">"+'表达式断言失败:'+str(expression[i])+"</font>")
                resultDict.append('表达式断言失败:'+str(expression[i]))
                status=1
            
        if currentItera==Iteration-1:
            if status==1:
                self.status2=self.status2+1
            else:
                self.status1=self.status1+1
            
        book.save(fileRes)  
        return resultDict
    
    
    '''
    @写入值并设置背景色
    @param sheet1:
    @param row:行号
    @param column:列号
    @param value:写入单元格的值
    @param color:单元格背景色(red,blue,green)
    '''
    def setValueColor(self,sheet1,row,column,value,color):
        sheet1.cell(row = row, column =column , value = value)
        if color=='':
            pass
        else:
            if color=='blue':
                color='00FFFF'
            elif color=='red':
                color='FF0000'
            elif color=='green':
                color='00FF00'
            color_fill = PatternFill("solid", fgColor=color)
            sheet1.cell(row, column).fill = color_fill
    
    ''' 
    @deprecated: 执行－－单行执行或全量执行（无参数）
    @param file:用例文件
    @param model:模式(普通,简洁)
    @param n:行号
    @param sheetName:页签名
    @param userParams:用户变量数组
    @param userParamsValue:用户变量值数组
    @param sheet:
    @param nrows:行数
    @param ncols:列数
    @param book:
    @param sheet1:  
    @param fileRes:用例结果文件    
    @param column:列号 
    @param allRows:全部用例数
    '''   
    def run(self,file,model,n,sheetName,userParams,userParamsValue,sheet,nrows,book,sheet1,fileRes,column,allRows): 
        testResult=[]
        dict={}
        self.console.append("<font color=\"#000000\"></font>")
        if n=='':
            '''
            @全量执行
            '''
            self.console.append("<font size=4 color=blue>"+"【"+sheetName+"】"+"</font>")
            for row in range(3,nrows+1):
                className = str(getValue(file,sheet,row-1,column[23]))
                Iteration=getValue(file,sheet,int(row)-1, column[24])
                if isinstance(Iteration, int):
                    for i in range(0,Iteration):    
                        print(row)
                        self.console.append("<font color=\"#000000\"></font>")
                        self.console.append("<font color=green>"+str(row)+' '+className+"</font>")
                        getToLog("☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆"+"【"+str(sheetName)+"】"+"第"+str(row)+"个接口【"+className+"】请求开始☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆")
                        if i==Iteration-1:
                            testResult.append(self.write(file,model,row-1,sheetName,userParams,userParamsValue,sheet,book,sheet1,fileRes,column,i,Iteration))
                        else:
                            self.write(file,model,row-1,sheetName,userParams,userParamsValue,sheet,book,sheet1,fileRes,column,i,Iteration)
                        self.successNum.setText(str(self.status1))
                        self.failNum.setText(str(self.status2))
                        self.skipNum.setText(str(self.status3))
                        self.result.setText(str(self.status1+self.status2+self.status3)+'/'+str(allRows))
                        getToLog("☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆"+"【"+str(sheetName)+"】"+"第"+str(row)+"个接口【"+className+"】请求结束☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆")
                else:
                    print(row)
                    self.console.append("<font color=\"#000000\"></font>")
                    self.console.append("<font color=green>"+str(row)+' '+className+"</font>")
                    getToLog("☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆"+"【"+str(sheetName)+"】"+"第"+str(row)+"个接口【"+className+"】请求开始☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆")
                    testResult.append(self.write(file,model,row-1,sheetName,userParams,userParamsValue,sheet,book,sheet1,fileRes,column,0,1))
                    self.successNum.setText(str(self.status1))
                    self.failNum.setText(str(self.status2))
                    self.skipNum.setText(str(self.status3))
                    self.result.setText(str(self.status1+self.status2+self.status3)+'/'+str(allRows))
                    getToLog("☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆"+"【"+str(sheetName)+"】"+"第"+str(row)+"个接口【"+className+"】请求结束☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆")
        else:
            '''
            @debug
            '''
            className = str(getValue(file,sheet,n-1,column[23]))
            Iteration=getValue(file,sheet,int(n)-1, column[24])
            if isinstance(Iteration, int):
                for i in range(0,Iteration):
                    print(n)
                    self.console.append("<font color=\"#000000\"></font>")
                    self.console.append("<font color=green>"+str(n)+' '+className+"</font>")
                    getToLog("☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆"+"【"+str(sheetName)+"】"+"第"+str(n)+"个接口【"+className+"】请求开始☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆")
                    if i==Iteration-1:
                        testResult.append(self.write(file,model,n-1,sheetName,userParams,userParamsValue,sheet,book,sheet1,fileRes,column,i,Iteration))
                    else:
                        self.write(file,model,n-1,sheetName,userParams,userParamsValue,sheet,book,sheet1,fileRes,column,i,Iteration)
                    self.successNum.setText(str(self.status1))
                    self.failNum.setText(str(self.status2))
                    self.skipNum.setText(str(self.status3))
                    self.result.setText(str(self.status1+self.status2+self.status3)+'/'+str(allRows))
                    getToLog("☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆"+"【"+str(sheetName)+"】"+"第"+str(n)+"个接口【"+className+"】请求结束☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆")
            else:        
                print(n)
                self.console.append("<font color=\"#000000\"></font>")
                self.console.append("<font color=green>"+str(n)+' '+className+"</font>")
                getToLog("☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆"+"【"+str(sheetName)+"】"+"第"+str(n)+"个接口【"+className+"】请求开始☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆")
                testResult.append(self.write(file,model,n-1,sheetName,userParams,userParamsValue,sheet,book,sheet1,fileRes,column,0,1))
                self.successNum.setText(str(self.status1))
                self.failNum.setText(str(self.status2))
                self.skipNum.setText(str(self.status3))
                self.result.setText(str(self.status1+self.status2+self.status3)+'/'+str(allRows))
                getToLog("☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆"+"【"+str(sheetName)+"】"+"第"+str(n)+"个接口【"+className+"】请求结束☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆")                

        '''
        @用于html测试报告
        '''
        dict['testAll']=self.status1+self.status2+self.status3
        dict['testPass']=self.status1
        dict['testFail']=self.status2
        dict['testSkip']=self.status3
        return dict,testResult
        
        