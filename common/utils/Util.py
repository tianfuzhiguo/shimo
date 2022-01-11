from common.utils.SmLog import SmLog
from common.init.Init import Init
import cx_Oracle,pymysql,pymssql
'''
@author: dujianxiao
'''
class Util(SmLog,Init):
    '''
    @连接数据库 
    @param file:用例文件
    @param sheet:  
    @param row:行号 
    '''
    def getConn(self,file,sheet,row):
        DB=self.getValue(file,sheet,row,self.DBCol)
        DB=self.repVar(str(DB))
        if DB=='':
            return [[]]
        else:
            try: 
                #DB1,DB2,DB3是配置文件中预置的3个数据库
                if DB=="DB1":
                    DB="${DB1}"
                if DB=="DB2":
                    DB="${DB2}"
                if DB=="DB3":
                    DB="${DB3}"
                DB=self.repVar(str(DB))
                conn=eval(DB)
            except Exception as e:
                print(e)
                self.getError(str(e))
                return [['数据库异常',self.DBCol],[str(e)]]
            return [[conn]] 
        
    '''
    @有SQL则数据库不允许为空
    '''
    def DBExists(self,file,sheet,row,conn):
        allSql = self.getArray(file,sheet,row,self.part301Col,self.section101Col) + \
                 self.getArray(file,sheet,row,self.section201Col,self.section301Col) + \
                 self.getArray(file,sheet,row,self.init001Col,self.key001Col)
        '''
        @如果sql不全为空而数据库连接为空，则返回数据库异常，否则返回空数组
        '''
        return [['数据库异常', self.DBCol],[]] if ''.join(allSql)!='' and conn==[[]] else []
    
    '''
    @获取sql结果数组
    @param file:用例文件
    @param sheet:  
    @param row:行号
    @param conn:数据库连接对象
    @param start: 
    @param end:  
    '''
    def getSqlResultArray(self,file,sheet,row,conn,start,end):
        data = []
        sqlArray = self.getArray(file,sheet,row,start,end)
        column=start
        '''
        @在此之前已经校验过sql不全为空而数据库为空的情况
        @所以如果此时数据库为空，说明sql全为空
        '''
        if conn==[[]]:
            return sqlArray
        else:
            cursor = conn[0][0].cursor()
            for item in sqlArray:
                if(item == ''):
                    data.append('')
                else:
                    cursor.execute(item)
                    dd = cursor.fetchone()
                    data.append(None if dd==None else dd[0])
                column=column+1
            cursor.close()
            conn[0][0].commit()
            return data
    
    '''
    @此方法仅用作验证SQL句是否正确
    @param file:用例文件
    @param sheet:  
    @param row:行号 
    @param conn:数据库连接对象
    '''
    def sqlExcept(self,file,sheet,row,conn):
        conn=self.getConn(file,sheet,row)
        msg = ['数据库异常']
        SqlMsg = []
        sqlArray = self.getArray(file,sheet,row,self.part301Col,self.section101Col) + \
                   self.getArray(file,sheet,row,self.section201Col,self.section301Col) + \
                   self.getArray(file,sheet,row,self.dyparam001Col,self.key001Col)
        for item in sqlArray:
            item=self.repVar(str(item))
            item=self.repRel(row,str(item))
        if conn==[[]]:
            return []
        else:   
            column1=self.part301Col 
            cursor = conn[0][0].cursor()
            for item in sqlArray:
                if item != '':
                    try:
                        '''                   
                        #这三部分的SQL只能是查询语句
                        '''
                        if (str(item).lower()).replace(' ', '').startswith('select')!=True:
                            msg.append(column1)
                            SqlMsg.append(column1)
                        else:
                            cursor.execute(item)   
                    except Exception as e:
                        print(e)
                        msg.append(column1)
                        self.getError(str(e))
                        SqlMsg.append(str(e))
    #                     msg.append(str(e))
                if column1==self.section101Col-1:
                    column1=self.section201Col-1
                if column1==self.section301Col-1:
                    column1=self.dyparam001Col-1
                column1=column1+1
            cursor.close()
            conn[0][0].close()
            return [msg,SqlMsg] if len(msg)>1 else []
        
    '''    
    @数据初始化
    @param file:用例文件
    @param sheet:  
    @param row:行号
    @param conn:数据库连接对象 
    '''
    def initData(self,file,sheet,row,conn):
        msg = ['数据库异常']
        SqlMsg = []
        column1=self.init001Col
        sqlArray = self.getArray(file,sheet,row,self.init001Col,self.restore001Col)
        for item in sqlArray:
            item=self.repVar(str(item))
            item=self.repRel(row,str(item))
        if conn ==[[]]:
            return []
        else:
            cursor = conn[0][0].cursor()
            for item in sqlArray:
                if item != '':
                    try:
                        item=self.rep(file,sheet,row,conn,item)
                        cursor.execute(item)
                        conn[0][0].commit()
                        self.getToLog('数据初始化：'+item)
                    except Exception as e:
                        print(e)
                        msg.append(str(column1))
                        self.getError(str(e))
                        SqlMsg.append(str(e))
                column1=column1+1
            cursor.close()
            return [msg,SqlMsg] if len(msg)>1 else []
        
    '''    
    @数据恢复
    @param file:用例文件
    @param sheet:  
    @param row: 行号
    @param conn:数据库连接对象 
    '''
    def restore(self,file,sheet,row,conn):
        msg = ['数据库异常']
        SqlMsg = []
        column1 = self.restore001Col
        sqlArray = self.getArray(file,sheet,row,self.restore001Col,self.dyparam001Col)
        for item in sqlArray:
            item=self.repVar(str(item))
            item=self.repRel(row,str(item))
        if conn == [[]]:
            return []
        else:
            cursor = conn[0][0].cursor()
            for item in sqlArray:
                if item != '':
                    try:
                        item=self.rep(file,sheet,row,conn,item)
                        cursor.execute(item)
                        conn[0][0].commit()
                        self.getToLog('数据恢复：'+item)
                    except Exception as e:
                        print(e)
                        msg.append(str(column1))
                        self.getError(str(e))
                        SqlMsg.append(str(e))
                column1=column1+1
            cursor.close()
            return [msg,SqlMsg] if len(msg)>1 else []
    
    '''    
    @动态化参数-数据库查询结果作为参数供同一行的其他地方调用
    @param file:用例文件
    @param sheet:  
    @param row: 行号
    @param conn: 数据库连接对象
    '''
    def dyparam(self,file,sheet,row,conn):
        data=[]
        sqlArray = self.getArray(file,sheet,row,self.dyparam001Col,self.key001Col)
        for item in sqlArray:
            item=self.repVar(str(item))
            item=self.repRel(row,str(item))
        if conn == [[]]:
            return []
        else:
            cursor = conn[0][0].cursor()
            for item in sqlArray:
                if(item == ''):
                    data.append('')
                else:
                    cursor.execute(item)
                    dd = cursor.fetchone()
                    data.append(None if dd==None else dd[0])
                    conn[0][0].commit()
            cursor.close()
            return data
            
    '''
    @用户变量替换
    @param param: 参数
    '''
    def repVar(self,param):
        for i in range(len(self.userParams)):
            if '${'+self.userParams[i]+'}' in param:
                param=param.replace('${'+str(self.userParams[i])+'}',str(self.userParamsValue[i]))
        return param    
    
            
    '''
    @接口变量替换
    @param row: 行号
    @param param:参数
    '''
    def repRel(self,row,param):
        for i in range(len(self.userVar)):
            if '${'+str(self.userVar[i])+'}' in str(param):
                param=param.replace('${'+str(self.userVar[i])+'}',str(self.userVarValue[i]))
        return param        
    
    '''
    @动态参数替换
    @param file:用例文件
    @param sheet:  
    @param row: 行号
    @param conn: 数据库连接对象
    @param param: 动态参数
    '''
    def rep(self,file,sheet,row,conn,param):
        dypArr=[]
        dypar=self.dyparam(file,sheet,row,conn)
        if dypar==None:
            return param
        else:
            for i in range(1,len(dypar)+1):
                dypArr.append('dyparam'+str(i).zfill(3))
            for i in range(len(dypar)):
                if '${'+str(dypArr[i])+'}' in param:
                    param=param.replace('${'+str(dypArr[i])+'}',str(dypar[i]))
        return param
    
    '''
    @三者替换，替换用户变量、动态参数、接口变量
    @param file:用例文件
    @param sheet:  
    @param row: 行号
    @param conn: 数据库连接对象
    '''
    def repAll(self,strValue,file,sheet,row,conn):
        strValue=self.repVar(str(strValue))
        strValue=self.rep(file,sheet,row,conn,strValue)
        strValue=self.repRel(row,strValue)
        return strValue
    