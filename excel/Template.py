from common.init.Init import Init
'''
@模板校验:校验各关键字是否存在，顺序是否正确，是否有重复关键字;校验各部分数量是否一致
@author: dujianxiao
'''
class Template(Init):
    '''
    @获取各标志位之间的数组个数
    @param start:开始列
    @param end:  结束列
    '''
    def getArrLenth(self,start, end):
        return [column for column in range(start, end)]
    
    '''
    @模板校验
    @param sheetName:页签名
    @param sheet:用例文件
    @param book:用例结果文件
    @param sheetRes:用例结果文件
    @param fileRes:用例结果文件
    '''
    def verTemp(self,sheetName,sheet,bookRes,sheetRes,fileRes):
        msg=self.verKeyWordExist(fileRes,sheet)
        blue=self.setCellStyle(7)
        if '未找到关键字' in str(msg):
            self.consoleFunc('green', sheetName+':', 'size=4')
            self.consoleFunc('red', str(msg))
            return msg
        elif '存在重复的关键字' in str(msg):
            self.consoleFunc('green', sheetName+':', 'size=4')
            self.consoleFunc('red', str(msg))
            return msg
        elif '关键字顺序不正确' in str(msg):
            ss="['关键字顺序不正确',"+"'"+str(self.getValue(fileRes,sheet,1,msg[1]))+"','"+str(self.getValue(fileRes,sheet,1,msg[2]))+"']"
            if fileRes.endswith('xls'):
                sheetRes.write(1,msg[1],self.getValue(fileRes,sheet,1, msg[1]),blue)
                sheetRes.write(1,msg[2],self.getValue(fileRes,sheet,1, msg[2]),blue)
            elif fileRes.endswith('xlsx'):
                self.setValueColor(sheetRes,2,msg[1],self.getValue(fileRes,sheet,1,msg[1]),"blue")
                self.setValueColor(sheetRes,2,msg[2],self.getValue(fileRes,sheet,1,msg[2]),"blue")
            bookRes.save(fileRes)
            self.consoleFunc('green', sheetName+':','size=4')
            self.consoleFunc('red', str(ss))
            return msg
        else:
            info=self.verLength()
            if '数量不一致' in str(info):
                print('校验字段和预期结果数量不一致')
                for i in range(1,len(info)):
                    if fileRes.endswith('.xls'):
                        sheetRes.write(1,int(info[i]),self.getValue(fileRes,sheet,1,int(info[i])),blue)
                    elif fileRes.endswith('.xlsx'):
                        self.setValueColor(sheetRes,2,int(info[i]),self.getValue(fileRes,sheet,1,int(info[i])),"blue")
                bookRes.save(fileRes)
                self.consoleFunc('green', sheetName+':','size=4')
                self.consoleFunc('red', str(info))
                return info
            else:
                return ''
       
    '''
    @校验各关键字是否存在，顺序是否正确，是否有重复关键字
    @param file:用例文件
    @param sheet:用例文件
    '''
    def verKeyWordExist(self,file,sheet):
        '''
        @定义关键字数组
        '''
        arr=['name','url','method','param','file','header','part101','part201','part301','section101','section201',
             'section301','resText','resHeader','statusCode','expression','status','time','init001','restore001','dyparam001',
             'key001','value001','headerManager','数据库','Iteration']
        arrCopy=[]
        order=[]
        msg=[]
        sts="未找到关键字"
        msg.append(sts)
        keyWord=''
        repeat = ['存在重复的关键字']
        for item in arr:
            for i in range(self.ncols):
                try:
                    if file.endswith('xls'):
                        keyWord = sheet.cell(1, i).value
                    elif file.endswith('xlsx'):
                        keyWord = sheet.cell(row=2, column=i).value
                except Exception as e:
                    print(e)
                if(item == keyWord):
                    order.append(self.findStr(file,sheet,item))
                    if len(arrCopy)>0:
                        for k in range(len(arrCopy)):
                            if item == arrCopy[k]:
                                repeat.append(item)
                                break
                    arrCopy.append(item)                  
            '''
            @如果没找到关键字，则返回未找到的关键字
            '''
            if item not in arrCopy:
                msg.append(item)
        if len(repeat)>1:
            return repeat
        if len(msg)>1:
            return msg
        '''
        @找到全部关键字则校验顺序
        '''
        msg=[]
        st="关键字顺序不正确"
        msg.append(st)
        for i in range(len(order)):
            try:
                if i+1<len(order) and order[i+1]<order[i]:
                    msg.append(order[i])
                    msg.append(order[i+1])
            except Exception as e:
                print(e)
        return msg if len(msg)>1 else ''
        
    '''
    @校验各部分数量是否一致
    '''
    def verLength(self): 
        msg=[]
        st='数量不一致'
        msg.append(st)
        '''
        @校验字段
        '''
        len1=self.getArrLenth(self.part101Col, self.part201Col)
        len2=self.getArrLenth(self.part201Col, self.part301Col)
        len3=self.getArrLenth(self.part301Col, self.section101Col)
        '''
        @预期结果
        '''
        len4=self.getArrLenth(self.section101Col, self.section201Col)
        len5=self.getArrLenth(self.section201Col, self.section301Col)
        len6=self.getArrLenth(self.section301Col, self.resTextCol)
        '''
        @接口变量
        '''
        len7=self.getArrLenth(self.key001Col, self.value001Col)
        len8=self.getArrLenth(self.value001Col, self.headerManagerCol)
        if len(len1)!=len(len4):
            return msg+len1+len4
        elif len(len2)!=len(len5):
            return msg+len2+len5
        elif len(len3)!=len(len6):
            return msg+len3+len6
        elif len(len7)!=len(len8):
            return msg+len7+len8
        else:
            return ''
    
    