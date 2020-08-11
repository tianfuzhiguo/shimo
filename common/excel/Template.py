from common.init.Init import Init
from common.utils.Util import setStyle,getValue,findStr

'''
@模板校验
@author: dujianxiao
'''
class Template(Init):
    '''
    @获取各标志位之间的数组个数
    @param start:
    @param end:  
    '''
    def getArrLenth(self,start, end):
        return [column for column in range(start, end)]
    
    '''
    @模板校验
    @param file:用例文件
    @param sheetName:页签名
    @param sheet:
    @param ncols:列数
    @param book:
    @param sheet1:
    @param fileRes:用例结果文件
    @param column:列号        
    '''
    def verTemp(self,file,sheetName,sheet,ncols,book,sheet1,fileRes,column):
        msg=self.verKeyWordExist(file,sheetName,sheet,ncols)
        blue=setStyle(7)
        if '未找到关键字' in str(msg):
            self.console.append("<font size=4 color=green>"+sheetName+':'+"</font>")
            self.console.append("<font color=\"#000000\"></font> ")
            self.console.append("<font color=\"#FF0000\">"+str(msg)+"</font>")
            self.console.append("<font color=\"#000000\"></font> ")
            return msg
        elif '存在重复的关键字' in str(msg):
            self.console.append("<font size=4 color=green>"+sheetName+':'+"</font>")
            self.console.append("<font color=\"#000000\"></font> ")
            self.console.append("<font color=\"#FF0000\">"+str(msg)+"</font>")
            self.console.append("<font color=\"#000000\"></font> ")
            return msg
        elif '关键字顺序不正确' in str(msg):
            ss="['关键字顺序不正确',"+"'"+str(getValue(file,sheet,1,msg[1]))+"','"+str(getValue(file,sheet,1,msg[2]))+"']"
            print(ss)
            if fileRes[-4:]=='.xls':
                sheet1.write(1,msg[1],getValue(file,sheet,1, msg[1]),blue)
                sheet1.write(1,msg[2],getValue(file,sheet,1, msg[2]),blue)
            elif fileRes[-5:]=='.xlsx':
                self.setValueColor(sheet1,2,msg[1],getValue(file,sheet,1,msg[1]),"blue")
                self.setValueColor(sheet1,2,msg[2],getValue(file,sheet,1,msg[2]),"blue")
            
            book.save(fileRes)
            self.console.append("<font size=4 color=green>"+sheetName+':'+"</font>")
            self.console.append("<font color=\"#000000\"></font> ")
            self.console.append("<font color=\"#FF0000\">"+str(ss)+"</font>")
            self.console.append("<font color=\"#000000\"></font> ")
            return msg
        else:
            info=self.verLength(sheetName,column)
            if '数量不一致' in str(info):
                print('校验字段和预期结果数量不一致')
                for i in range(1,len(info)):
                    if fileRes[-4:]=='.xls':
                        sheet1.write(1,int(info[i]),getValue(file,sheet,1,int(info[i])),blue)
                    elif fileRes[-5:]=='.xlsx':
                        self.setValueColor(sheet1,2,int(info[i]),getValue(file,sheet,1,int(info[i])),"blue")
                book.save(fileRes)
                self.console.append("<font size=4 color=green>"+sheetName+':'+"</font>")
                self.console.append("<font color=\"#000000\"></font> ")
                self.console.append("<font color=\"#FF0000\">"+str(info)+"</font>")
                self.console.append("<font color=\"#000000\"></font> ")
                return info
            else:
                return ''
       
    '''
    @校验各关键字是否存在，顺序是否正确，是否有重复关键字
    @param file:用例文件
    @param sheetName:页签名
    @param sheet:
    @param ncols:列数
    '''
    def verKeyWordExist(self,file,sheetName,sheet,ncols):
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
            for i in range(0, ncols):
                try:
                    if file[-4:]=='.xls':
                        keyWord = sheet.cell(1, i).value
                    elif file[-5:]=='.xlsx':
                        keyWord = sheet.cell(row=2, column=i).value
                except Exception:
                    pass
                if(item == keyWord):
                    order.append(findStr(file,sheet,ncols,item))
                    if len(arrCopy)>0:
                        for k in range(0,len(arrCopy)):
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
        for i in range(0,len(order)):
            try:
                if order[i+1]<order[i]:
                    msg.append(order[i])
                    msg.append(order[i+1])
            except Exception:
                pass
            
        return msg if len(msg)>1 else ''
        
    '''
    @校验各部分数量是否一致
    @param sheetName:页签名
    @param column:列号
    '''
    def verLength(self,sheetName,column): 
        msg=[]
        st='数量不一致'
        msg.append(st)
        '''
        @校验字段
        '''
        len1=self.getArrLenth(column[5], column[6])
        len2=self.getArrLenth(column[6], column[7])
        len3=self.getArrLenth(column[7], column[8])
        '''
        @预期结果
        '''
        len4=self.getArrLenth(column[8], column[9])
        len5=self.getArrLenth(column[9], column[10])
        len6=self.getArrLenth(column[10], column[11])
        '''
        @接口变量
        '''
        len7=self.getArrLenth(column[19], column[20])
        len8=self.getArrLenth(column[20], column[21])
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
    
    