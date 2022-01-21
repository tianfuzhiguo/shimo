from common.excel.Write import Write
from common.excel.Report import Report
from common.excel.Template import Template
from apscheduler.schedulers.background import BackgroundScheduler 
from apscheduler.triggers.date import DateTrigger
from PyQt5.QtCore import QThread,Qt
from PyQt5 import QtCore,QtGui, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog
from common.ui.Ui_mainWindow import Ui_mainWindow
import time,datetime,sys,os,yagmail

'''
@主类
@author: dujianxiao
'''
class DetailUI(Ui_mainWindow,QMainWindow,Write,Report,Template):

    def resource_path(self,relative_path):
        base_path = getattr(sys,'_MEIPASS',os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)
    
    def __init__(self):
        self.img=self.resource_path(os.path.join(".","source/1.ico"))
        splash = QtWidgets.QSplashScreen(QtGui.QPixmap(self.img))
        splash.show()                           # 显示启动界面
        QtWidgets.qApp.processEvents()          # 处理主进程事件
        font = QtGui.QFont()
        font.setPointSize(16)
        font.setBold(True)
        font.setWeight(75)
        splash.setFont(font)
        self.load_data(splash)                # 加载数据
        
        super().__init__()
        
        splash.finish(self)                   # 隐藏启动界面
        self.setupUi(self)
        self.setWindowTitle('时默')
        now_time = datetime.datetime.now()
        ss=datetime.datetime.strptime(str(now_time)[:-7],'%Y-%m-%d %H:%M:%S') 
        self.taskTime.setMinimumDateTime(ss)
        '''
        @设置窗口标题栏图标
        '''
        filename = self.resource_path(os.path.join(".","source/1.ico"))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(filename), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.setWindowIcon(icon)
        '''
        @计数
        '''
        self.allRows=0
        self.status1=0#success
        self.status2=0#fail
        self.status3=0#skip
        
    '''
    @启动动画
    '''
    def load_data(self, sp):
        sp.showMessage("  正在加载...", Qt.AlignCenter,Qt.black)
        QtWidgets.qApp.processEvents()  # 允许主进程处理事件
        
    '''
    @打开用例文件
    '''
    def getFile(self):
        ex.console.clear()
        global sheetNames,sheet,bookRes,sheetRes,fileRes,path,file
        try:
            '''
            @加载文件时清除上一个文件可能存在的请求头信息、接口变量和用户变量
            '''
            self.headerManager=''
            self.userVar=[]
            self.userParams
            reportDate=time.strftime("%Y%m%d", time.localtime())
            self.example.clear()
            self.result.setText('0/0')
            self.initTextNum()
            
            '''
            @集成jenkins时，自动加载配置文件和用例文件，文件需与执行程序在同一目录下
            '''
            path= os.getcwd()
            path=path.replace('\\','/')
            a,b,c,userParamsValue=self.initConfig(path)
            self.initLog(path)
            file=str(userParamsValue[-1:])
            file=file[2:-2]
            fname=path+'/'+file
            sheetNames = self.getSheetNames(fname)
            self.fileName.setText(fname)
            '''
            @以上
            '''           
            
            '''
            @如果未选择文件，页签下拉列表置空
            '''       
            if fname in ['请选择文件','']:
                self.qSheetName.clear()
                self.console.clear()
            else:
                '''
                @初始化日志和配置文件
                '''
                self.initConfig(path)
                self.initLog(path)
                '''
                @创建用例结果文件
                '''
                sheetNames = self.getSheetNames(path+'/'+file)
                bookRes,sheetRes,fileRes=self.createReport(reportDate,path,file,sheetNames)
                self.qSheetName.clear()
                self.example.clear()
                fname=self.fileName.text()            
                '''
                @填充页签下拉列表
                '''
                for i in range(len(sheetNames)+1):
                    if i==0:
                        self.qSheetName.addItem('全部')
                    else:
                        self.qSheetName.addItem(str(sheetNames[i-1]))           
        except Exception as e:
            print(e)
            self.qSheetName.clear()
    
    '''
    @点击文件名打开文件
    '''
    def openExample(self):
        try:
            fname=self.fileName.text()
            if fname in ['请选择文件','']:
                ex.console.clear()
            else:
                os.startfile(eval('r'+"'"+fname+"'"))
        except Exception as e:
            print(e)
            ex.console.clear()
            
    '''
    @切换页签
    '''
    def changeSheet(self):
        self.initTextNum()
        ex.console.clear()
        self.example.clear()
        global sheetName,sheet,nrows,fname,path
        try:
            reportDate=time.strftime("%Y%m%d", time.localtime())
            sheetName=self.qSheetName.currentText() 
            bookRes,sheetRes,fileRes=self.createReport(reportDate,path,file,sheetNames)
            if (sheetName=='全部' and ex.qSheetName.currentIndex()==0) or sheetName=='':
                self.example.setCurrentText('')
                for i in range(len(sheetNames)):
                    self.example.clear
                    self.example.items.clear()
                self.example.loadItems([])
                allRows=0
                allRpt=''                
                '''
                @每次切换页签时都校验一遍模板，防止使用过程中对模板有改动
                '''                
                for i in range(len(sheetNames)):
                    sheet,nrows=self.initFile(reportDate,path,file,sheetNames[i])
                    rpt=ex.verTemp(sheetNames[i],sheet,bookRes,sheetRes[i],fileRes)
                    allRpt=allRpt+str(rpt)
                    if rpt=='':
                        noRuns=0
                        IterationCol=self.findStr(file,sheet,'Iteration')
                        for i in range(3,nrows+1):
                            if str(self.getValue(file,sheet,i-1, IterationCol))=='0':
                                noRuns=noRuns+1
                        allRows=allRows+nrows-2-noRuns
                if allRpt=='':
                    ex.result.setText('0/'+str(allRows))
            else:
                items=[]
                st=[]
                '''
                @每次切换页签时都校验一遍模板，防止使用过程中对模板有改动
                '''
                sheet,nrows=self.initFile(reportDate,path,file,sheetName)
                rpt=ex.verTemp(sheetName,sheet,bookRes,sheetRes[0],fileRes)
                noRuns = 0
                if rpt=='':
                    for i in range(3,nrows+1):
                        st.append(str(i)+' '+str(self.getValue(file,sheet,i-1,ex.nameCol)))
                        st.append(str(self.getValue(file,sheet,i-1,ex.IterationCol)))
                        items.append(st)
                        st=[]
                    self.example.loadItems(items)
                    IterationCol=self.findStr(file,sheet,'Iteration')
                    for i in range(3,nrows+1):
                        if str(self.getValue(file,sheet,i-1, IterationCol)).upper()=='0':
                            noRuns=noRuns+1
                    ex.result.setText('0/'+str(nrows-2-noRuns))
                else:
                    ex.console.clear()
                    ex.consoleFunc('red',str(rpt))                   
        except Exception as e:
            print(e)
            
    '''
    @每次点击页签下拉框重新加载下拉列表
    @考虑到性能问题,这个方法没有被调用
    @应同事要求，在界面加了一个刷新按钮调用此方法
    '''      
    def reloadSheet(self):
        try:
            '''
            @刷新时清除上一个文件可能存在的请求头信息、接口变量和用户变量
            '''
            self.headerManager=''
            self.userVar=[]
            self.userParams=[]
            fname=self.fileName.text()
            if fname in ['请选择文件','']:
                pass
            else:
                self.qSheetName.clear()
                sheetNames = self.getSheetNames(path+'/'+file)
                '''
                @填充页签下拉列表
                '''
                for i in range(len(sheetNames)+1):
                    if i==0:
                        self.qSheetName.addItem('全部')
                    else:
                        self.qSheetName.addItem(str(sheetNames[i-1]))
                self.qSheetName.setCurrentIndex(0)
        except Exception as e:
            print(e)
     
    '''
    @点击用例下拉框时重新填充下拉列表
    '''
    def reload(self):
        self.initTextNum()
        ex.console.clear()
        global sheetName,sheet,nrows,fname,path
        try:
            reportDate=time.strftime("%Y%m%d", time.localtime())
            sheetName=self.qSheetName.currentText() 
            if (sheetName=='全部' and ex.qSheetName.currentIndex()==0) or sheetName=='':
                self.example.setCurrentText('')
                for i in range(len(sheetNames)):
                    self.example.clear
                    self.example.items.clear()                    
                allRows=0
                allRpt=''                
                '''
                @每次切换页签时都校验一遍模板，防止使用过程中对模板有改动
                '''
                for i in range(len(sheetNames)):
                    sheet,nrows=self.initFile(reportDate,path,file,sheetNames[i])
                    rpt=ex.verTemp(sheetNames[i],sheet,bookRes,sheetRes[i],fileRes)
                    allRpt=allRpt+str(rpt)
                    if rpt=='':
                        noRuns = 0
                        IterationCol=self.findStr(file,sheet,'Iteration')
                        for i in range(3,nrows+1):
                            if str(self.getValue(file,sheet,i-1, IterationCol)).upper()=='0':
                                noRuns=noRuns+1
                        allRows=allRows+nrows-2-noRuns
                if allRpt=='':
                    ex.result.setText('0/'+str(allRows))
            else:
                '''
                @每次切换页签时都校验一遍模板，防止使用过程中对模板有改动
                '''
                items=[]
                st=[]
                sheet,nrows=self.initFile(reportDate,path,file,sheetName)
                rpt=ex.verTemp(sheetName,sheet,bookRes,sheetRes,fileRes)
                '''
                @模板校验通过
                '''
                if rpt=='':
                    #获取被选中的用例
                    exa=str(ex.example.currentText()).replace('(', '').replace(')','').replace("'",'').split(",")
                    for value in exa:
                        if value=='':
                            exa.remove(value)
                    exa=[int(i) for i in exa if i.isdigit()]
                    self.example.clear()
                    self.example.items.clear()
                    noRuns = 0
                    IterationCol=self.findStr(file,sheet,'Iteration')
                    for i in range(3,nrows+1):
                        if str(self.getValue(file,sheet,i-1, IterationCol)).upper()=='0':
                            noRuns=noRuns+1
                    for i in range(3,nrows+1):
                        st.append(str(i)+' '+str(self.getValue(file,sheet,i-1,ex.nameCol)))
                        st.append(str(self.getValue(file,sheet,i-1,ex.IterationCol)))
                        items.append(st)
                        st=[]
                    self.example.loadItems(items)    
                    '''
                    @保持上一次的选中状态
                    '''
                    if exa!=[]:
                        for i in range(len(exa)):
                            try:
                                self.example.qCheckBox[exa[i]-2].setChecked(True)
                            except Exception as e:
                                print(e)
                    ex.result.setText('0/'+str(nrows-2-noRuns))
                else:
                    ex.console.clear()
                    ex.consoleFunc('red',str(rpt))                  
        except Exception as e:
            print(e)
     
    '''
    @切换用例
    '''
    def changeExample(self):
        try:        
            exa=self.example.Selectlist()
            ex.result.setText('0/'+str(len(exa)))
        except Exception as e:
            print(e)
     
    '''
    @打开excel报告
    '''
    def openExcelReport(self):
        runTime=time.strftime("%Y%m%d", time.localtime())
        try:
            fname=self.fileName.text()
            if fname in ['请选择文件','']:
                ex.console.clear()
            else:
                reportName=file[:file.index('.xls')]+'-'+str(runTime)+'-report.xls'
                if file.endswith('xlsx'):
                    reportName=reportName=+'x'
                excel='r'+"'"+path+'/result/'+reportName+"'"
                os.startfile(eval(excel))
        except Exception as e:
            print(e)
            ex.console.clear()
            ex.consoleFunc('red','打开报告失败')            
            
    '''
    @创建html测试报告
    '''        
    def createHTMLReport(self,reportDate,js,file,path):
        try:
            html=self.resource_path("source/template")
            f1=open(html,"r",encoding="utf-8")
            htmlData = f1.read()
            html=htmlData.replace('${resultData}',str(js))
            f1.close()
            file=file[:file.index('.xls')]
            try:
                os.remove(path+'/result/'+file+'-'+str(reportDate)+'-report.html')
            except Exception as e:
                print(e)
            htmlReportName=path+'/result/'+file+'-'+str(reportDate)+'-report.html'
            f2 = open(htmlReportName,'w',encoding='utf-8')
            f2.write(html)
            f2.close()
            return htmlReportName
        except Exception as e:
            print(e)
    
    '''
    @打开html报告
    '''       
    def openHtmlReport(self):
        runTime=time.strftime("%Y%m%d", time.localtime())
        try:
            fname=self.fileName.text()
            if fname in ['请选择文件','']:
                ex.console.clear()
            else:
                reportName=file[:file.index('.xls')]+'-'+str(runTime)+'-report.html'
                sss='r'+"'"+path+'/result/'+reportName+"'"
                os.startfile(eval(sss))
        except Exception as e:
            print(e)
            ex.console.clear()
            ex.consoleFunc('red','打开报告失败')           
                 
    '''
    @打开日志
    '''      
    def openLog(self):
        try:
            fname=self.fileName.text()
            if fname in ['请选择文件','']:
                ex.console.clear()
            else:
                sss='r'+"'"+path+'/result/info.log'+"'"
                os.startfile(eval(sss))
        except Exception as e:
            print(e)
            ex.console.clear()
            ex.consoleFunc('red','打开日志失败')           
            
    '''
    @获取文件路径
    '''       
    def getPath(self,path):
        n=0
        s1=''
        s2=''
        for i in range(len(path)):
            if path[i:i+1]=='/':
                n=n+1
            if n==path.count('/')-1:
                s1=path[:i+1]
                s2=path[i+2:]
        return s1,s2  
    
    '''
    @定时任务或debug时勾选邮件时候发送邮件
    '''
    def sendEmail(self,htmlReportName):
        isMail=ex.checkMail.checkState()
        if isMail==2:            
            ex.consoleFunc('black','邮件发送中...')
            try:
                yag=yagmail.SMTP( user=self.email[0], password=self.email[2], host=self.email[1])
                receList=self.email[3].split(',')
                yag.send(receList,self.email[4],self.email[5],[htmlReportName])
                ex.consoleFunc('black','邮件发送成功')
                yag.close()
            except Exception as e:
                print(e)
                yag.close()
                ex.consoleFunc('red',str(e))
                ex.consoleFunc('red','邮件发送失败')               
    
    '''
    @重写窗口大小改变事件
    '''
    def resizeEvent(self,evt):
        res=self.result.text()
        ss=self.example.currentText()
        if ss==[]:
            ss=''
        elif '(' in str(ss):
            ss=ss.replace('(', '').replace(')', '').replace("'", "")
        try: 
            '''
            @给用例框赋值(非下拉列表)
            '''
            items=[]
            st=[]
            reportDate=time.strftime("%Y%m%d", time.localtime())
            sheet,nrows=self.initFile(reportDate,path,file,sheetName)
            for i in range(3,nrows+1):
                st.append(str(i)+' '+str(self.getValue(file,sheet,i-1,ex.nameCol)))
                st.append(str(self.getValue(file,sheet,i-1,ex.IterationCol)))
                items.append(st)
                st=[]
            [items.remove(str(item[0])) for item in items if str(item[0])=='全部']
        except Exception as e:
            print(e)
        self.example.loadItems(items)
        self.example.setCurrentText(str(ss))        
        '''
        @给预览结果赋值
        '''
        self.result.setText(res)      
        
    '''
    @接口解析
    '''
    def analyFunction(self):
        fname=self.fileName.text()
        if fname in ['请选择文件','']:
            pass
        else:
            ss=self.analyJSON.text()
            if ss == '解析JSON':
                ex.console.clear()
                ex.analy_thread=analyFunctionClass()
                ex.analy_thread.start()
                ex.analyJSON.setText('停止')
            elif ss=='停止':
                ex.analyJSON.setText('解析JSON')
                ex.analy_thread.terminate()
                self.buttonStatus(True)
                
    '''
    @执行用例
    '''
    def start(self):
        fname=self.fileName.text()
        if fname in ['请选择文件','']:
            pass
        else:
            ss=self.debug.text()
            if ss == '开始':
                ex.console.clear()
                ex.debug_thread=debugClass()
                ex.debug_thread.start()
                ex.debug.setText('停止')
            elif ss=='停止':
                ex.debug.setText('开始')
                ex.debug_thread.terminate()
                self.buttonStatus(True)

    '''
    @开始任务
    '''   
    def startTask(self):
        ex.console.clear()
        self.task_thread=taskClass()
        self.task_thread.start()
        
    '''
    @取消任务
    '''
    def abortTask(self):
        scheduler.shutdown()
        self.buttonStatus(True)
        ex.console.clear()
        ex.consoleFunc('black','定时任务已取消')

    '''
    @设置按钮状态
    @param flag:True/False
    '''
    def buttonStatus(self,flag):
        ex.task.setEnabled(flag)
        ex.abort.setEnabled(flag)
        ex.taskTime.setEnabled(flag)
        ex.debug.setEnabled(flag)
        ex.file.setEnabled(flag)
        ex.analyJSON.setEnabled(flag)
        ex.dtailReport.setEnabled(flag)
        ex.html.setEnabled(flag)
        ex.dtailLog.setEnabled(flag)
        ex.qSheetName.setEnabled(flag)
        ex.refresh.setEnabled(flag)
    
    '''
    @初始化成功、失败、异常数量为0
    '''
    def initTextNum(self):
        ex.successNum.setText('0')
        ex.failNum.setText('0')
        ex.skipNum.setText('0')
'''
@接口解析
'''        
class analyFunctionClass(QThread,DetailUI):
        
    def __init__(self):
        super(DetailUI, self).__init__()
    
    def run(self):
        try:
            ex.buttonStatus(False)
            ex.analyJSON.setEnabled(True)
            self.initTextNum()
            ex.result.setText('0/0')
            exa=ex.example.currentText()
            if exa==[]:
                ex.consoleFunc('red', '请选择接口')               
            else:
                exa=exa.replace("'", '').replace('(', '').replace(')','').split(',')
                exa=[int(item) for item in exa]
                ex.initConfig(path)
                for i in range(len(exa)):                    
                    ex.analyFunc(file,exa[i]-1, sheetName,sheet)
        except Exception as e:
            print(e)
        self.buttonStatus(True)
        ex.analyJSON.setText('解析JSON')    
            
class debugClass(QThread,DetailUI):
        
    def __init__(self):
        super(DetailUI, self).__init__()

    def run(self):
        dict={}
        htmlReportName=''
        ex.runFlag=True
        cText=ex.debug.text()
        if cText=='停止':
            self.initTextNum()
            ex.status1=0#success
            ex.status2=0#fail
            ex.status3=0#skip
            ex.allRows=0
        try:
            reportDate=time.strftime("%Y%m%d", time.localtime())
            model='普通' if ex.model1.isChecked() else '简洁'
            startTime = datetime.datetime.now()
            self.initTextNum()           
            text=ex.result.text()
            ex.result.setText('0'+text[text.index('/'):])
            ex.buttonStatus(False)
            ex.debug.setEnabled(True)
            ex.initConfig(path)
            bookRes,sheetRes,fileRes=ex.createReport(reportDate,path,file,sheetNames)
            sheetValue=ex.qSheetName.currentText()
            allRpt=''
            testResult=[]
            '''
            @全量
            '''
            if sheetValue=='全部' and ex.qSheetName.currentIndex()==0:
                
                for i in range(len(sheetNames)):
                    sheet,nrows=ex.initFile(reportDate,path,file,sheetNames[i])   
                    rpt=ex.verTemp(sheetNames[i],sheet,bookRes,sheetRes[i],fileRes)
                    '''
                    @模板校验通过
                    '''
                    if rpt=='':
                        noRuns=0
                        '''
                        @找出迭代次数为0（不执行）的用例
                        '''
                        for i in range(3,nrows+1):
                            if str(self.getValue(file,sheet,i-1, ex.IterationCol)).upper()=='0':
                                noRuns=noRuns+1
                        '''
                        @全部用例数为各页签的用例数相加减去不执行的用例数
                        '''
                        ex.allRows=ex.allRows+nrows-2-noRuns#全部用例数
                    allRpt=allRpt+str(rpt)
                '''
                @模板全部校验通过
                '''
                if allRpt=='':                  
                    for i in range(len(sheetNames)):
                        sheet,nrows=ex.initFile(reportDate,path,file,sheetNames[i])
                        '''
                        @执行该页签中的用例
                        '''
                        dict,tr=ex.run(model,'',sheetNames[i],sheet,nrows,bookRes,sheetRes[i],fileRes,ex.allRows)
                        testResult=testResult+tr
            else: #单个或多个
                for i in range(len(sheetNames)):
                    if sheetValue==sheetNames[i]:
                        sheetRes=sheetRes[i]
                        break
                sheet,nrows=ex.initFile(reportDate,path,file,sheetName)
                rpt=ex.verTemp(sheetName,sheet,bookRes,sheetRes,fileRes)
                '''
                @模板校验通过才进行之后的操作
                '''
                if rpt=="":
                    exa=ex.example.currentText()
                    en=[]
                    '''
                    @未选中任何用例则执行该页签的全部用例
                    '''
                    if exa==[]:
                        for i in range(3,nrows+1):
                            '''
                            @迭代次数为0表示此用例不执行
                            '''
                            if str(self.getValue(file,sheet,i-1, ex.IterationCol))!='0':
                                exa.append(i)
                    else:
                        '''
                        @是否存在大于当前页签行数的接口号(删除用例引起)
                        '''
                        exa=exa.replace("'", '').replace('(', '').replace(')','').split(',')
                        exa=[int(item) for item in exa]
                        en=[item for item in exa if item>nrows]
                    '''
                    @如果当前选中的用例(序列号大于nrows)已被删除,则提示
                    '''
                    if en != []:
                        ex.consoleFunc('red', '用例'+str(en)+"不存在")
                    else:
                        for item in exa:
                            if str(self.getValue(file,sheet,item-1, ex.IterationCol))!='0':#不执行迭代次数为0的用例
                                dict,tr=ex.run(model,item,sheetName,sheet,nrows,bookRes,sheetRes,fileRes,len(exa))
                                testResult=testResult+tr
            '''
            @格式化html报告中的运行时间和时长
            '''
            endTime = datetime.datetime.now()
            second=str(endTime-startTime)
            duration=second[:second.index('.')]
            dd=duration.split(':')
            duration=dd[0]+'小时 '+dd[1]+'分 '+dd[2]+'秒'
            taskName=file[:-4] if file.endswith('xls') else file[:-5]           
            '''
            @测试结果存到字典中，用于html测试报告
            '''
            dict['testName']=taskName#项目名称
            startTime=str(startTime)
            dict['beginTime']=startTime[:startTime.index('.')]#开始时间
            dict['totalTime']=duration#运行时长
            dict['testResult']=testResult#结果集
            htmlReportName=ex.createHTMLReport(reportDate,dict,file,path)
            ex.sendEmail(htmlReportName)
        except Exception as e:
            print(e)
        ex.status1=0#success
        ex.status2=0#fail
        ex.status3=0#skip
        ex.allRows=0
        self.buttonStatus(True)
        ex.debug.setText('开始')  
        '''
        @此行代码用于集成jenkins时，当用例执行完毕后自动退出程序
        '''
        sys.exit()
            
'''
@定时任务
@全量执行－－除了迭代次数为0的
'''
class taskClass(QThread,DetailUI): 
    def __init__(self):
        super(DetailUI, self).__init__()
        
    def run(self):
        global scheduler
        try:
            fname=ex.fileName.text()
            if fname in ['请选择文件','']:
                pass
            else:
                inTime=ex.taskTime.text()
                now=datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                flag=True
                '''
                @定时任务时间不得小于当前时间
                '''
                if inTime<now:
                    flag=False
                    ex.consoleFunc('red', '时间不得小于当前时间,请重新输入')                    
                if flag==True:
                    trigger = DateTrigger(inTime)
                    scheduler = BackgroundScheduler ()
                    scheduler.add_job(self.taskJob, trigger)
                    ex.consoleFunc('black', '定时任务将于'+str(inTime)+'执行:')
                    ex.task.setEnabled(False)
                    ex.taskTime.setEnabled(False)
                    scheduler.start()
        except Exception as e:
            print(e)
                   
    def taskJob(self):
        dict={}
        htmlReportName=''
        try:
            ex.status1=0#success
            ex.status2=0#fail
            ex.status3=0#skip
            ex.allRows=0
            fname=ex.fileName.text()
            reportDate=time.strftime("%Y%m%d", time.localtime())
            if fname in ['请选择文件','']:
                pass
            else:
                ex.consoleFunc('black', '定时任务开始:')
                startTime = datetime.datetime.now()
                self.initTextNum()
                ex.result.setText('0/0')
                ex.buttonStatus(False)
                bookRes,sheetRes,fileRes=ex.createReport(reportDate,path,file,sheetNames)
                allRpt=''
                try:
                    '''
                    @定时任务只支持执行全部用例，
                    '''
                    model='普通' if ex.model1.isChecked() else '简洁'
                    for i in range(len(sheetNames)):
                        sheet,nrows=ex.initFile(reportDate,path,file,sheetNames[i])
                        rpt=ex.verTemp(sheetNames[i],sheet,bookRes,sheetRes[i],fileRes)
                        if rpt=='':
                            noRuns=0
                            for i in range(3,nrows+1):
                                if str(self.getValue(file,sheet,i-1,ex.IterationCol)).upper()=='0':
                                    noRuns=noRuns+1
                            ex.allRows=ex.allRows+nrows-2-noRuns
                        allRpt=allRpt+str(rpt)
                    '''
                    @所有页签的模板校验通过
                    '''
                    if allRpt=='':
                        testResult=[]
                        for i in range(len(sheetNames)):
                            sheet,nrows=ex.initFile(reportDate,path,file,sheetNames[i])
                            dict,tr=ex.run(model,'',sheetNames[i],sheet,nrows,bookRes,sheetRes[i],fileRes,ex.allRows)
                            testResult=testResult+tr
                except Exception as e:
                    print(e)
                    ex.consoleFunc('red','定时任务执行失败')                                      
                endTime = datetime.datetime.now()
                second=str(endTime-startTime)
                duration=second[:second.index('.')]
                dd=duration.split(':')
                duration=dd[0]+'小时 '+dd[1]+'分 '+dd[2]+'秒'
                taskName=file[:-4] if file.endswith('xls') else file[:-5]
                '''
                @测试结果存到字典中，用于html测试报告
                '''
                dict['testName']=taskName#项目名称
                startTime=str(startTime)
                dict['beginTime']=startTime[:startTime.index('.')]#开始时间
                dict['totalTime']=duration#运行时长
                dict['testResult']=testResult#结果集
                htmlReportName=ex.createHTMLReport(reportDate,dict,file,path)
                ex.sendEmail(htmlReportName)               
                ex.consoleFunc('black', '定时任务执行成功')
                scheduler.shutdown()
        except Exception as e:
            print(e)
        ex.allRows=0
        ex.status1=0#success
        ex.status2=0#fail
        ex.status3=0#skip
        self.buttonStatus(True)
            
if __name__ == "__main__":
    app=0
    app = QApplication(sys.argv) 
    ex = DetailUI()
    ex.show()
    ex.getFile()
    ex.start()
    sys.exit(app.exec_())
    