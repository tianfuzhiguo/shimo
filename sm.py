import time
import datetime
import sys
import os
from common.utils.Log import *
from common.init.Init import Init
from common.excel.Write import Write
from common.excel.Template import Template
from common.utils.Log import initLog
from common.utils.Util import getValue,readExcel,getSheetNames,findStr
import yagmail
from apscheduler.schedulers.background import BackgroundScheduler 
from apscheduler.triggers.date import DateTrigger
from PyQt5.QtCore import QThread,Qt
from PyQt5 import QtGui, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog
from common.ui.Ui_mainWindow import Ui_mainWindow

'''
@主类
@author: dujianxiao
'''
class DetailUI(Ui_mainWindow,QMainWindow,Write,Template,Init):
    
    def resource_path(self,relative_path):
        base_path = getattr(sys,'_MEIPASS',os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)
    
    
    def __init__(self):
        img=self.resource_path(os.path.join(".","source/1.ico"))
        splash = QtWidgets.QSplashScreen(QtGui.QPixmap(img))
        splash.show()                           # 显示启动界面
        QtWidgets.qApp.processEvents()          # 处理主进程事件
        font = QtGui.QFont()
        font.setPointSize(16)
        font.setBold(True)
        font.setWeight(75)
        splash.setFont(font)
        self.load_data(splash)                # 加载数据
        
        super(DetailUI, self).__init__()
        
        splash.finish(self)                   # 隐藏启动界面
        self.setupUi(self)
        self.setWindowTitle('dujianxiao7@163.com')
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
        global sheetNames,sheetName,fileData,email,userParams,userParamsValue,data,sheet,nrows,ncols,book,sheet1,fileRes,column,fname,path,file,exampleList
        try:
            reportDate=time.strftime("%Y%m%d", time.localtime())
            self.example.clear()
            '''
            @支持两种格式的excel文件
            '''
            fname , _ = QFileDialog.getOpenFileName(self, 'open file', '/',"files (*.xls *.xlsx)")
            self.fileName.setToolTip(fname)
            '''
            @获取文件的路径和名称，供其他方法使用
            '''
            self.fileName.setText(fname)
            path,file=self.getPath(fname)
            '''
            @初始化日志、配置文件
            '''
            fileData,email,userParams,userParamsValue=self.initConfig(path)
            initLog(path)
            '''
            @读取文件内容
            '''
            data = readExcel(path+'/'+file)
            sheetNames = getSheetNames(file,data)
            book,sheet1,fileRes=self.createReport(reportDate,path, file, data, sheetNames)
            self.qSheetName.clear()
            self.example.clear()
            fname=self.fileName.text()
            '''
            @填充页签下拉列表
            '''
            for i in range(0,len(sheetNames)+1):
                if i==0:
                    self.qSheetName.addItem('全部')
                else:
                    self.qSheetName.addItem(str(sheetNames[i-1]))
                 
            '''
            @如果未选择文件，页签下拉列表置空
            '''       
            if fname=='请选择文件' or fname=='':
                self.qSheetName.clear()
                self.console.clear()
            self.successNum.setText('0')
            self.failNum.setText('0')
            self.skipNum.setText('0')
        except Exception as e:
            print(e)
            self.qSheetName.clear()
    
    '''
    @点击文件名，打开文件
    '''
    def openExample(self):
        global sheetName,fileData,email,userParams,userParamsValue,data,sheet,nrows,ncols,column,UVCol,fname,path                          
        try:
            fname=self.fileName.text()
            if fname=='请选择文件' or fname=='':
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
        self.successNum.setText('0')
        self.failNum.setText('0')
        self.skipNum.setText('0')
        ex.console.clear()
        self.example.clear()
        global sheetName,fileData,email,userParams,userParamsValue,data,sheet,nrows,ncols,column,fname,path
        try:
            reportDate=time.strftime("%Y%m%d", time.localtime())
            sheetName=self.qSheetName.currentText() 
            book,sheet1,fileRes=self.createReport(reportDate,path, file, data, sheetNames)
            if (sheetName=='全部' and ex.qSheetName.currentIndex()==0) or sheetName=='':
                self.example.setCurrentText('')
                for i in range(0,len(sheetNames)):
                    self.example.clear
                    self.example.items.clear()
                self.example.loadItems([])
                allRows=0
                allRpt=''
                
                '''
                @每次切换页签时都校验一遍模板，防止使用过程中对模板有改动
                '''
                
                for i in range(0,len(sheetNames)):
                    fileData,email,userParams,userParamsValue,data,sheet,nrows,ncols,column=self.init(reportDate,path,file,sheetNames[i])
                    rpt=ex.verTemp(file,sheetNames[i],sheet,ncols,book,sheet1[i],fileRes,column)
                    allRpt=allRpt+str(rpt)
                    if rpt=='':
                        noRuns=0
                        IterationCol=findStr(file,sheet,ncols,'Iteration')
                        for i in range(3,nrows+1):
                            if str(getValue(file,sheet,i-1, IterationCol))=='0':
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
                fileData,email,userParams,userParamsValue,data,sheet,nrows,ncols,column=self.init(reportDate,path,file,sheetName)
                rpt=ex.verTemp(file,sheetName,sheet,ncols,book,sheet1[0],fileRes,column)
                noRuns = 0
                if rpt=='':
                    for i in range(3,nrows+1):
                        st.append(str(i)+' '+str(getValue(file,sheet,i-1,column[23])))
                        st.append(str(getValue(file,sheet,i-1,column[24])))
                        items.append(st)
                        st=[]
                    self.example.loadItems(items)
                    IterationCol=findStr(file,sheet,ncols,'Iteration')
                    for i in range(3,nrows+1):
                        if str(getValue(file,sheet,i-1, IterationCol)).upper()=='0':
                            noRuns=noRuns+1
                    ex.result.setText('0/'+str(nrows-2-noRuns))
                else:
                    ex.console.clear()
                    ex.console.append("<font color=\"#FF0000\">"+str(rpt)+"</font> ")
                    ex.console.append("<font color=\"#000000\"></font>")
        except Exception as e:
            print(e)
            
    '''
    @每次点击页签下拉框重新加载下拉列表
    @考虑到性能问题,这个方法没有被调用
    @应同事要求，在界面加了一个刷新按钮调用此方法
    '''      
    def reloadSheet(self):
        try:
            fname=self.fileName.text()
            if fname=='请选择文件' or fname=='':
                pass
            else:
                self.qSheetName.clear()
                '''
                @读取文件内容
                '''
                data = readExcel(path+'/'+file)
                sheetNames = getSheetNames(file,data)
                '''
                @填充页签下拉列表
                '''
                for i in range(0,len(sheetNames)+1):
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
        self.successNum.setText('0')
        self.failNum.setText('0')
        self.skipNum.setText('0')
        ex.console.clear()
        global sheetName,fileData,email,userParams,userParamsValue,data,sheet,nrows,ncols,column,fname,path
        try:
            reportDate=time.strftime("%Y%m%d", time.localtime())
            sheetName=self.qSheetName.currentText() 
            if (sheetName=='全部' and ex.qSheetName.currentIndex()==0) or sheetName=='':
                self.example.setCurrentText('')
                for i in range(0,len(sheetNames)):
                    self.example.clear
                    self.example.items.clear()
                    
                allRows=0
                allRpt=''
                
                '''
                @每次切换页签时都校验一遍模板，防止使用过程中对模板有改动
                '''
                for i in range(0,len(sheetNames)):
                    fileData,email,userParams,userParamsValue,data,sheet,nrows,ncols,column=self.init(reportDate,path,file,sheetNames[i])
                    rpt=ex.verTemp(file,sheetNames[i],sheet,ncols,book,sheet1[i],fileRes,column)
                    allRpt=allRpt+str(rpt)
                    if rpt=='':
                        noRuns = 0
                        IterationCol=findStr(file,sheet,ncols,'Iteration')
                        for i in range(3,nrows+1):
                            if str(getValue(file,sheet,i-1, IterationCol)).upper()=='0':
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
                fileData,email,userParams,userParamsValue,data,sheet,nrows,ncols,column=self.init(reportDate,path,file,sheetName)
                rpt=ex.verTemp(file,sheetName,sheet,ncols,book,sheet1,fileRes,column)
                '''
                @模板校验通过
                '''
                if rpt=='':
                    exa=ex.example.currentText()#选中的用例
                    self.example.clear()
                    self.example.items.clear()
                    noRuns = 0
                    IterationCol=findStr(file,sheet,ncols,'Iteration')
                    for i in range(3,nrows+1):
                        if str(getValue(file,sheet,i-1, IterationCol)).upper()=='0':
                            noRuns=noRuns+1
                    for i in range(3,nrows+1):
                        st.append(str(i)+' '+str(getValue(file,sheet,i-1,column[23])))
                        st.append(str(getValue(file,sheet,i-1,column[24])))
                        items.append(st)
                        st=[]
                    self.example.loadItems(items)
                    
                    '''
                    @保持上一次的选中状态
                    '''
                    if exa!=[]:
                        exa=exa.replace("'", '').replace('(', '').replace(')','')
                        exa=exa.split(',')
                        for i in range(0,len(exa)):
                            try:
                                self.example.qCheckBox[int(exa[i])-2].setChecked(True)
                            except Exception:
                                pass
                    ex.result.setText('0/'+str(nrows-2-noRuns))
                else:
                    ex.console.clear()
                    ex.console.append("<font color=\"#FF0000\">"+str(rpt)+"</font> ")
                    ex.console.append("<font color=\"#000000\"></font>")
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
            if fname=='请选择文件' or fname=='':
                ex.console.clear()
            else:
                ss=os.listdir(path+'/result/')
                reportName=''
                fffile=''
                if file[-4:]=='.xls':
                    fffile=file[:-4]
                    reportName=fffile+'-'+str(runTime)+'-report.xls'
                elif file[-5:]=='.xlsx':
                    fffile=file[:-5]
                    reportName=fffile+'-'+str(runTime)+'-report.xlsx'
                if reportName!='':
                    sss='r'+"'"+path+'/result/'+reportName+"'"
                    os.startfile(eval(sss))
                else:
                    ex.console.clear()
                    ex.console.append("<font color=\"#FF0000\">"+'打开报告失败'+"</font> ")
                    ex.console.append("<font color=\"#000000\"></font>")
        except Exception as e:
            print(e)
            
    '''
    @创建html测试报告
    '''        
    def createHTMLReport(self,reportDate,js,file,path):
        try:
            ss=self.resource_path("source/template")
            f1=open(ss,"r",encoding="utf-8")
            data = f1.read()
            ss=data.replace('${resultData}',str(js))
            f1.close()
            htmlReportName=''   
            if file[-4:]=='.xls':
                file=file[:-4]
            elif file[-5:]=='.xlsx':
                file=file[:-5]      
            try:
                os.remove(path+'/result/'+file+'-'+str(reportDate)+'-report.html')
            except Exception as e:
                print(e)
            htmlReportName=path+'/result/'+file+'-'+str(reportDate)+'-report.html'
            f2 = open(htmlReportName,'w',encoding='utf-8')
            f2.write(ss)
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
            if fname=='请选择文件' or fname=='':
                ex.console.clear()
            else:
                reportName=''
                realFileName=''
                if file[-4:]=='.xls':
                    realFileName=file[:-4]
                elif file[-5:]=='.xlsx':
                    realFileName=file[:-5]
                reportName=realFileName+'-'+str(runTime)+'-report.html'
                
                
                if reportName!='':
                    sss='r'+"'"+path+'/result/'+reportName+"'"
                    os.startfile(eval(sss))
                else:
                    ex.console.clear()
                    ex.console.append("<font color=\"#FF0000\">"+'打开报告失败'+"</font> ")
                    ex.console.append("<font color=\"#000000\"></font>")
        except Exception as e:
            print(e)
            
            
    '''
    @打开日志
    '''      
    def openLog(self):
        try:
            fname=self.fileName.text()
            if fname=='请选择文件' or fname=='':
                ex.console.clear()
            else:
                sss='r'+"'"+path+'/result/info.log'+"'"
                os.startfile(eval(sss))
        except Exception as e:
            print(e)
            ex.console.clear()
            ex.console.append("<font color=\"#FF0000\">"+'打开日志失败'+"</font> ")
            ex.console.append("<font color=\"#000000\"></font>")
            
            
    '''
    @获取文件路径
    '''       
    def getPath(self,path):
        n=0
        s1=''
        s2=''
        for i in range(0,len(path)):
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
            ex.console.append("<font color=\"#000000\"></font>")
            ex.console.append("<font color=\"#000000\">"+'邮件发送中...'+"</font>")
            try:
                yag = yagmail.SMTP( user=email[0], password=email[2], host=email[1])
                yag.send(eval(email[3]),email[4],email[5],[htmlReportName])
                ex.console.append("<font color=\"#000000\">"+'邮件发送成功'+"</font>")
                yag.close()
            except Exception as e:
                print(e)
                yag.close()
                ex.console.append("<font color=\"#FF0000\">"+str(e)+"</font> ")
                ex.console.append("<font color=\"#FF0000\">"+'邮件发送失败'+"</font> ")
                ex.console.append("<font color=\"#000000\"></font>")
            
    
    '''
    @重写窗口大小改变事件
    '''
    def resizeEvent(self,evt):
        res=self.result.text()
        ss=self.example.currentText()
        if ss==[]:
            ss=''
        elif '(' in str(ss):
            ss=ss.replace('(', '')
            ss=ss.replace(')', '')
            ss=ss.replace("'", "")
        try: 
            '''
            @给用例框赋值(非下拉列表)
            '''
            items=[]
            st=[]
            reportDate=time.strftime("%Y%m%d", time.localtime())
            fileData,email,userParams,userParamsValue,data,sheet,nrows,ncols,column=self.init(reportDate,path,file,sheetName)
            for i in range(3,nrows+1):
                st.append(str(i)+' '+str(getValue(file,sheet,i-1,column[23])))
                st.append(str(getValue(file,sheet,i-1,column[24])))
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
        if fname=='请选择文件' or fname=='':
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
                ex.task.setEnabled(True)
                ex.abort.setEnabled(True)
                ex.file.setEnabled(True)
                ex.debug.setEnabled(True)
                ex.dtailReport.setEnabled(True)
                ex.html.setEnabled(True)
                ex.dtailLog.setEnabled(True)
                ex.qSheetName.setEnabled(True)
                ex.refresh.setEnabled(True)
    
    '''
    @执行用例
    '''
    def start(self):
        fname=self.fileName.text()
        if fname=='请选择文件' or fname=='':
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
                ex.task.setEnabled(True)
                ex.abort.setEnabled(True)
                ex.file.setEnabled(True)
                ex.analyJSON.setEnabled(True)
                ex.dtailReport.setEnabled(True)
                ex.html.setEnabled(True)
                ex.dtailLog.setEnabled(True)
                ex.qSheetName.setEnabled(True)
                ex.refresh.setEnabled(True)

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
        try:
            scheduler.shutdown()
            self.task.setEnabled(True)
            self.checkMail.setCheckable(True)
            self.taskTime.setEnabled(True)
            ex.console.clear()
            ex.console.append("<font color=\"#000000\">"+'定时任务已取消'+"</font>")
        except Exception as e:
            print(e)
        
'''
@接口解析
'''        
class analyFunctionClass(QThread,DetailUI):
        
    def __init__(self):
        super(DetailUI, self).__init__()
    
    def run(self):
        try:
            ex.debug.setEnabled(False)
            ex.task.setEnabled(False)
            ex.abort.setEnabled(False)
            ex.file.setEnabled(False)
            ex.dtailReport.setEnabled(False)
            ex.html.setEnabled(False)
            ex.dtailLog.setEnabled(False)
            ex.qSheetName.setEnabled(False)
            ex.refresh.setEnabled(False)
            ex.successNum.setText('0')
            ex.failNum.setText('0')
            ex.skipNum.setText('0')
            ex.result.setText('0/0')
            exa=ex.example.currentText()
            if exa==[]:
                ex.console.append("<font color=\"#FF0000\">"+'请选择接口'+"</font> ")
                ex.console.append("<font color=\"#000000\"></font>")
            else:
                fileData,email,userParams,userParamsValue=ex.initConfig(path)
                exa=exa.replace("'", '').replace('(', '').replace(')','')
                exa=exa.split(',')
                for i in range(0,len(exa)):
                    ex.console.append("<font color=\"#000000\"></font>")
                    ex.analyFunc(file,int(exa[i])-1, sheetName,userParams,userParamsValue,sheet,fileRes,column)
        except Exception as e:
            print(e)
        ex.debug.setEnabled(True)
        ex.task.setEnabled(True)
        ex.abort.setEnabled(True)
        ex.file.setEnabled(True)
        ex.dtailReport.setEnabled(True)
        ex.html.setEnabled(True)
        ex.dtailLog.setEnabled(True)
        ex.qSheetName.setEnabled(True)
        ex.refresh.setEnabled(True)
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
            ex.successNum.setText('0')
            ex.failNum.setText('0')
            ex.skipNum.setText('0')
            ex.status1=0#success
            ex.status2=0#fail
            ex.status3=0#skip
            ex.allRows=0
            
        try:
            reportDate=time.strftime("%Y%m%d", time.localtime())
            model='普通' if ex.model1.isChecked() else '简洁'
            now_time = datetime.datetime.now()
            startTime=datetime.datetime.strptime(str(now_time)[:-7],'%Y-%m-%d %H:%M:%S') 
            ex.successNum.setText('0')
            ex.failNum.setText('0')
            ex.skipNum.setText('0')
            ex.console.append("<font color=\"#000000\"></font>")
            text=ex.result.text()
            ex.result.setText('0'+text[text.index('/'):])
            ex.task.setEnabled(False)
            ex.abort.setEnabled(False)
            ex.file.setEnabled(False)
            ex.analyJSON.setEnabled(False)
            ex.dtailReport.setEnabled(False)
            ex.html.setEnabled(False)
            ex.dtailLog.setEnabled(False)
            ex.qSheetName.setEnabled(False)
            ex.refresh.setEnabled(False)
            data = readExcel(path+'/'+file)
            fileData,email,userParams,userParamsValue=ex.initConfig(path)
            book,sheet1,fileRes=ex.createReport(reportDate,path, file, data, sheetNames)
            sheetValue=ex.qSheetName.currentText()
            allRpt=''
            testResult=[]
            '''
            @全量
            '''
            if sheetValue=='全部' and ex.qSheetName.currentIndex()==0:
                ex.console.append("<font color=\"#000000\"></font>")
                for i in range(0,len(sheetNames)):
                    fileData,email,userParams,userParamsValue,data,sheet,nrows,ncols,column=ex.init(reportDate,path,file,sheetNames[i])   
                    rpt=ex.verTemp(file,sheetNames[i],sheet,ncols,book,sheet1[i],fileRes,column)
                    '''
                    @模板校验通过
                    '''
                    if rpt=='':
                        noRuns=0
                        '''
                        @找出迭代次数为0（不执行）的用例
                        '''
                        for i in range(3,nrows+1):
                            if str(getValue(file,sheet,i-1, column[24])).upper()=='0':
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
                    for i in range(0,len(sheetNames)):
                        '''
                        @初始化页签
                        '''
                        fileData,email,userParams,userParamsValue,data,sheet,nrows,ncols,column=ex.init(reportDate,path,file,sheetNames[i])
                        '''
                        @执行该页签中的用例
                        '''
                        dict,tr=ex.run(file,model,'',sheetNames[i],userParams,userParamsValue,sheet,nrows,book,sheet1[i],fileRes,column,ex.allRows)
                        '''
                        @拼接结果集
                        '''
                        testResult=testResult+tr
            else: #单个或多个
                for i in range(0,len(sheetNames)):
                    if sheetValue==sheetNames[i]:
                        sheet1=sheet1[i]
                        break
                fileData,email,userParams,userParamsValue,data,sheet,nrows,ncols,column=ex.init(reportDate,path,file,sheetName)
                rpt=ex.verTemp(file,sheetName,sheet,ncols,book,sheet1,fileRes,column)
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
                            if str(getValue(file,sheet,i-1, column[24]))!='0':
                                exa.append(i)
                    else:
                        '''
                        @取选中的用例列表                        
                        '''
                        exa=exa.replace("'", '').replace('(', '').replace(')','').split(',')
                        '''
                        @是否存在大于当前页签行数的接口号(删除用例引起)
                        '''
                        en=[int(item) for item in exa if int(item)>nrows]
                    '''
                    @如果当前选中的用例(序列号大于nrows)已被删除,则提示
                    '''
                    if en != []:
                        ex.console.append("<font color=\"#FF0000\">"+'用例'+str(en)+"不存在"+"</font>")
                    else:
                        for item in exa:
                            if str(getValue(file,sheet,int(item)-1, column[24]))!='0':#不执行迭代次数为0的用例
                                dict,tr=ex.run(file,model,int(item),sheetName,userParams,userParamsValue,sheet,nrows,book,sheet1,fileRes,column,len(exa))
                                testResult=testResult+tr
            '''
            @格式化html报告中的运行时间和时长
            '''
            now_time = datetime.datetime.now()
            endTime=datetime.datetime.strptime(str(now_time)[:-7],'%Y-%m-%d %H:%M:%S') 
            duration=datetime.datetime.strptime(str(endTime-startTime),'%H:%M:%S')
            dd=str(duration)[11:].split(':')
            for item in dd:
                if item[0:1]=='0':
                    item=item[1:]
            duration=dd[0]+'小时 '+dd[1]+'分 '+dd[2]+'秒'
            taskName=file[:-4] if file[-4:]=='.xls' else file[:-5]
                    
            '''
            @测试结果存到字典中，用于html测试报告
            '''
            dict['testName']=taskName#项目名称
            dict['beginTime']=str(startTime)#开始时间
            dict['totalTime']=duration#运行时长
            dict['testResult']=testResult#结果集
            htmlReportName=ex.createHTMLReport(reportDate,dict,file,path)
            ex.sendEmail(htmlReportName)
            ex.status1=0#success
            ex.status2=0#fail
            ex.status3=0#skip
            ex.allRows=0
            ex.debug.setText('开始')
        except Exception as e:
            print(e)
            ex.status1=0#success
            ex.status2=0#fail
            ex.status3=0#skip
            ex.allRows=0
        ex.task.setEnabled(True)
        ex.abort.setEnabled(True)
        ex.file.setEnabled(True)
        ex.analyJSON.setEnabled(True)
        ex.dtailReport.setEnabled(True)
        ex.html.setEnabled(True)
        ex.dtailLog.setEnabled(True)
        ex.qSheetName.setEnabled(True) 
        ex.refresh.setEnabled(True)
        ex.debug.setText('开始')  
            
    
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
            if fname=='请选择文件' or fname=='':
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
                    ex.console.append("<font color=\"#FF0000\">"+'时间不得小于当前时间,请重新输入'+"</font> ")
                    ex.console.append("<font color=\"#000000\"></font>")
                if flag==True:
                    trigger = DateTrigger(inTime)
                    scheduler = BackgroundScheduler ()
                    scheduler.add_job(self.taskJob, trigger)
                    ex.console.append("<font color=\"#000000\">"+'定时任务将于'+str(inTime)+'执行:'+"</font>")
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
            if fname=='请选择文件' or fname=='':
                pass
            else:
                ex.console.append("<font color=\"#000000\">"+'定时任务开始:'+"</font>")
                now_time = datetime.datetime.now()
                startTime=datetime.datetime.strptime(str(now_time)[:-7],'%Y-%m-%d %H:%M:%S')
                ex.successNum.setText('0')
                ex.failNum.setText('0')
                ex.skipNum.setText('0')
                ex.result.setText('0/0')
                ex.abort.setEnabled(False)
                ex.debug.setEnabled(False)
                ex.file.setEnabled(False)
                ex.analyJSON.setEnabled(False)
                ex.dtailReport.setEnabled(False)
                ex.html.setEnabled(False)
                ex.dtailLog.setEnabled(False)
                ex.qSheetName.setEnabled(False)
                ex.refresh.setEnabled(False)
                data = readExcel(path+'/'+file)
                book,sheet1,fileRes=ex.createReport(reportDate,path, file, data, sheetNames)
                allRpt=''
                try:
                    '''
                    @定时任务只支持执行全部用例，
                    '''
                    model='普通' if ex.model1.isChecked() else '简洁'
                    for i in range(0,len(sheetNames)):
                        fileData,email,userParams,userParamsValue,data,sheet,nrows,ncols,column=ex.init(reportDate,path,file,sheetNames[i])   
                        rpt=ex.verTemp(file,sheetNames[i],sheet,ncols,book,sheet1[i],fileRes,column)
                        if rpt=='':
                            noRuns=0
                            IterationCol=findStr(file,sheet,ncols,'Iteration')
                            for i in range(3,nrows+1):
                                if str(getValue(file,sheet,i-1, IterationCol)).upper()=='0':
                                    noRuns=noRuns+1
                            ex.allRows=ex.allRows+nrows-2-noRuns
                        allRpt=allRpt+str(rpt)
                    '''
                    @所有页签的模板校验通过
                    '''
                    if allRpt=='':
                        testResult=[]
                        for i in range(0,len(sheetNames)):
                            fileData,email,userParams,userParamsValue,data,sheet,nrows,ncols,column=self.init(reportDate,path,file,sheetNames[i])
                            dict,tr=ex.run(file,model,'',sheetNames[i],userParams,userParamsValue,sheet,nrows,book,sheet1[i],fileRes,column,ex.allRows)
                            testResult=testResult+tr
                except Exception as e:
                    print(e)
                    ex.console.append("<font color=\"#FF0000\">"+'定时任务执行失败'+"</font> ")
                    ex.console.append("<font color=\"#000000\"></font>")
                    
                now_time = datetime.datetime.now()
                endTime=datetime.datetime.strptime(str(now_time)[:-7],'%Y-%m-%d %H:%M:%S') 
                duration=datetime.datetime.strptime(str(endTime-startTime),'%H:%M:%S')
                dd=str(duration)[11:].split(':')
                for i in range(0,len(dd)):
                    if dd[i][0:1]==str(0):
                        dd[i]=dd[i][1:]
                duration=dd[0]+'小时 '+dd[1]+'分 '+dd[2]+'秒' 
                taskName=file[:-4] if file[-4:]=='.xls' else file[:-5]
                
                '''
                @测试结果存到字典中，用于html测试报告
                '''
                dict['testName']=taskName#项目名称
                dict['beginTime']=str(startTime)#开始时间
                dict['totalTime']=duration#运行时长
                dict['testResult']=testResult#结果集
                htmlReportName=ex.createHTMLReport(reportDate,dict,file,path)
                ex.sendEmail(htmlReportName)
                ex.console.append("<font color=\"#000000\"></font>")
                ex.console.append("<font color=\"#000000\">"+'定时任务执行成功'+"</font>")
                scheduler.shutdown()
                ex.status1=0#success
                ex.status2=0#fail
                ex.status3=0#skip
                ex.allRows=0
        except Exception as e:
            print(e)
            ex.status1=0#success
            ex.status2=0#fail
            ex.status3=0#skip
            ex.allRows=0
        ex.task.setEnabled(True)
        ex.abort.setEnabled(True)
        ex.taskTime.setEnabled(True)
        ex.debug.setEnabled(True)
        ex.file.setEnabled(True)
        ex.analyJSON.setEnabled(True)
        ex.dtailReport.setEnabled(True)
        ex.html.setEnabled(True)
        ex.dtailLog.setEnabled(True)
        ex.qSheetName.setEnabled(True)
        ex.refresh.setEnabled(True)
            
if __name__ == "__main__":
    app=0
    app = QApplication(sys.argv)    
    ex = DetailUI()
    ex.show()
    sys.exit(app.exec_())
