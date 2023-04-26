import cgitb
import datetime
import os
import sys
import time
import traceback

from PyQt5 import QtGui, QtCore
from PyQt5.QtCore import QThread
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QListView

from common.excel.Report import Report
from common.excel.Template import Template
from common.excel.Write import Write
from common.ui.ExampleBox import ExampleBox
from common.ui.MainWindow import Ui_MainWindow

'''
#主类
#author: dujianxiao
'''
date = time.strftime("%Y%m%d")
sheetNames = fname = bookRes = sheetRes = fileRes = path = file = sheetName = sheet = nrows = ""


class DetailUI(Ui_MainWindow, QMainWindow, Write, Report, Template):

    def resource_path(self, relative_path):
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.example = ExampleBox()
        self.example.setMinimumSize(QtCore.QSize(0, 25))
        self.example.setObjectName("example")
        self.gridLayout.addWidget(self.example, 1, 1, 1, 1)
        self.qSheetName.setView(QListView())

        self.file.clicked.connect(self.getFile)
        self.fileName.clicked.connect(self.openExample)
        self.qSheetName.currentIndexChanged['int'].connect(self.changeSheet)
        self.refresh.clicked.connect(self.reloadSheet)
        self.example.popupAboutToBeShown.connect(self.reload)
        self.dtailReport.clicked.connect(self.openExcelReport)
        self.html.clicked.connect(self.openHtmlReport)
        self.dtailLog.clicked.connect(self.openLog)
        self.debug.clicked.connect(self.start)
        self.analyJSON.clicked.connect(self.analyFunction)
        self.example.currentTextChanged.connect(self.changeExample)

        self.setWindowTitle('时默')
        img = self.resource_path(os.path.join(".", "source/1.ico"))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(img), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.setWindowIcon(icon)

        self.allRows = 0
        self.status1 = 0  # success
        self.status2 = 0  # fail
        self.status3 = 0  # skip

    def getFile(self):
        """
        打开用例文件
        """
        self.console.clear()
        global sheetNames, fname, bookRes, sheetRes, fileRes, path, file
        try:
            # 加载文件时清除上一个文件可能存在的请求头信息、接口变量和用户变量
            self.headerManager = ''
            self.interData = {}
            self.fileData = []
            self.example.clear()
            self.result.setText('0/0')
            self.initTextNum()

            # 以下8行是jenkins.py的代码，再往下4行是sm.py的代码
            # path = os.getcwd()
            # path = path.replace('\\', '/')
            # fileData = self.initConfig(path)
            # self.initLog(path)
            # file = str(fileData[-1][1])  # 需要在conf.ini最后一行写入用例文件的名称如jenkinsFile=ems.xls
            # fname = f'{path}/{file}'
            # sheetNames = self.getSheetNames(fname)
            # self.fileName.setText(fname)

            # 以下4行是sm.py的代码，以上8行是jenkins.py的代码
            fname, _ = QFileDialog.getOpenFileName(self, 'open file', '/', "files (*.xls *.xlsx)")
            self.fileName.setToolTip(fname)
            self.fileName.setText(fname)
            path, file = self.getPath(fname)
            # 如果未选择文件，页签下拉列表置空
            if fname in ['请选择文件', '']:
                self.qSheetName.clear()
                self.console.clear()
            else:
                # 初始化日志和配置文件
                self.initConfig(path)
                self.initLog(path)
                # 创建用例结果文件
                sheetNames = self.getSheetNames(f'{path}/{file}')
                bookRes, sheetRes, fileRes = self.createReport(date, path, file, sheetNames)
                self.qSheetName.clear()
                self.example.clear()
                # 填充页签下拉列表
                for i in range(len(sheetNames) + 1):
                    if i == 0:
                        self.qSheetName.addItem('全部')
                    else:
                        self.qSheetName.addItem(str(sheetNames[i - 1]))
        except Exception as e:
            print(e)
            self.qSheetName.clear()

    def openExample(self):
        """
        点击文件名打开文件
        """
        try:
            fname = ex.fileName.text()
            if fname in ['请选择文件', '']:
                ex.console.clear()
            else:
                os.startfile(eval(f"r'{fname}'"))
        except Exception as e:
            print(e)
            ex.console.clear()

    def changeSheet(self):
        """
        切换页签
        """
        global sheetName, sheet, nrows
        self.initTextNum()
        self.console.clear()
        self.example.clear()
        try:
            sheetName = self.qSheetName.currentText()
            bookRes, sheetRes, fileRes = self.createReport(date, path, file, sheetNames)
            if (sheetName == '全部' and self.qSheetName.currentIndex() == 0) or sheetName == '':
                self.example.setCurrentText('')
                for i in range(len(sheetNames)):
                    self.example.clear
                    self.example.items.clear()
                self.example.loadItems([])
                allRows = 0
                allRpt = ''
                # 每次切换页签时都校验一遍模板，防止使用过程中对模板有改动
                for i in range(len(sheetNames)):
                    sheet, nrows = self.initFile(date, path, file, sheetNames[i])
                    rpt = ex.verTemp(sheetNames[i], sheet, bookRes, sheetRes[i], fileRes)
                    allRpt = allRpt + str(rpt)
                    if rpt == '':
                        noRuns = 0
                        IterationCol = self.findStr(file, sheet, 'Iteration')
                        for i in range(3, nrows + 1):
                            if str(self.getValue(file, sheet, i - 1, IterationCol)) == '0':
                                noRuns = noRuns + 1
                        allRows = allRows + nrows - 2 - noRuns
                if allRpt == '':
                    ex.result.setText(f'0/{allRows}')
            else:
                items = []
                st = []
                # 每次切换页签时都校验一遍模板，防止使用过程中对模板有改动
                sheet, nrows = self.initFile(date, path, file, sheetName)
                rpt = ex.verTemp(sheetName, sheet, bookRes, sheetRes[0], fileRes)
                noRuns = 0
                if rpt == '':
                    for i in range(3, nrows + 1):
                        st.append(f'{i} {self.getValue(file, sheet, i - 1, ex.nameCol)}')
                        st.append(str(self.getValue(file, sheet, i - 1, ex.IterationCol)))
                        items.append(st)
                        st = []
                    self.example.loadItems(items)
                    IterationCol = self.findStr(file, sheet, 'Iteration')
                    for i in range(3, nrows + 1):
                        if str(self.getValue(file, sheet, i - 1, IterationCol)).upper() == '0':
                            noRuns = noRuns + 1
                    ex.result.setText(f'0/{(nrows - 2 - noRuns)}')
                else:
                    ex.console.clear()
                    ex.consoleFunc('red', str(rpt))
        except Exception as e:
            print(e)

    def reloadSheet(self):
        """
        每次点击页签下拉框重新加载下拉列表
        考虑到性能问题,这个方法没有被调用
        应同事要求，在界面加了一个刷新按钮调用此方法
        """
        try:
            # 刷新时清除上一个文件可能存在的请求头信息、接口变量和用户变量
            self.headerManager = ''
            self.interData = {}
            self.fileData = []
            fname = ex.fileName.text()
            if fname in ['请选择文件', '']:
                pass
            else:
                self.qSheetName.clear()
                sheetNames = self.getSheetNames(f"{path}/{file}")
                # 填充页签下拉列表
                for i in range(len(sheetNames) + 1):
                    if i == 0:
                        self.qSheetName.addItem('全部')
                    else:
                        self.qSheetName.addItem(str(sheetNames[i - 1]))
                self.qSheetName.setCurrentIndex(0)
        except Exception as e:
            print(e)

    def reload(self):
        """
        点击用例下拉框时重新填充下拉列表
        """
        try:
            sheetName = self.qSheetName.currentText()
            if (sheetName == '全部' and ex.qSheetName.currentIndex() == 0) or sheetName == '':
                self.example.setCurrentText('')
                for i in range(len(sheetNames)):
                    self.example.clear
                    self.example.items.clear()
                allRows = 0
                allRpt = ''
                # 每次切换页签时都校验一遍模板，防止使用过程中对模板有改动
                for i in range(len(sheetNames)):
                    sheet, nrows = self.initFile(date, path, file, sheetNames[i])
                    rpt = ex.verTemp(sheetNames[i], sheet, bookRes, sheetRes[i], fileRes)
                    allRpt = allRpt + str(rpt)
                    if rpt == '':
                        noRuns = 0
                        IterationCol = self.findStr(file, sheet, 'Iteration')
                        for i in range(3, nrows + 1):
                            if str(self.getValue(file, sheet, i - 1, IterationCol)).upper() == '0':
                                noRuns = noRuns + 1
                        allRows = allRows + nrows - 2 - noRuns
                if allRpt == '':
                    ex.result.setText(f'0/{allRows}')
            else:
                # 每次切换页签时都校验一遍模板，防止使用过程中对模板有改动
                items = []
                st = []
                sheet, nrows = self.initFile(date, path, file, sheetName)
                rpt = ex.verTemp(sheetName, sheet, bookRes, sheetRes, fileRes)
                # 模板校验通过
                if rpt == '':
                    # 获取被选中的用例
                    exa = str(ex.example.currentText())[1:-1].replace("'", '').split(",")
                    for value in exa:
                        if value == '':
                            exa.remove(value)
                    exa = [int(i) for i in exa if i.isdigit()]
                    self.example.clear()
                    self.example.items.clear()
                    noRuns = 0
                    IterationCol = self.findStr(file, sheet, 'Iteration')
                    for i in range(3, nrows + 1):
                        if str(self.getValue(file, sheet, i - 1, IterationCol)).upper() == '0':
                            noRuns = noRuns + 1
                    for i in range(3, nrows + 1):
                        st.append(str(i) + ' ' + str(self.getValue(file, sheet, i - 1, ex.nameCol)))
                        st.append(str(self.getValue(file, sheet, i - 1, ex.IterationCol)))
                        items.append(st)
                        st = []
                    self.example.loadItems(items)
                    # 保持上一次的选中状态
                    if exa:
                        for i in range(len(exa)):
                            try:
                                self.example.qCheckBox[exa[i] - 2].setChecked(True)
                            except Exception as e:
                                print(e)
                    ex.result.setText(f"0/{nrows - 2 - noRuns}")
                else:
                    ex.console.clear()
                    ex.consoleFunc('red', str(rpt))
        except Exception as e:
            print(e)

    def changeExample(self):
        """
        切换用例
        """
        try:
            exa = self.example.Selectlist()
            ex.result.setText(f"0/{len(exa)}")
        except Exception as e:
            print(e)

    def openExcelReport(self):
        """
        打开excel报告
        """
        try:
            fname = ex.fileName.text()
            if fname in ['请选择文件', '']:
                ex.console.clear()
            else:
                reportName = file[:file.index('.xls')] + '-' + date + '-report.xls'
                if file.endswith('xlsx'):
                    reportName += 'x'
                excel = f"r'{path}/result/{reportName}'"
                os.startfile(eval(excel))
        except Exception as e:
            print(e)
            ex.console.clear()
            ex.consoleFunc('red', '打开报告失败')

    def createHTMLReport(self, js):
        """
        创建html测试报告
        @param path:
        @param file:
        @param js: json格式的测试结果
        """
        try:
            try:
                html = self.resource_path("source/template")
                f1 = open(html, "r", encoding="utf-8")
                htmlData = f1.read()
                html = htmlData.replace('${resultData}', str(js))
                f1.close()
                file_name = file[:file.index('.xls')]
                html_file = f"{path}/result/{file_name}-{date}-report.html"
                if os.path.exists(html_file):
                    os.remove(html_file)
                htmlReportName = f"{path}/result/{file_name}-{date}-report.html"
                f2 = open(htmlReportName, 'w', encoding='utf-8')
                f2.write(html)
                f2.close()
            except:
                print(traceback.format_exc())
            return htmlReportName
        except Exception as e:
            print(e)

    def openHtmlReport(self):
        """
        打开html报告
        """
        try:
            fname = ex.fileName.text()
            if fname in ['请选择文件', '']:
                ex.console.clear()
            else:
                file_name = file[:file.index('.xls')]
                reportName = f"{file_name}-{date}-report.html"
                html_path = f"r'{path}/result/{reportName}'"
                os.startfile(eval(html_path))
        except Exception as e:
            print(e)
            ex.console.clear()
            ex.consoleFunc('red', '打开报告失败')

    def openLog(self):
        """
        打开日志
        """
        try:
            fname = ex.fileName.text()
            if fname in ['请选择文件', '']:
                ex.console.clear()
            else:
                log_path = f"r'{path}/result/info.log'"
                os.startfile(eval(log_path))
        except Exception as e:
            print(e)
            ex.console.clear()
            ex.consoleFunc('red', '打开日志失败')

    def getPath(self, path: str):
        """
        获取文件路径
        """
        index = path.rfind("/")
        file_path = path[:index]
        file_name = path[index + 1:]
        return file_path, file_name

    def analyFunction(self):
        """
        接口解析
        """
        fname = ex.fileName.text()
        if fname in ['请选择文件', '']:
            pass
        else:
            ss = self.analyJSON.text()
            if ss == '解  析':
                ex.console.clear()
                ex.analy_thread = analyFunctionClass()
                ex.analy_thread.start()
                ex.analyJSON.setText('停  止')
            elif ss == '停  止':
                ex.analyJSON.setText('解  析')
                ex.analy_thread.terminate()
                self.buttonStatus(True)

    def start(self):
        """
        执行用例
        """
        fname = ex.fileName.text()
        if fname in ['请选择文件', '']:
            pass
        else:
            ss = self.debug.text()
            if ss == '开  始':
                ex.console.clear()
                ex.debug_thread = debugClass()
                ex.debug_thread.start()
                ex.debug.setText('停  止')
            elif ss == '停  止':
                ex.debug.setText('开  始')
                ex.debug_thread.terminate()
                self.buttonStatus(True)

    def buttonStatus(self, flag):
        ex.debug.setEnabled(flag)
        ex.file.setEnabled(flag)
        ex.analyJSON.setEnabled(flag)
        ex.dtailReport.setEnabled(flag)
        ex.html.setEnabled(flag)
        ex.dtailLog.setEnabled(flag)
        ex.qSheetName.setEnabled(flag)
        ex.refresh.setEnabled(flag)

    def initTextNum(self):
        ex.successNum.setText('0')
        ex.failNum.setText('0')
        ex.skipNum.setText('0')


"""
接口解析
"""


class analyFunctionClass(QThread, DetailUI):

    def __init__(self):
        super(DetailUI, self).__init__()

    def run(self):
        try:
            ex.buttonStatus(False)
            ex.analyJSON.setEnabled(True)
            self.initTextNum()
            ex.result.setText('0/0')
            exa = ex.example.currentText()
            if not exa:
                ex.consoleFunc('red', '请选择接口')
            else:
                exa = exa[1:-1].replace("'", '').split(',')
                exa = [int(item) for item in exa]
                ex.initConfig(path)
                for i in range(len(exa)):
                    ex.analyFunc(file, exa[i] - 1, sheetName, sheet)
        except Exception as e:
            print(e)
        self.buttonStatus(True)
        ex.analyJSON.setText('解  析')


class debugClass(QThread, DetailUI):

    def __init__(self):
        super(DetailUI, self).__init__()

    def run(self):
        dict = {}
        ex.runFlag = True
        cText = ex.debug.text()
        if cText == '停  止':
            self.initTextNum()
            ex.status1 = 0  # success
            ex.status2 = 0  # fail
            ex.status3 = 0  # skip
            ex.allRows = 0
        try:
            model = '普  通' if ex.model1.isChecked() else '简  洁'
            startTime = datetime.datetime.now()
            self.initTextNum()
            text = ex.result.text()
            ex.result.setText('0' + text[text.index('/'):])
            ex.buttonStatus(False)
            ex.debug.setEnabled(True)
            ex.initConfig(path)
            bookRes, sheetRes, fileRes = ex.createReport(date, path, file, sheetNames)
            sheetValue = ex.qSheetName.currentText()
            allRpt = ''
            testResult = []
            # 全量
            if sheetValue == '全部' and ex.qSheetName.currentIndex() == 0:
                for i in range(len(sheetNames)):
                    sheet, nrows = ex.initFile(date, path, file, sheetNames[i])
                    rpt = ex.verTemp(sheetNames[i], sheet, bookRes, sheetRes[i], fileRes)
                    # 模板校验通过
                    if rpt == '':
                        noRuns = 0
                        # 找出迭代次数为0（不执行）的用例
                        for i in range(3, nrows + 1):
                            if str(self.getValue(file, sheet, i - 1, ex.IterationCol)).upper() == '0':
                                noRuns = noRuns + 1
                        # 全部用例数为各页签的用例数相加减去不执行的用例数
                        ex.allRows = ex.allRows + nrows - 2 - noRuns  # 全部用例数
                    allRpt = allRpt + str(rpt)
                # 模板全部校验通过
                if allRpt == '':
                    for i in range(len(sheetNames)):
                        sheet, nrows = ex.initFile(date, path, file, sheetNames[i])
                        # 执行该页签中的用例
                        dict, tr = ex.run(model, '', sheetNames[i], sheet, nrows, bookRes, sheetRes[i], fileRes,
                                          ex.allRows)
                        testResult = testResult + tr
            else:  # 单个或多个
                for i in range(len(sheetNames)):
                    if sheetValue == sheetNames[i]:
                        sheetRes = sheetRes[i]
                        break
                sheet, nrows = ex.initFile(date, path, file, sheetName)
                rpt = ex.verTemp(sheetName, sheet, bookRes, sheetRes, fileRes)
                # 模板校验通过才进行之后的操作
                if rpt == "":
                    exa = ex.example.currentText()
                    en = []
                    # 未选中任何用例则执行该页签的全部用例
                    if not exa:
                        for i in range(3, nrows + 1):
                            # 迭代次数为0表示此用例不执行
                            if str(self.getValue(file, sheet, i - 1, ex.IterationCol)) != '0':
                                exa.append(i)
                    else:
                        # 是否存在大于当前页签行数的接口号(删除用例引起)
                        exa = exa[1:-1].replace("'", '').split(',')
                        exa = [int(item) for item in exa]
                        en = [item for item in exa if item > nrows]
                    # 如果当前选中的用例(序列号大于nrows)已被删除,则提示
                    if en:
                        ex.consoleFunc('red', f"用例{en}不存在")
                    else:
                        for item in exa:
                            if str(self.getValue(file, sheet, item - 1, ex.IterationCol)) != '0':  # 不执行迭代次数为0的用例
                                dict, tr = ex.run(model, item, sheetName, sheet, nrows, bookRes, sheetRes, fileRes,
                                                  len(exa))
                                testResult = testResult + tr
            # 格式化html报告中的运行时间和时长
            endTime = datetime.datetime.now()
            second = str(endTime - startTime)
            duration = second[:second.index('.')]
            dd = duration.split(':')
            duration = f"{dd[0]}小时 {dd[1]}分 {dd[2]}秒"
            taskName = file[:file.index(".")]
            # 测试结果存到字典中，用于html测试报告
            dict['testName'] = taskName  # 项目名称
            startTime = str(startTime)
            dict['beginTime'] = startTime[:startTime.index('.')]  # 开始时间
            dict['totalTime'] = duration  # 运行时长
            dict['testResult'] = testResult  # 结果集
            ex.createHTMLReport(dict)
        except Exception as e:
            print(e)
        ex.status1 = 0  # success
        ex.status2 = 0  # fail
        ex.status3 = 0  # skip
        ex.allRows = 0
        self.buttonStatus(True)
        ex.debug.setText('开  始')
        # 此行代码用于集成jenkins时，当用例执行完毕后自动退出程序
        # sys.exit()


if __name__ == "__main__":
    log_dir = os.path.join(os.getcwd(), 'log')
    if not os.path.exists(log_dir):
        os.mkdir(log_dir)
    cgitb.enable(format='text', logdir=log_dir)

    app = 0
    app = QApplication(sys.argv)  #
    QssStyle1 = '''
                    QPushButton:hover{
                    color: rgb(0,51,153);
                    font-weight:bold;
                    transition-duration: 0.3s;
                    -webkit-transition-duration: 0.3s;
                    }
                    '''

    app.setStyleSheet(QssStyle1)
    ex = DetailUI()
    ex.show()
    # jenkins需要加以下3行
    # ex.model2.click()
    # ex.getFile()
    # ex.start()
    sys.exit(app.exec_())
