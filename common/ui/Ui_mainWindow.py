from PyQt5 import QtCore, QtGui, QtWidgets
from common.ui.ComboCheckBox import ComboCheckBox
from PyQt5.QtWidgets import QListView
import sys
import os
from common.ui.TextEdit import TextEdit

'''
@author: dujianxiao
'''
class Ui_mainWindow(object):
    def resource_path(self,relative_path):
        base_path = getattr(sys,'_MEIPASS',os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)
    
    def setupUi(self, mainWindow):
        mainWindow.setObjectName("mainWindow")
        mainWindow.resize(650, 600)
        mainWindow.setIconSize(QtCore.QSize(12, 24))
        self.centralwidget = QtWidgets.QWidget(mainWindow)
        mainWindow.setMinimumSize(650, 600)
#         mainWindow.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)  窗口置顶
        self.example=ComboCheckBox(mainWindow)
        self.console=TextEdit(mainWindow)
        self.setCentralWidget(self.centralwidget)
        self.centralwidget.setMouseTracking(True)
        self.setMouseTracking(True)
        font = QtGui.QFont()
        font.setPointSize(9)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        font.setKerning(True)
        font.setStyleStrategy(QtGui.QFont.PreferDefault)
        self.centralwidget.setFont(font)
        self.centralwidget.setAutoFillBackground(False)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        
        self.fileName = QtWidgets.QPushButton(self.centralwidget)
        self.fileName.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.fileName.setAutoRepeatDelay(300)
        self.fileName.setObjectName("fileName")
        self.gridLayout.addWidget(self.fileName, 0, 0, 1, 2)
        
        self.file = QtWidgets.QPushButton(self.centralwidget)
        self.file.setObjectName("file")
        self.gridLayout.addWidget(self.file, 0, 2, 1, 1)
        
        self.refresh = QtWidgets.QPushButton(self.centralwidget)
        self.refresh.setObjectName("file")
        self.gridLayout.addWidget(self.refresh, 0, 3, 1, 1)
        
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setLineWidth(0)
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 0, 4, 1, 1)
        
        self.result = QtWidgets.QLabel(self.centralwidget)
        self.result.setLineWidth(0)
        self.result.setAlignment(QtCore.Qt.AlignCenter)
        self.result.setObjectName("result")
        self.gridLayout.addWidget(self.result, 0, 5, 1, 1)
        
        self.qSheetName = QtWidgets.QComboBox(self.centralwidget)
        self.qSheetName.setWhatsThis("")
        self.qSheetName.setObjectName("qSheetName")
        self.qSheetName.setFixedHeight(22)
        self.qSheetName.setView(QListView())
        self.qSheetName.setStyleSheet("QAbstractItemView::item {height: 18px;} QScrollBar::vertical{width:0px;background:rgb(186,211,218);border:none;border-radius:5px;}")
        self.gridLayout.addWidget(self.qSheetName, 1, 0, 1, 1)       
        
        '''
        @用例下拉框加入窗口
        '''
        self.example.setWhatsThis("")
        self.example.setObjectName("example")
        self.example.setFixedHeight(22)
        self.gridLayout.addWidget(self.example, 1, 1, 1, 1)        
        
        self.debug = QtWidgets.QPushButton(self.centralwidget)
        self.debug.setObjectName("debug")
        self.gridLayout.addWidget(self.debug, 1, 2, 1, 1)
        
        self.analyJSON = QtWidgets.QPushButton(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.analyJSON.sizePolicy().hasHeightForWidth())
        self.analyJSON.setSizePolicy(sizePolicy)
        self.analyJSON.setIconSize(QtCore.QSize(8, 16))
        self.analyJSON.setObjectName("analyJSON")
        self.gridLayout.addWidget(self.analyJSON, 1, 3, 1, 1)
        
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setAlignment(QtCore.Qt.AlignCenter)
        self.label_4.setObjectName("label_4")
        self.gridLayout.addWidget(self.label_4, 1, 4, 1, 1)
        
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setAlignment(QtCore.Qt.AlignCenter)
        self.label_5.setObjectName("label_5")
        self.gridLayout.addWidget(self.label_5, 1, 5, 1, 1)
        
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setAlignment(QtCore.Qt.AlignCenter)
        self.label_6.setObjectName("label_6")
        self.gridLayout.addWidget(self.label_6, 1, 6, 1, 1)
        
        self.checkMail = QtWidgets.QCheckBox(self.centralwidget)
        self.checkMail.setText("发送邮件")
        self.checkMail.setCheckable(True)
        self.checkMail.setChecked(False)
        self.checkMail.setObjectName("checkMail")
        self.gridLayout.addWidget(self.checkMail, 2, 1, 1, 1 )
        
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 2, 0, 1, 1)
        
        
        self.task = QtWidgets.QPushButton(self.centralwidget)
        self.task.setObjectName("task")
        self.gridLayout.addWidget(self.task, 2, 2, 1, 1)
        
        self.abort = QtWidgets.QPushButton(self.centralwidget)
        self.abort.setObjectName("abort")
        self.gridLayout.addWidget(self.abort, 2, 3, 1, 1)
        
        self.successNum = QtWidgets.QLabel(self.centralwidget)
        self.successNum.setAlignment(QtCore.Qt.AlignCenter)
        self.successNum.setObjectName("successNum")
        self.gridLayout.addWidget(self.successNum, 2, 4, 1, 1)
        
        self.failNum = QtWidgets.QLabel(self.centralwidget)
        self.failNum.setAlignment(QtCore.Qt.AlignCenter)
        self.failNum.setObjectName("failNum")
        self.gridLayout.addWidget(self.failNum, 2, 5, 1, 1)
        
        self.skipNum = QtWidgets.QLabel(self.centralwidget)
        self.skipNum.setAlignment(QtCore.Qt.AlignCenter)
        self.skipNum.setObjectName("skipNum")
        self.gridLayout.addWidget(self.skipNum, 2, 6, 1, 1)
        
        self.model1 = QtWidgets.QRadioButton(self.centralwidget)
        self.model1.setFixedHeight(22)
        self.model1.setText('普通')
        self.model1.setChecked(True)
        self.gridLayout.addWidget(self.model1, 3, 0, 1, 1)
        
        self.model2 = QtWidgets.QRadioButton(self.centralwidget)
        self.model2.setFixedHeight(22)
        self.model2.setText('简洁')
        self.gridLayout.addWidget(self.model2, 3, 1, 1, 1)
        
        self.taskTime = QtWidgets.QDateTimeEdit(self.centralwidget)
        self.taskTime.setAlignment(QtCore.Qt.AlignCenter)
        self.taskTime.setReadOnly(False)
        self.taskTime.setMaximumDateTime(QtCore.QDateTime(QtCore.QDate(2099, 12, 31), QtCore.QTime(23, 59, 59)))
        self.taskTime.setMinimumDateTime(QtCore.QDateTime(QtCore.QDate(2020, 4, 14), QtCore.QTime(0, 0, 0)))
        self.taskTime.setMinimumDate(QtCore.QDate(2020, 4, 14))
        self.taskTime.setCurrentSection(QtWidgets.QDateTimeEdit.YearSection)
        self.taskTime.setObjectName("taskTime")
        self.gridLayout.addWidget(self.taskTime, 3, 2, 1, 2)
        
        self.dtailReport = QtWidgets.QPushButton(self.centralwidget)
        self.dtailReport.setObjectName("dtailReport")
        self.gridLayout.addWidget(self.dtailReport, 3, 4, 1, 1)
        
        self.html = QtWidgets.QPushButton(self.centralwidget)
        self.html.setObjectName("html")
        self.gridLayout.addWidget(self.html, 3, 5, 1, 1)
        
        self.dtailLog = QtWidgets.QPushButton(self.centralwidget)
        self.dtailLog.setObjectName("dtailLog")
        self.gridLayout.addWidget(self.dtailLog, 3, 6, 1, 1)
        
        
        '''
        @控制台加入窗口
        '''
        self.gridLayout.addWidget(self.console, 4, 0, 2000, 7)
        
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.gridLayout.addWidget(self.frame, 5, 2, 1, 5)
        
        mainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(mainWindow)
        self.statusbar.setObjectName("statusbar")
        mainWindow.setStatusBar(self.statusbar)
        self.menuBar = QtWidgets.QMenuBar(mainWindow)
        self.menuBar.setGeometry(QtCore.QRect(0, 0, 639, 23))
        self.menuBar.setObjectName("menuBar")
        mainWindow.setMenuBar(self.menuBar)
        
        '''
        @主窗口背景图
        '''
        img2=self.resource_path("source/2.png")
        img2=img2.replace('\\','/')
        img2="background-image: url("+img2+");"
        mainWindow.setStyleSheet(img2)  
        '''
        @控制台背景图
        '''
        img1=self.resource_path("source/1.png")
        img1=img1.replace('\\','/')
        img1="background-image: url("+img1+");"
        self.console.setStyleSheet(img1)

        self.successNum.setStyleSheet("font: 12pt 'Arial';color: green;")
        self.failNum.setStyleSheet("font: 12pt 'Arial';color: rgb(255,0,0);")
        self.skipNum.setStyleSheet("font: 12pt 'Arial';color: rgb(248,172,89);")
        self.result.setStyleSheet("font: 12pt 'Arial';")
        
        self.retranslateUi(mainWindow)
        self.file.clicked.connect(mainWindow.getFile)
        self.fileName.clicked.connect(mainWindow.openExample)
        self.qSheetName.currentIndexChanged['int'].connect(mainWindow.changeSheet)
        self.refresh.clicked.connect(mainWindow.reloadSheet)
        self.example.popupAboutToBeShown.connect(mainWindow.reload)
        self.dtailReport.clicked.connect(mainWindow.openExcelReport)
        self.html.clicked.connect(mainWindow.openHtmlReport)
        self.dtailLog.clicked.connect(mainWindow.openLog)
        self.debug.clicked.connect(mainWindow.start)
        self.task.clicked.connect(mainWindow.startTask)
        self.analyJSON.clicked.connect(mainWindow.analyFunction)
        self.abort.clicked.connect(mainWindow.abortTask)  
        self.example.currentTextChanged.connect(mainWindow.changeExample)
        QtCore.QMetaObject.connectSlotsByName(mainWindow)

    def retranslateUi(self, mainWindow):
        _translate = QtCore.QCoreApplication.translate
        mainWindow.setWindowTitle(_translate("mainWindow", "MainWindow"))
        self.failNum.setText(_translate("mainWindow", "0"))
        self.dtailReport.setText(_translate("mainWindow", "excel报告"))
        self.html.setText(_translate("mainWindow", "html报告"))
        self.debug.setText(_translate("mainWindow", "开始"))
        self.label.setText(_translate("mainWindow", "定时任务:"))
        self.label_5.setText(_translate("mainWindow", "失败"))
        self.label_4.setText(_translate("mainWindow", "成功"))
        self.task.setText(_translate("mainWindow", "执行"))
        self.abort.setText(_translate("mainWindow", "取消"))
        self.skipNum.setText(_translate("mainWindow", "0"))
        self.analyJSON.setText(_translate("mainWindow", "解析JSON"))
        self.taskTime.setDisplayFormat(_translate("mainWindow", "yyyy-MM-dd HH:mm:ss"))
        self.successNum.setText(_translate("mainWindow", "0"))
        self.label_2.setText(_translate("mainWindow", "结果预览:"))
        self.result.setText(_translate("mainWindow", "0/0"))
        self.dtailLog.setText(_translate("mainWindow", "查看日志"))
        self.label_6.setText(_translate("mainWindow", "异常"))
        self.fileName.setText(_translate("mainWindow", "请选择文件"))
        self.file.setText(_translate("mainWindow", "选择"))
        self.refresh.setText(_translate("mainWindow", "刷新"))
        
    