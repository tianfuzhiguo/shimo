# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'MainWindow.ui'
#
# Created by: PyQt5 UI code generator 5.15.9
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(880, 600)
        MainWindow.setMinimumSize(QtCore.QSize(880, 600))
        MainWindow.setStyleSheet("background:rgb(186,211,218);border:none;")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_4 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_4.setObjectName("gridLayout_4")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setObjectName("gridLayout")
        self.qSheetName = QtWidgets.QComboBox(self.centralwidget)
        self.qSheetName.setMinimumSize(QtCore.QSize(0, 25))
        self.qSheetName.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.qSheetName.setStyleSheet("background-color:#ffffff;")
        self.qSheetName.setCurrentText("")
        self.qSheetName.setObjectName("qSheetName")
        self.gridLayout.addWidget(self.qSheetName, 1, 0, 1, 1)
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setStyleSheet("background-color:#ffffff;")
        self.label.setText("")
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 1, 1, 1, 1)
        self.fileName = QtWidgets.QPushButton(self.centralwidget)
        self.fileName.setMinimumSize(QtCore.QSize(0, 25))
        self.fileName.setStyleSheet(";background-color:#ffffff;")
        self.fileName.setObjectName("fileName")
        self.gridLayout.addWidget(self.fileName, 0, 0, 1, 2)
        self.model1 = QtWidgets.QRadioButton(self.centralwidget)
        self.model1.setMinimumSize(QtCore.QSize(0, 25))
        self.model1.setStyleSheet("background-color:#ffffff;")
        self.model1.setChecked(True)
        self.model1.setObjectName("model1")
        self.gridLayout.addWidget(self.model1, 5, 0, 1, 1)
        self.model2 = QtWidgets.QRadioButton(self.centralwidget)
        self.model2.setMinimumSize(QtCore.QSize(0, 25))
        self.model2.setStyleSheet("background-color:#ffffff;")
        self.model2.setObjectName("model2")
        self.gridLayout.addWidget(self.model2, 5, 1, 1, 1)
        self.horizontalLayout.addLayout(self.gridLayout)
        self.gridLayout_2 = QtWidgets.QGridLayout()
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.html = QtWidgets.QPushButton(self.centralwidget)
        self.html.setMinimumSize(QtCore.QSize(0, 25))
        self.html.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.html.setStyleSheet("    color: white;\n"
"    background-color: #5d8aad;\n"
"    border: 1px #3079ED solid;\n"
"    box-shadow: inset 0 1px 0 #80B0FB;\n"
"    border-radius:5px;")
        self.html.setObjectName("html")
        self.gridLayout_2.addWidget(self.html, 1, 1, 1, 1)
        self.debug = QtWidgets.QPushButton(self.centralwidget)
        self.debug.setMinimumSize(QtCore.QSize(105, 25))
        self.debug.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.debug.setStyleSheet("    color: white;\n"
"    background-color: #5d8aad;\n"
"    border: 1px #3079ED solid;\n"
"    box-shadow: inset 0 1px 0 #80B0FB;\n"
"    border-radius:5px;")
        self.debug.setObjectName("debug")
        self.gridLayout_2.addWidget(self.debug, 1, 0, 1, 1)
        self.dtailLog = QtWidgets.QPushButton(self.centralwidget)
        self.dtailLog.setMinimumSize(QtCore.QSize(0, 25))
        self.dtailLog.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.dtailLog.setStyleSheet("    color: white;\n"
"    background-color: #5d8aad;\n"
"    border: 1px #3079ED solid;\n"
"    box-shadow: inset 0 1px 0 #80B0FB;\n"
"    border-radius:5px;")
        self.dtailLog.setObjectName("dtailLog")
        self.gridLayout_2.addWidget(self.dtailLog, 2, 1, 1, 1)
        self.file = QtWidgets.QPushButton(self.centralwidget)
        self.file.setMinimumSize(QtCore.QSize(0, 25))
        self.file.setStyleSheet("    color: white;\n"
"    background-color: #5d8aad;\n"
"    border: 1px #3079ED solid;\n"
"    box-shadow: inset 0 1px 0 #80B0FB;\n"
"    border-radius:5px;")
        self.file.setObjectName("file")
        self.gridLayout_2.addWidget(self.file, 0, 0, 1, 1)
        self.analyJSON = QtWidgets.QPushButton(self.centralwidget)
        self.analyJSON.setMinimumSize(QtCore.QSize(105, 25))
        self.analyJSON.setStyleSheet("    color: white;\n"
"    background-color: #5d8aad;\n"
"    border: 1px #3079ED solid;\n"
"    box-shadow: inset 0 1px 0 #80B0FB;\n"
"    border-radius:5px;")
        self.analyJSON.setObjectName("analyJSON")
        self.gridLayout_2.addWidget(self.analyJSON, 2, 0, 1, 1)
        self.dtailReport = QtWidgets.QPushButton(self.centralwidget)
        self.dtailReport.setMinimumSize(QtCore.QSize(0, 25))
        self.dtailReport.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.dtailReport.setStyleSheet("    color: white;\n"
"    background-color: #5d8aad;\n"
"    border: 1px #3079ED solid;\n"
"    box-shadow: inset 0 1px 0 #80B0FB;\n"
"    border-radius:5px;")
        self.dtailReport.setObjectName("dtailReport")
        self.gridLayout_2.addWidget(self.dtailReport, 0, 1, 1, 1)
        self.horizontalLayout.addLayout(self.gridLayout_2)
        self.gridLayout_3 = QtWidgets.QGridLayout()
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.failNum = QtWidgets.QLabel(self.centralwidget)
        self.failNum.setMinimumSize(QtCore.QSize(0, 25))
        self.failNum.setStyleSheet("height:23px;background-color:#ffffff;font: 12pt \'Arial\';color: rgb(255,0,0);")
        self.failNum.setAlignment(QtCore.Qt.AlignCenter)
        self.failNum.setObjectName("failNum")
        self.gridLayout_3.addWidget(self.failNum, 2, 1, 1, 1)
        self.label_9 = QtWidgets.QLabel(self.centralwidget)
        self.label_9.setMinimumSize(QtCore.QSize(0, 25))
        self.label_9.setStyleSheet("height:23px;background-color:#ffffff;")
        self.label_9.setAlignment(QtCore.Qt.AlignCenter)
        self.label_9.setObjectName("label_9")
        self.gridLayout_3.addWidget(self.label_9, 1, 0, 1, 1)
        self.label_8 = QtWidgets.QLabel(self.centralwidget)
        self.label_8.setMinimumSize(QtCore.QSize(105, 25))
        self.label_8.setStyleSheet("height:23px;background-color:#ffffff;")
        self.label_8.setAlignment(QtCore.Qt.AlignCenter)
        self.label_8.setObjectName("label_8")
        self.gridLayout_3.addWidget(self.label_8, 1, 2, 1, 1)
        self.result = QtWidgets.QLabel(self.centralwidget)
        self.result.setMinimumSize(QtCore.QSize(0, 25))
        self.result.setStyleSheet("height:23px;background-color:#ffffff;font: 12pt \'Arial\';")
        self.result.setAlignment(QtCore.Qt.AlignCenter)
        self.result.setObjectName("result")
        self.gridLayout_3.addWidget(self.result, 0, 2, 1, 1)
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setMinimumSize(QtCore.QSize(0, 25))
        self.label_2.setStyleSheet("height:23px;background-color:#ffffff;")
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setObjectName("label_2")
        self.gridLayout_3.addWidget(self.label_2, 0, 1, 1, 1)
        self.skipNum = QtWidgets.QLabel(self.centralwidget)
        self.skipNum.setMinimumSize(QtCore.QSize(0, 25))
        self.skipNum.setStyleSheet("height:23px;background-color:#ffffff;font: 12pt \'Arial\';color: rgb(248,172,89);")
        self.skipNum.setAlignment(QtCore.Qt.AlignCenter)
        self.skipNum.setObjectName("skipNum")
        self.gridLayout_3.addWidget(self.skipNum, 2, 2, 1, 1)
        self.label_7 = QtWidgets.QLabel(self.centralwidget)
        self.label_7.setMinimumSize(QtCore.QSize(105, 25))
        self.label_7.setStyleSheet("height:23px;background-color:#ffffff;")
        self.label_7.setAlignment(QtCore.Qt.AlignCenter)
        self.label_7.setObjectName("label_7")
        self.gridLayout_3.addWidget(self.label_7, 1, 1, 1, 1)
        self.successNum = QtWidgets.QLabel(self.centralwidget)
        self.successNum.setMinimumSize(QtCore.QSize(0, 25))
        self.successNum.setStyleSheet("height:23px;background-color:#ffffff;font: 12pt \'Arial\';color: green;")
        self.successNum.setAlignment(QtCore.Qt.AlignCenter)
        self.successNum.setObjectName("successNum")
        self.gridLayout_3.addWidget(self.successNum, 2, 0, 1, 1)
        self.refresh = QtWidgets.QPushButton(self.centralwidget)
        self.refresh.setMinimumSize(QtCore.QSize(0, 25))
        self.refresh.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.refresh.setStyleSheet("    color: white;\n"
"    background-color: #5d8aad;\n"
"    border: 1px #3079ED solid;\n"
"    box-shadow: inset 0 1px 0 #80B0FB;\n"
"    border-radius:5px;")
        self.refresh.setObjectName("refresh")
        self.gridLayout_3.addWidget(self.refresh, 0, 0, 1, 1)
        self.horizontalLayout.addLayout(self.gridLayout_3)
        self.horizontalLayout.setStretch(0, 3)
        self.horizontalLayout.setStretch(1, 2)
        self.horizontalLayout.setStretch(2, 3)
        self.gridLayout_4.addLayout(self.horizontalLayout, 0, 0, 1, 1)
        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setStyleSheet("background:white")
        self.textEdit.setObjectName("textEdit")
        self.gridLayout_4.addWidget(self.textEdit, 1, 0, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.fileName.setText(_translate("MainWindow", "请选择文件"))
        self.model1.setText(_translate("MainWindow", "普  通"))
        self.model2.setText(_translate("MainWindow", "简  洁"))
        self.html.setText(_translate("MainWindow", "html报告"))
        self.debug.setText(_translate("MainWindow", "开  始"))
        self.dtailLog.setText(_translate("MainWindow", "查看日志"))
        self.file.setText(_translate("MainWindow", "选  择"))
        self.analyJSON.setText(_translate("MainWindow", "解  析"))
        self.dtailReport.setText(_translate("MainWindow", "excel报告"))
        self.failNum.setText(_translate("MainWindow", "0"))
        self.label_9.setText(_translate("MainWindow", "成  功"))
        self.label_8.setText(_translate("MainWindow", "异  常"))
        self.result.setText(_translate("MainWindow", "0/0"))
        self.label_2.setText(_translate("MainWindow", "结果预览:"))
        self.skipNum.setText(_translate("MainWindow", "0"))
        self.label_7.setText(_translate("MainWindow", "失  败"))
        self.successNum.setText(_translate("MainWindow", "0"))
        self.refresh.setText(_translate("MainWindow", "刷  新"))
