from PyQt5.QtWidgets import QComboBox, QLineEdit, QListWidget, QCheckBox, QListWidgetItem
from PyQt5.QtCore import pyqtSignal  #导入这个模块才可以创建信号

'''
@author: dujianxiao
'''
class ComboCheckBox(QComboBox):
    popupAboutToBeShown = pyqtSignal()   #创建一个信号
    def loadItems(self, items):
        if len(items)>0 and str(items[0])!=['全部','']:
            items.insert(0,['全部',''])
        self.items = items
        self.row_num = len(self.items)
        self.Selectedrow_num = 0
        self.qCheckBox = []
        self.qLineEdit = QLineEdit() 
        self.qLineEdit.setReadOnly(True)
        self.qListWidget = QListWidget()
        if self.row_num==0:
            pass
        else:
            self.addQCheckBox(0)
            self.qCheckBox[0].stateChanged.connect(self.All)
            for i in range(1, self.row_num):
                self.addQCheckBox(i)
                self.qCheckBox[i].stateChanged.connect(self.showMessage)
        self.setModel(self.qListWidget.model())
        self.setStyleSheet("QAbstractItemView::item {height: 18px;} QScrollBar::vertical{width:0px;border:none;border-radius:5px;}")
        self.setView(self.qListWidget)
        self.setLineEdit(self.qLineEdit)
        
    def showPopup(self):
        self.popupAboutToBeShown.emit()   #发送信号
        select_list = self.Selectlist()  # 当前选择数据
        self.loadItems(items=self.items[1:])  # 重新添加组件
        items=[]
        for i in range(0,len(self.items)):
            items.append(self.items[i][0])
        for select in select_list:
            index = items[:].index(select)
            self.qCheckBox[index].setChecked(True)   # 选中组件
        return QComboBox.showPopup(self)

    def addQCheckBox(self, i):
        self.qCheckBox.append(QCheckBox())
        qItem = QListWidgetItem(self.qListWidget)
        self.qCheckBox[i].setText(str(self.items[i][0]))
        self.qCheckBox[i].setToolTip(str(self.items[i][0]))
        if str(self.items[i][1])=="0":
            self.qCheckBox[i].setCheckable(False) 
            self.qCheckBox[i].setStyleSheet("color:#808080")
        self.qListWidget.setItemWidget(qItem, self.qCheckBox[i])

    def Selectlist(self):
        Outputlist = []
        for i in range(1, self.row_num):
            if self.qCheckBox[i].isChecked() == True:
                Outputlist.append(self.qCheckBox[i].text())
        self.Selectedrow_num = len(Outputlist)
        return Outputlist

    def showMessage(self):
        Outputlist = self.Selectlist()
        if len(Outputlist)>0:
            for i in range(0,len(Outputlist)):
                Outputlist[i]=Outputlist[i][0:str(Outputlist[i]).index(' ')]
        self.qLineEdit.setReadOnly(False)
        self.qLineEdit.clear()
        show = ';'.join(Outputlist)

        if self.Selectedrow_num == 0:
            self.qCheckBox[0].setCheckState(0)
        elif self.Selectedrow_num == self.row_num - 1:
            self.qCheckBox[0].setCheckState(2)
        else:
            self.qCheckBox[0].setCheckState(1)
        self.qLineEdit.setText(show)
        self.qLineEdit.setReadOnly(True)

    def All(self, zhuangtai):
        if zhuangtai == 2:
            for i in range(1, self.row_num):
                self.qCheckBox[i].setChecked(True)
        elif zhuangtai == 1:
            if self.Selectedrow_num == 0:
                self.qCheckBox[0].setCheckState(2)
        elif zhuangtai == 0:
            self.clear()

    def clear(self):
        for i in range(self.row_num):
            self.qCheckBox[i].setChecked(False)

    def currentText(self):
        text = QComboBox.currentText(self).split(';')
        if text.__len__() == 1:
            if not text[0]:
                return []
            else:
                return "('{}')".format("','".join(text))
        else:
            return "('{}')".format("','".join(text))


