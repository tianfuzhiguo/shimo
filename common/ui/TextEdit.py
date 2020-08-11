from PyQt5.QtWidgets import QTextEdit
from PyQt5 import QtCore

'''
@author: dujianxiao
'''
class TextEdit(QTextEdit):
    
    def __init__(self,parent=None):
        super(TextEdit, self).__init__(parent)

        self.console=QTextEdit()
        self.text=self.verticalScrollBar()
        self.text.setStyleSheet("background:rgb(203, 222, 236);width:12px;")
        self.text.hide()
        _translate = QtCore.QCoreApplication.translate
        self.console.setHtml(_translate("mainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p></body></html>"))
#         
        
        

    def enterEvent(self, evt):
        self.text.show()
        
    def leaveEvent(self, evt):
        self.text.hide()