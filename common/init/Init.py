from common.init.InitConfig import InitConfig
from common.init.InitExcel import InitExcel

'''
@初始化用例文件和配置文件
'''


class Init(InitConfig, InitExcel):

    def __init__(self):
        self.fileData = None
        self.ncols = None
        self.column = None
        self.nameCol = None
        self.urlCol = None
        self.methodCol = None
        self.paramCol = None
        self.fileCol = None
        self.headerCol = None
        self.part101Col = None
        self.part201Col = None
        self.part301Col = None
        self.section101Col = None
        self.section201Col = None
        self.section301Col = None
        self.resTextCol = None
        self.resHeaderCol = None
        self.statusCodeCol = None
        self.expressCol = None
        self.statusCol = None
        self.timeCol = None
        self.init001Col = None
        self.restore001Col = None
        self.dyparam001Col = None
        self.key001Col = None
        self.value001Col = None
        self.headerManagerCol = None
        self.DBCol = None
        self.iterationCol = None

    def initFile(self, date, path, file, sheetName):
        """
        初始化用例文件和配置文件
        @param date: 日期
        @param path: 路径
        @param file: 文件
        @param sheetName: 页签名
        """
        self.fileData = self.initConfig(path)
        sheet, nrows, self.ncols = InitExcel.getSheet(date, path, sheetName, file)
        self.column = self.getColumn(file, sheet)
        self.nameCol = self.column[0]
        self.urlCol = self.column[1]
        self.methodCol = self.column[2]
        self.paramCol = self.column[3]
        self.fileCol = self.column[4]
        self.headerCol = self.column[5]
        self.part101Col = self.column[6]
        self.part201Col = self.column[7]
        self.part301Col = self.column[8]
        self.section101Col = self.column[9]
        self.section201Col = self.column[10]
        self.section301Col = self.column[11]
        self.resTextCol = self.column[12]
        self.resHeaderCol = self.column[13]
        self.statusCodeCol = self.column[14]
        self.expressCol = self.column[15]
        self.statusCol = self.column[16]
        self.timeCol = self.column[17]
        self.init001Col = self.column[18]
        self.restore001Col = self.column[19]
        self.dyparam001Col = self.column[20]
        self.key001Col = self.column[21]
        self.value001Col = self.column[22]
        self.headerManagerCol = self.column[23]
        self.DBCol = self.column[24]
        self.iterationCol = self.column[25]
        return sheet, nrows
