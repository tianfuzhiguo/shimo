import re, chardet, os, json, datetime, time, demjson3, xmltodict

'''   
@获取校验字段和预期结果的原始值和结果值                                                                                                       
@author: dujianxiao          
'''


class Array():

    def check(self, file, sheet, row, conn):
        """
        校验字段数组－－原值
        :param file:用例文件
        :param sheet:
        :param row:行号
        :param conn:数据库连接对象
        """
        arr = self.getArray(file, sheet, row, self.part101Col, self.section101Col)
        return [self.repAll(str(item), file, sheet, row, conn) for item in arr]

    def expResultInit(self, file, sheet, row, conn):
        """
        预期结果数组－－原值
        :param file:用例文件
        :param sheet:
        :param row:行号
        :param conn:数据库连接对象
        """
        arr = self.getArray(file, sheet, row, self.section101Col, self.resTextCol)
        return [self.repAll(str(item), file, sheet, row, conn) for item in arr]

    def checkRes(self, r, file, sheet, row, conn):
        """
        校验字段结果数组
        :param r:
        :param file:用例文件
        :param sheet:
        :param row:行号
        :param conn:数据库连接对象
        """
        jss = self.getResType(r)
        # 固定值数组
        js = self.getArray(file, sheet, row, self.part101Col, self.part301Col)
        jsonValue = []
        for item in js:
            jsonValue.append('' if item == '' else eval(jss + item))
        # SQL数组
        sqlArr = self.getSqlResultArray(file, sheet, row, conn, self.part301Col, self.section101Col)
        arr = jsonValue + sqlArr
        arr = [self.repAll(item, file, sheet, row, conn) for item in arr]
        self.getToLog(f'校验字段：{arr}')
        return arr

    def expResult(self, file, sheet, row, conn):
        """
        预期结果值数组
        :param file:用例文件
        :param sheet:
        :param row:行号
        :param conn:数据库连接对象
        """
        arr1 = self.getArray(file, sheet, row, self.section101Col, self.section201Col)
        arr2 = self.getSqlResultArray(file, sheet, row, conn, self.section201Col, self.section301Col)
        arr3 = self.getArray(file, sheet, row, self.section301Col, self.resTextCol)
        arr = arr1 + arr2 + arr3
        arr = [self.repAll(item, file, sheet, row, conn) for item in arr]
        self.getToLog(f'预期结果：{arr}')
        return arr
