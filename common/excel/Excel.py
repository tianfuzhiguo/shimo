from common.utils.ExcelUtil import ExcelUtil
from common.utils.Util import Util
import re, chardet, os, json, datetime, time, demjson3, xmltodict

'''   
@获取校验字段和预期结果的原始值和结果值                                                                                                       
@author: dujianxiao          
'''


class ExcelData:

    def check(self, file, sheet, row, conn):
        """
        校验字段数组－－原值
        @param file: 用例文件
        @param sheet: 页签
        @param row: 行号
        @param conn: 数据库连接对象
        """
        data = ExcelUtil.getList(file, sheet, row, self.part101Col, self.section101Col)
        return [self.repAll(item, file, sheet, row, conn) for item in data]

    def expect(self, file, sheet, row, conn):
        """
        预期结果数组－－原值
        @param file: 用例文件
        @param sheet: 页签
        @param row: 行号
        @param conn: 数据库连接对象
        """
        expect = ExcelUtil.getList(file, sheet, row, self.section101Col, self.resTextCol)
        return [self.repAll(item, file, sheet, row, conn) for item in expect]

    def checkRes(self, r, file, sheet, row, conn):
        """
        校验字段结果数组
        @param r: 接口响应对象
        @param file: 用例文件
        @param sheet: 页签
        @param row: 行号
        @param conn: 数据库连接对象
        """
        jss = self.getResType(r)
        # 固定值数组
        js = ExcelUtil.getList(file, sheet, row, self.part101Col, self.part301Col)
        jsonValue = []
        for item in js:
            jsonValue.append('' if item == '' else eval(jss + item))
        # SQL数组
        sqlList = Util.getSqlResult(file, sheet, row, conn, self.part301Col, self.section101Col)
        result = jsonValue + sqlList
        result = [self.repAll(item, file, sheet, row, conn) for item in result]
        self.getToLog(f'校验字段：{result}')
        return result

    def expectResult(self, file, sheet, row, conn):
        """
        预期结果值数组
        @param file: 用例文件
        @param sheet: 页签
        @param row: 行号
        @param conn: 数据库连接对象
        """
        expectResult1 = ExcelUtil.getList(file, sheet, row, self.section101Col, self.section201Col)
        expectResult2 = Util.getSqlResult(file, sheet, row, conn, self.section201Col, self.section301Col)
        expectResult3 = ExcelUtil.getList(file, sheet, row, self.section301Col, self.resTextCol)
        expectResult = expectResult1 + expectResult2 + expectResult3
        expectResult = [self.repAll(item, file, sheet, row, conn) for item in expectResult]
        self.getToLog(f'预期结果：{expectResult}')
        return expectResult
