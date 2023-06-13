import cx_Oracle
import pymssql
import pymysql
from common.init.Init import Init
from common.utils.ExcelUtil import ExcelUtil
from common.utils.SmLog import SmLog

'''
@author: dujianxiao
'''


class Util(SmLog, Init):

    def getConn(self, file, sheet, row):
        """
        连接数据库
        @param file: 用例文件
        @param sheet: 页签
        @param row: 行号
        """
        DB = ExcelUtil.getValue(file, sheet, row, self.DBCol)
        if DB == '':
            return [[]]
        else:
            try:
                DB = self.repVar(f'${{{DB}}}')
                conn = eval(DB)
            except Exception as e:
                print(e)
                self.getError(e)
                return [['数据库异常', self.DBCol], [f'{e}']]
            return [[conn]]

    def DBExists(self, file, sheet, row, conn):
        """
        有SQL则数据库不允许为空
        """
        allSql = ExcelUtil.getList(file, sheet, row, self.part301Col, self.section101Col) + \
                 ExcelUtil.getList(file, sheet, row, self.section201Col, self.section301Col) + \
                 ExcelUtil.getList(file, sheet, row, self.init001Col, self.key001Col)
        # 如果sql不全为空而数据库连接为空，则返回数据库异常，否则返回空数组
        return [['数据库异常', self.DBCol], []] if ''.join(allSql) != '' and conn == [[]] else []

    @staticmethod
    def getSqlResult(file, sheet, row, conn, start, end):
        """
        获取sql结果数组
        @param file: 用例文件
        @param sheet: 页签
        @param row: 行号
        @param conn: 数据库连接对象
        @param start: 索引开始
        @param end: 索引结束
        @return List: sql结果
        """
        data = []
        sqlList = ExcelUtil.getList(file, sheet, row, start, end)
        column = start
        # 在此之前已经校验过sql不全为空而数据库为空的情况
        # 所以如果此时数据库为空，说明sql全为空
        if conn == [[]]:
            return sqlList
        else:
            cursor = conn[0][0].cursor()
            for item in sqlList:
                if item == '':
                    data.append('')
                else:
                    cursor.execute(item)
                    one = cursor.fetchone()
                    data.append(None if one is None else one[0])
                column = column + 1
            cursor.close()
            conn[0][0].commit()
            return data

    def sqlExcept(self, file, sheet, row, conn):
        """
        此方法仅用作验证SQL句是否正确
        @param file: 用例文件
        @param sheet: 页签
        @param row: 行号
        @param conn: 数据库连接对象
        """
        conn = self.getConn(file, sheet, row)
        msg = ['数据库异常']
        sqlMsg = []
        sqlList = (ExcelUtil.getList(file, sheet, row, self.part301Col, self.section101Col) +
                   ExcelUtil.getList(file, sheet, row, self.section201Col, self.section301Col) +
                   ExcelUtil.getList(file, sheet, row, self.dyparam001Col, self.key001Col))
        sqlList = [self.repRel(item) for item in sqlList]
        sqlList = [self.repVar(item) for item in sqlList]
        if conn == [[]]:
            return []
        else:
            column1 = self.part301Col
            cursor = conn[0][0].cursor()
            for item in sqlList:
                if item != '':
                    try:
                        # 这三部分的SQL只能是查询语句
                        if not (f'{item}'.lower()).replace(' ', '').startswith('select'):
                            self.getToLog('part301Col,section201Col,dyparam001Col这三部分只能是select语句:' + f'{item}')
                            msg.append(column1)
                            sqlMsg.append(f'{item}')
                        else:
                            cursor.execute(item)
                    except Exception as e:
                        print(e)
                        msg.append(column1)
                        self.getToLog(item)
                        self.getError(e)
                        sqlMsg.append(f'{e}')
                if column1 == self.section101Col - 1:
                    column1 = self.section201Col - 1
                if column1 == self.section301Col - 1:
                    column1 = self.dyparam001Col - 1
                column1 = column1 + 1
            cursor.close()
            conn[0][0].close()
            return [msg, sqlMsg] if len(msg) > 1 else []

    def initData(self, file, sheet, row, conn):
        """
        数据初始化
        @param file: 用例文件
        @param sheet: 页签
        @param row: 行号
        @param conn: 数据库连接对象
        """
        msg = ['数据库异常']
        sqlMsg = []
        column1 = self.init001Col
        sqlList = ExcelUtil.getList(file, sheet, row, self.init001Col, self.restore001Col)
        sqlList = [self.repRel(item) for item in sqlList]
        sqlList = [self.repVar(item) for item in sqlList]
        if conn == [[]]:
            return []
        else:
            cursor = conn[0][0].cursor()
            for item in sqlList:
                if item != '':
                    try:
                        item = self.rep(file, sheet, row, conn, item)
                        cursor.execute(item)
                        conn[0][0].commit()
                        self.getToLog(f'数据初始化：{item}')
                    except Exception as e:
                        print(e)
                        msg.append(f'{column1}')
                        self.getToLog(item)
                        self.getError(e)
                        sqlMsg.append(f'{e}')
                column1 = column1 + 1
            cursor.close()
            return [msg, sqlMsg] if len(msg) > 1 else []

    def restore(self, file, sheet, row, conn):
        """
        数据恢复
        @param file: 用例文件
        @param sheet: 页签
        @param row: 行号
        @param conn: 数据库连接对象
        """
        msg = ['数据库异常']
        sqlMsg = []
        column1 = self.restore001Col
        sqlList = ExcelUtil.getList(file, sheet, row, self.restore001Col, self.dyparam001Col)
        sqlList = [self.repRel(item) for item in sqlList]
        sqlList = [self.repVar(item) for item in sqlList]
        if conn == [[]]:
            return []
        else:
            cursor = conn[0][0].cursor()
            for item in sqlList:
                if item != '':
                    try:
                        item = self.rep(file, sheet, row, conn, item)
                        cursor.execute(item)
                        conn[0][0].commit()
                        self.getToLog(f'数据恢复：{item}')
                    except Exception as e:
                        print(e)
                        msg.append(f'{column1}')
                        self.getError(e)
                        sqlMsg.append(f'{e}')
                column1 = column1 + 1
            cursor.close()
            return [msg, sqlMsg] if len(msg) > 1 else []

    def dyparam(self, file, sheet, row, conn):
        """
        动态化参数-数据库查询结果作为参数供同一行的其他地方调用
        @param file: 用例文件
        @param sheet: 页签
        @param row: 行号
        @param conn: 数据库连接对象
        """
        data = []
        sqlList = ExcelUtil.getList(file, sheet, row, self.dyparam001Col, self.key001Col)
        sqlList = [self.repRel(item) for item in sqlList]
        sqlList = [self.repVar(item) for item in sqlList]
        if conn == [[]]:
            return []
        else:
            cursor = conn[0][0].cursor()
            for item in sqlList:
                if item == '':
                    data.append('')
                else:
                    cursor.execute(item)
                    one = cursor.fetchone()
                    data.append(None if one is None else one[0])
                    conn[0][0].commit()
            cursor.close()
            return data

    def repVar(self, string):
        """
        替换字符串中的用户变量为真实值
        @param string: 需要替换的字符串
        """
        for key, value in self.fileData:
            string = string.replace(f'${{{key}}}', f'{value}')
        return string

    def repRel(self, string):
        """
        替换字符串中的接口变量为真实值
        @param string: 需要替换的字符串
        """
        for key, value in self.interData.items():
            string = string.replace(f'${{{key}}}', f'{value}')
        return string

    def rep(self, file, sheet, row, conn, string):
        """
        替换字符串中的动态参数为真实值
        @param file: 用例文件
        @param sheet: 页签
        @param row: 行号
        @param conn: 数据库连接对象
        @param string: 需要替换的字符串
        """
        dypar = self.dyparam(file, sheet, row, conn)
        if not dypar:
            return string
        for i, value in enumerate(dypar, start=1):
            dyparam = 'dyparam' + f'{i}'.zfill(3)
            string = string.replace(f'${{{dyparam}}}', f'{value}')
        return string

    def repAll(self, string, file, sheet, row, conn):
        """
        替换字符串中的用户变量、动态参数、接口变量为真实值
        @param string:
        @param file: 用例文件
        @param sheet: 页签
        @param row: 行号
        @param conn: 数据库连接对象
        """
        string = f'{string}'
        string = self.repVar(string)
        string = self.rep(file, sheet, row, conn, string)
        string = self.repRel(string)
        return string
