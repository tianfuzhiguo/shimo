import cx_Oracle
import pymssql
import pymysql

from common.init.Init import Init
from common.utils.SmLog import SmLog

'''
@author: dujianxiao
'''


class Util(SmLog, Init):

    def getConn(self, file, sheet, row):
        """
        连接数据库
        :param file:用例文件
        :param sheet:
        :param row:行号
        """
        DB = self.getValue(file, sheet, row, self.DBCol)
        DB = self.repVar(DB)
        if DB == '':
            return [[]]
        else:
            try:
                DB = self.repVar('${' + f'{DB}' + '}')
                conn = eval(DB)
            except Exception as e:
                print(e)
                self.getError(str(e))
                return [['数据库异常', self.DBCol], [str(e)]]
            return [[conn]]

    def DBExists(self, file, sheet, row, conn):
        """
        有SQL则数据库不允许为空
        """
        allSql = self.getArray(file, sheet, row, self.part301Col, self.section101Col) + \
                 self.getArray(file, sheet, row, self.section201Col, self.section301Col) + \
                 self.getArray(file, sheet, row, self.init001Col, self.key001Col)
        # 如果sql不全为空而数据库连接为空，则返回数据库异常，否则返回空数组
        return [['数据库异常', self.DBCol], []] if ''.join(allSql) != '' and conn == [[]] else []

    def getSqlResultArray(self, file, sheet, row, conn, start, end):
        """
        获取sql结果数组
        :param file:用例文件
        :param sheet:
        :param row:行号
        :param conn:数据库连接对象
        :param start:
        :param end:
        """
        data = []
        sqlArray = self.getArray(file, sheet, row, start, end)
        column = start
        # 在此之前已经校验过sql不全为空而数据库为空的情况
        # 所以如果此时数据库为空，说明sql全为空
        if conn == [[]]:
            return sqlArray
        else:
            cursor = conn[0][0].cursor()
            for item in sqlArray:
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
        :param file:用例文件
        :param sheet:
        :param row:行号
        :param conn:数据库连接对象
        """
        conn = self.getConn(file, sheet, row)
        msg = ['数据库异常']
        SqlMsg = []
        sqlArray = (self.getArray(file, sheet, row, self.part301Col, self.section101Col) +
                    self.getArray(file, sheet, row, self.section201Col, self.section301Col) +
                    self.getArray(file, sheet, row, self.dyparam001Col, self.key001Col))
        sqlArray = [self.repRel(item) for item in sqlArray]
        sqlArray = [self.repVar(item) for item in sqlArray]
        if conn == [[]]:
            return []
        else:
            column1 = self.part301Col
            cursor = conn[0][0].cursor()
            for item in sqlArray:
                if item != '':
                    try:
                        # 这三部分的SQL只能是查询语句
                        if not (str(item).lower()).replace(' ', '').startswith('select'):
                            self.getToLog('part301Col,section201Col,dyparam001Col这三部分只能是select语句:' + str(item))
                            msg.append(column1)
                            SqlMsg.append(str(item))
                        else:
                            cursor.execute(item)
                    except Exception as e:
                        print(e)
                        msg.append(column1)
                        self.getToLog(item)
                        self.getError(str(e))
                        SqlMsg.append(str(e))
                #                     msg.append(str(e))
                if column1 == self.section101Col - 1:
                    column1 = self.section201Col - 1
                if column1 == self.section301Col - 1:
                    column1 = self.dyparam001Col - 1
                column1 = column1 + 1
            cursor.close()
            conn[0][0].close()
            return [msg, SqlMsg] if len(msg) > 1 else []

    def initData(self, file, sheet, row, conn):
        """
        数据初始化
        :param file:用例文件
        :param sheet:
        :param row:行号
        :param conn:数据库连接对象
        """
        msg = ['数据库异常']
        SqlMsg = []
        column1 = self.init001Col
        sqlArray = self.getArray(file, sheet, row, self.init001Col, self.restore001Col)
        sqlArray = [self.repRel(item) for item in sqlArray]
        sqlArray = [self.repVar(item) for item in sqlArray]
        if conn == [[]]:
            return []
        else:
            cursor = conn[0][0].cursor()
            for item in sqlArray:
                if item != '':
                    try:
                        item = self.rep(file, sheet, row, conn, item)
                        cursor.execute(item)
                        conn[0][0].commit()
                        self.getToLog('数据初始化：' + item)
                    except Exception as e:
                        print(e)
                        msg.append(str(column1))
                        self.getError(str(e))
                        SqlMsg.append(str(e))
                column1 = column1 + 1
            cursor.close()
            return [msg, SqlMsg] if len(msg) > 1 else []

    def restore(self, file, sheet, row, conn):
        """
        数据恢复
        :param file:用例文件
        :param sheet:
        :param row: 行号
        :param conn:数据库连接对象
        """
        msg = ['数据库异常']
        SqlMsg = []
        column1 = self.restore001Col
        sqlArray = self.getArray(file, sheet, row, self.restore001Col, self.dyparam001Col)
        sqlArray = [self.repRel(item) for item in sqlArray]
        sqlArray = [self.repVar(item) for item in sqlArray]
        if conn == [[]]:
            return []
        else:
            cursor = conn[0][0].cursor()
            for item in sqlArray:
                if item != '':
                    try:
                        item = self.rep(file, sheet, row, conn, item)
                        cursor.execute(item)
                        conn[0][0].commit()
                        self.getToLog('数据恢复：' + item)
                    except Exception as e:
                        print(e)
                        msg.append(str(column1))
                        self.getError(str(e))
                        SqlMsg.append(str(e))
                column1 = column1 + 1
            cursor.close()
            return [msg, SqlMsg] if len(msg) > 1 else []

    def dyparam(self, file, sheet, row, conn):
        """
        动态化参数-数据库查询结果作为参数供同一行的其他地方调用
        :param file:用例文件
        :param sheet:
        :param row: 行号
        :param conn: 数据库连接对象
        """
        data = []
        sqlArray = self.getArray(file, sheet, row, self.dyparam001Col, self.key001Col)
        sqlArray = [self.repRel(item) for item in sqlArray]
        sqlArray = [self.repVar(item) for item in sqlArray]
        if conn == [[]]:
            return []
        else:
            cursor = conn[0][0].cursor()
            for item in sqlArray:
                if item == '':
                    data.append('')
                else:
                    cursor.execute(item)
                    one = cursor.fetchone()
                    data.append(None if one is None else one[0])
                    conn[0][0].commit()
            cursor.close()
            return data

    def repVar(self, param):
        """
        用户变量替换
        :param param:需要替换的值
        """
        param = str(param)
        for i in range(len(self.userParams)):
            if '${' + self.userParams[i] + '}' in param:
                param = param.replace('${' + str(self.userParams[i]) + '}', str(self.userParamsValue[i]))
        return param

    def repRel(self, param):
        """
        接口变量替换
        :param param:需要替换的值
        """
        param = str(param)
        for i in range(len(self.userVar)):
            if '${' + str(self.userVar[i]) + '}' in str(param):
                param = param.replace('${' + str(self.userVar[i]) + '}', str(self.userVarValue[i]))
        return param

    def rep(self, file, sheet, row, conn, param):
        """
        动态参数替换
        :param file:用例文件
        :param sheet:
        :param row: 行号
        :param conn: 数据库连接对象
        :param param: 需要替换的值
        """
        dypArr = []
        dypar = self.dyparam(file, sheet, row, conn)
        if dypar is None:
            return param
        else:
            for i in range(1, len(dypar) + 1):
                dypArr.append('dyparam' + str(i).zfill(3))
            for i in range(len(dypar)):
                if '${' + str(dypArr[i]) + '}' in param:
                    param = param.replace('${' + str(dypArr[i]) + '}', str(dypar[i]))
        return param

    def repAll(self, param, file, sheet, row, conn):
        """
        三者替换，替换用户变量、动态参数、接口变量
        :param param:
        :param file:用例文件
        :param sheet:
        :param row: 行号
        :param conn: 数据库连接对象
        """
        param = str(param)
        param = self.repVar(param)
        param = self.rep(file, sheet, row, conn, param)
        param = self.repRel(param)
        return param
