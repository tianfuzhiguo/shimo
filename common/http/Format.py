from common.http.Http import Http
from common.utils.Analy import Analy
from jsonschema import validate
from time import sleep
import re, chardet, os, json, datetime, time, demjson3, xmltodict

'''
@author: dujianxiao
'''


class Format(Http, Analy):

    def checkFormat(self, file, sheet, row, conn):
        """
        合法性校验
        @param file: 用例文件
        @param sheet: 页签
        @param row: 行号
        @param conn: 数据库连接对象
        @return: 返回3个值，分别为：http响应、响应时间、异常信息
        """
        # 数据库Ip、用户名、密码错误等引起的异常
        if '数据库异常' in f'{conn}':
            return '', '---', conn
        # 有SQL无数据库连接引起的异常
        DBMsg = self.DBExists(file, sheet, row, conn)
        if DBMsg:
            return '', '---', DBMsg
        # 查询语句错误引起的异常
        DBMsg = self.sqlExcept(file, sheet, row, conn)
        if '数据库异常' in f'{DBMsg}':
            return '', '---', DBMsg
        initMsg = self.initData(file, sheet, row, conn)
        # 数据库初始化语句异常
        if initMsg:
            return '', '---', initMsg
        r, duration, msg = self.httpRequest(file, sheet, row, conn)
        # 不直接返回，列出所有可能的异常
        if '参数异常' in msg:
            return r, duration, msg
        elif '请求头异常' in msg:
            return r, duration, msg
        elif 'url异常' in msg:
            return r, duration, msg
        elif '请求方式异常' in msg:
            return r, duration, msg
        elif '接口请求异常' in msg:
            return r, duration, msg
        elif 'json异常' in msg:
            return r, duration, msg
        elif '信息头管理器异常' in msg:
            return r, duration, msg
        elif '表达式异常' in msg:
            return r, duration, msg
        else:
            return r, duration, msg

    def jsonFormat(self, file, sheet, row, conn):
        """
        把接口返回对象解析成path+value的形式
        @param file: 用例文件
        @param sheet: 页签
        @param row: 行号
        @param conn: 数据库连接对象
        """
        try:
            self.initData(file, sheet, row, conn)
            r, duration, msg = self.httpRequest(file, sheet, row, conn)
            self.restore(file, sheet, row, conn)
            # 处理字符集
            encoding = chardet.detect(r.content).get('encoding')
            if '8859' in f'{encoding}':
                r.encoding = 'utf-8'
            elif '2312' in f'{encoding}' or 'gbk' in f'{encoding}'.lower() or 'gb18130' in f'{encoding}'.lower():
                r.encoding = 'gbk'
            else:
                r.encoding = 'utf-8'
            s1, s2 = self.analy(eval(self.getResType(r)))
            return s1, s2
        except Exception as e:
            self.getError(msg)
            return f'{e}', '解析失败'
