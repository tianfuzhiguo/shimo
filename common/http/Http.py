from common.utils.ExcelUtil import ExcelUtil
from common.utils.Util import Util
from time import sleep
from jsonschema import validate
from requests.adapters import HTTPAdapter
import requests, re, chardet, os, json, datetime, time, demjson3, xmltodict

'''
@author: dujianxiao
'''


class Http(Util):
    res = requests.session()
    interData = {}  # 接口变量数组
    headerManager = ''  # 请求头

    def httpRequest(self, file, sheet, row, conn):
        """
        http请求
        @param file: 用例文件
        @param sheet: 页签
        @param row: 行号
        @param conn: 数据库连接对象
        """
        interface = ExcelUtil.getList(file, sheet, row, self.key001Col, self.value001Col)  # 接口变量数组
        url, method, body, files, header = [self.repAll(ExcelUtil.getValue(file, sheet, row, self.column[i + 1]),
                                                        file, sheet, row, conn) for i in range(5)]
        # 如果请求头为空且信息头管理器不为空，则使用信息头管理器覆盖
        if header == '' and self.headerManager != '':
            header = self.headerManager
        self.getToLog(f"url：{url}")
        self.getToLog(f"method：{method}")
        self.getToLog(f"params：{body}")
        self.getToLog(f"files：{files}")
        self.getToLog(f"headers：{header}")

        # 校验参数格式，需要是JSON格式
        try:
            body = body.replace('&', '%26').replace('=', '%3D')  # 对&和=进行转义
            body = {} if body == '' else json.loads(body)
        except Exception as e:
            print(e)
            msg = ['参数异常', self.paramCol]
            return '', '---', msg
        # 校验请求头格式，需要是JSON格式
        try:
            if header == '':
                self.headerManager
            else:
                header = json.loads(header)
        except Exception as e:
            print(e)
            msg = ['请求头异常', self.headerCol]
            return '', '---', msg
        # 暂时只支持四种请求方式
        if f'{method}'.upper() not in ['GET', 'POST', 'PUT', 'DELETE']:
            msg = ['请求方式异常', self.methodCol]
            return '', '---', msg
        r, duration, msg = self.sendHttp(url, method, body, header, files)
        try:
            # 字符集处理
            encoding = chardet.detect(r.content).get('encoding')
            # 接口请求异常返回的可能不是response,为了不影响逻辑try一下
            if '8859' in f'{encoding}':
                r.encoding = 'utf-8'
            elif '2312' in f'{encoding}' or 'gbk' in f'{encoding}'.lower() or 'gb18130' in f'{encoding}'.lower():
                r.encoding = 'gbk'
            else:
                r.encoding = 'utf-8'
        except Exception as e:
            print(e)
        if '异常' in f'{msg}':
            return r, duration, msg
        else:
            self.getToLog(r)
            try:
                # 某些接口返回值是html格式,会出现大量的转义字符,使用loads进行反序列化
                self.getToLog(f"接口响应:{json.loads(r.text)}")
            except:
                # 又由于有些接口返回值不是json格式,不能loads,所以如果反序列化失败即不再进行反序列化
                self.getToLog(f"接口响应:{r.text}")
            self.getToLog(f"响应头：{r.headers}")
            msg3 = Http.analyJSON(self, file, sheet, row, r)
            # 校验字段或接口变量字段等JSON字段错误
            if 'json异常' in msg3:
                return r, duration, msg3
            else:
                msg = msg3
            self.setParams(r, interface, file, sheet, row)
            msg = self.setHeader(file, sheet, row, conn)
            if len(msg) > 1:
                return r, duration, msg
        expressionMsg = self.validateExp(r, file, sheet, row, conn)
        if len(expressionMsg) > 1:
            return r, duration, expressionMsg
        return r, duration, ''

    def setHeader(self, file, sheet, row, conn):
        """
        设置信息头管理器，供后续接口隐式调用，如果接口设置了信息头，则信息头管理器在此接口中无效
        """
        msg = ['信息头管理器异常']
        header = ExcelUtil.getValue(file, sheet, row, self.headerManagerCol)
        try:
            if header != '':
                json.loads(header)
                self.headerManager = self.repAll(header, file, sheet, row, conn)
        except Exception as e:
            print(e)
            msg.append(self.headerManagerCol)
        return msg

    def setParams(self, r, interface, file, sheet, row):
        """
        接口请求成功后存接口变量
        如果已经存在同名的变量则覆盖，否则新建一个变量
        @param r: 接口响应对象
        @param interface: 接口变量数组
        @param file: 用例文件
        @param sheet: 页签
        @param row: 行号
        """
        js = self.getResType(r)
        num = len(interface)
        for i in range(len(interface)):
            relate = ExcelUtil.getValue(file, sheet, row, self.key001Col + i + num)
            if interface[i] != '':
                self.interData.update({relate: eval(js + interface[i])})

    def validateExp(self, r, file, sheet, row, conn):
        """
        表达式异常校验
        @param conn: 数据库连接对象
        @param row: 行号
        @param sheet: 页签
        @param r: 接口响应对象
        @param file: 用例文件
        """
        expressMsg = ['表达式异常']
        column1 = self.expressCol
        js = self.getResType(r)
        express = ExcelUtil.getList(file, sheet, row, self.expressCol, self.statusCol)
        for i in range(len(express)):
            express[i] = self.repAll(express[i], file, sheet, row, conn)
            # 表达式可能涉及到接口响应
            express[i] = f'{express[i]}'.replace("r.json()", js)
            if express[i] != '':
                try:
                    eval(express[i])
                except Exception as e:
                    print(e)
                    self.getError(f'{express[i]}:{e}')
                    expressMsg.append(f'{column1}')
            column1 = column1 + 1
        return expressMsg

    def analyJSON(self, file, sheet, row, r):
        """
        解析接口响应
        @param file: 用例文件
        @param sheet: 页签
        @param row: 行号
        @param r: 接口请求返回对象
        """
        msg = ['json异常']
        res = []
        # 取出用例中所有的JSON字段
        checkList = ExcelUtil.getList(file, sheet, row, self.part101Col, self.part301Col) + \
                    ExcelUtil.getList(file, sheet, row, self.key001Col, self.value001Col)
        col = self.part101Col
        js = self.getResType(r)
        for item in checkList:
            if item == '':
                res.append('')
            else:
                try:
                    item = self.repVar(item)
                    res.append(eval(js + item))
                except Exception as e:
                    print(e)
                    msg.append(f'{col}')
            if col < self.part301Col - 1:
                col = col + 1
            elif col == self.part301Col - 1:
                col = col + (self.key001Col - self.part301Col) + 1
        # 返回异常信息或JSON值数组
        return msg if len(msg) > 1 else res

    def getResType(self, r):
        """
        return 响应类型,json.jsonp,xml
        """
        if f'{r}' == '':
            return ['']
        else:
            js1 = 'demjson3.decode(r.text)'
            js2 = 'json.dumps(xmltodict.parse(r.text))'
            js3 = 'json.loads(re.search("^[^(]*?\((.*)\)[^)]*$",r.text,re.S).group(1))'
        try:
            # 返回类型为json
            eval(js1)
            js = js1
        except:
            # 返回类型为xml
            try:
                js = eval(js2)
            except:
                # 返回类型为jsonp
                try:
                    eval(js3)
                    js = js3
                except:
                    js = f'{r.text}'
        return js

    def sendHttp(self, url, method, body, header, files):
        """
        @param url: 地址
        @param method: 请求方式
        @param body: 参数
        @param header: 请求头
        @param files: 上传文件
        @return: r,响应时间,异常信息
        """
        msg = ['请求方式异常', self.methodCol]
        methods = {'GET': 'self.get(url,body,header)',
                   'POST': 'self.post(url,body,header,files)',
                   'DELETE': 'self.delete(url,body,header)',
                   'PUT': 'self.put(url,body,header,files)'}
        key = [f'{method}'.upper()] + ExcelUtil.filterList(methods.keys(), f'{method}'.upper())  # 把传入的method放到第一个，提高效率
        try:
            resp, duration = eval(methods[key[0]])
        except Exception as e:
            print(e)
            self.getError(e)
            if 'Invalid URL' in f'{e}':
                msg = ['url异常', self.urlCol]
                return f'{e}', '---', msg
            elif f'Failed to parse: {url}' == f'{e}':
                msg = ['url异常', self.urlCol]
                return f'{e}', '---', msg
            else:
                msg = ['接口请求异常', self.urlCol, self.methodCol, self.paramCol, self.fileCol, self.headerCol, ]
                return f'{e}', '---', msg
        # 如果请求失败，则调用其他请求方式，其他方式有200说明请求方式写错了
        if '200' in f'{resp}':
            return resp, duration, ['']
        elif '200' in f'{eval(methods[key[1]])}':
            return resp, duration, msg
        elif '200' in f'{eval(methods[key[2]])}':
            return resp, duration, msg
        elif '200' in f'{eval(methods[key[3]])}':
            return resp, duration, msg
        else:
            return resp, duration, ['']

    def get(self, url, body={}, header={}):
        r = self.res.get(url, params=body, headers=header, timeout=30)
        duration = f'{r.elapsed.total_seconds()}'
        return r, duration[:-3]

    def post(self, url, body={}, header={}, files=""):
        if body != {} and 'application/json' in f'{header}'.lower():
            body = json.dumps(body)
        r = self.res.post(url, data=body, headers=header, files=files, timeout=30)
        duration = f'{r.elapsed.total_seconds()}'
        return r, duration[:-3]

    def delete(self, url, body={}, header={}):
        r = self.res.delete(url, json=body, headers=header, timeout=30)
        duration = f'{r.elapsed.total_seconds()}'
        return r, duration[:-3]

    def put(self, url, body={}, header={}, files=""):
        if body != {} and 'application/json' in f'{header}'.lower():
            body = json.dumps(body)
        r = self.res.put(url, data=body, headers=header, files=files, timeout=30)
        duration = f'{r.elapsed.total_seconds()}'
        return r, duration[:-3]
