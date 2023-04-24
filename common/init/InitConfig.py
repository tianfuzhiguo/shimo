import configparser

'''
@author: dujianxiao
'''


class InitConfig():

    def initConfig(self, path):
        """
        初始化config.ini
        :param path:配置文件路径
        """
        try:
            fileData = []
            config = configparser.RawConfigParser()
            config.read(path + '/conf.ini', encoding="utf-8-sig")
            #预置3个数据库
            DB1 = config.get("section", "DB1")
            DB2 = config.get("section", "DB2")
            DB3 = config.get("section", "DB3")
            fileData.append(DB1)
            fileData.append(DB2)
            fileData.append(DB3)
            items = config.items('section')
            userParams = []
            userParamsValue = []
            for i in range(0, len(items)):
                userParams.append(items[i][0])
                userParamsValue.append(items[i][1])
            return fileData, userParams, userParamsValue
        except Exception as e:
            self.consoleFunc('red', '初始化conf.ini失败:')
            self.consoleFunc('black', str(e))
            print(e)

    def consoleFunc(self, color, content='', size=''):
        """
        设置字体颜色和大小
        :param size: 字号
        :param content: 内容
        :param color:字体颜色
        """
        self.console.append(f"<font {size} color={color}>{content}</font>")
