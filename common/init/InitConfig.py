import configparser

'''
@author: dujianxiao
'''


class InitConfig():
    '''
    @初始化config.ini
    @param path:配置文件路径 
    '''

    def initConfig(self, path):
        try:
            fileData = []
            config = configparser.ConfigParser()
            config.read(path + '/conf.ini', encoding="utf-8-sig")
            '''
            @预置3个数据库
            '''
            DB1 = config.get("section", "DB1")
            DB2 = config.get("section", "DB2")
            DB3 = config.get("section", "DB3")
            fileData.append(DB1)
            fileData.append(DB2)
            fileData.append(DB3)
            try:
                '''
                @读取conf.ini中的用户自定义变量
                '''
                userParams = []
                userParamsValue = []
                with open(path + '/conf.ini', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip('\n')
                        if 0 < len(line) < 9:
                            if line[0] != '#':
                                userParams.append(line[0:line.find('=')])
                                userParamsValue.append(line[line.find('=') + 1:])
                        elif len(line) >= 9:
                            if line[0] != '#' and line[0:9] != '[section]':
                                userParams.append(line[0:line.find('=')])
                                userParamsValue.append(line[line.find('=') + 1:])
            except Exception as e:
                print(e)
            return fileData, userParams, userParamsValue
        except Exception as e:
            self.consoleFunc('red', '初始化conf.ini失败:')
            self.consoleFunc('black', str(e))
            print(e)

    '''
    @设置字体颜色和大小
    @param color
    '''

    def consoleFunc(self, color, content='', size=''):
        self.console.append("<font " + size + " color=" + color + ">" + content + "</font>")
