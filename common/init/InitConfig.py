import configparser
import os.path

'''
@author: dujianxiao
'''


class InitConfig:

    def initConfig(self, path):
        """
        初始化config.ini
        :param path:配置文件路径
        """
        try:
            # 读取conf.ini中的用户自定义变量
            config = configparser.ConfigParser()
            config.read(os.path.join(path, 'conf.ini'), encoding="utf-8")
            return config.items('section')
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
