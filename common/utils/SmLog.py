import logging, os

'''
@author: dujianxiao
'''

class SmLog:

    def initLog(self, path):
        global logger
        filename = path + '/result/info.log'
        try:
            isExists = os.path.exists(path + '/result')
            if not isExists:
                os.makedirs(path + '/result')
            filename = path + '/result/info.log'
            # 每次运行删除之前的日志，
            os.remove(filename)
        except:
            pass
        try:
            if logger is None:
                pass
        except:
            logger = None
        if logger is None:
            logger = logging.getLogger()
        else:
            for handler in logger.handlers[:]:
                logger.removeHandler(handler)
        logger.setLevel(logging.DEBUG)
        formatter = logging.Formatter(fmt='%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        fh = logging.FileHandler(filename, encoding='utf-8')
        fh.setFormatter(formatter)
        logger.addHandler(fh)

    def getToLog(self, info):
        logging.debug(info)

    def getError(self, info):
        logging.error(info)
