import logging
import os

'''
@author: dujianxiao
'''


class SmLog:
    @staticmethod
    def initLog(path):
        logger = None
        try:
            filename = f'{path}/result/info.log'
            isExists = os.path.exists(f'{path}/result')
            if os.path.exists(filename):
                os.remove(filename)
            if not isExists:
                os.makedirs(f'{path}/result')
        except Exception as e:
            print(e)
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

    @staticmethod
    def getToLog(info):
        info = str(info)
        logging.debug(info)

    @staticmethod
    def getError(info):
        info = str(info)
        logging.error(info)
