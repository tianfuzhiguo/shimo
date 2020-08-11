import time
import configparser

'''
@author: dujianxiao
'''
class InitConfig():
        
    '''
    @初始化config.ini
    @param path:配置文件路径 
    '''
    def initConfig(self,path):
        try:
            email=[]
            fileData=[]
            config=configparser.ConfigParser()
            config.read(path+'/conf.ini', encoding="utf-8-sig")
            '''
            @预置3个数据库
            '''
            DB1=config.get("section","DB1")
            DB2=config.get("section","DB2")
            DB3=config.get("section","DB3")
            fileData.append(DB1)
            fileData.append(DB2)
            fileData.append(DB3)
            '''         
            @邮件相关
            '''
            sendEmail=config.get("section","sendEmail")
            stmphost=config.get("section","stmphost")
            pwd=config.get("section","pwd")
            receive=config.get("section","receive")
            emailTitle=config.get("section","emailTitle")
            emailContent=config.get("section","emailContent")
            email.append(sendEmail)
            email.append(stmphost)
            email.append(pwd)
            email.append(receive)
            email.append(emailTitle)
            email.append(emailContent)
            '''
            @用户变量
            '''
            try:
                '''
                @读取conf.ini中的用户自定义变量
                '''
                userParams=[]
                userParamsValue=[]
                userParamFile=open(path+'/conf.ini',encoding='utf-8')
                for line in userParamFile.readlines():
                    line=line.strip('\n')
                    if 0<len(line)<9:
                        if line[0]!='#':
                            userParams.append(line[0:line.find('=')])
                            userParamsValue.append(line[line.find('=')+1:])
                    elif len(line)>=9:
                        if line[0]!='#' and line[0:9]!='[section]':
                            userParams.append(line[0:line.find('=')])
                            userParamsValue.append(line[line.find('=')+1:])
            except Exception as e:
                print(e)
            return fileData,email,userParams,userParamsValue
        except Exception as e:
            self.console.append("<font color=\"#000000\"></font> ")
            self.console.append("<font color=\"#FF0000\">"+'初始化conf.ini失败:'+"</font> ")
            self.console.append("<font color=\"#FF0000\">"+str(e)+"</font> ")
            self.console.append("<font color=\"#000000\"></font> ")
            print(e)
    