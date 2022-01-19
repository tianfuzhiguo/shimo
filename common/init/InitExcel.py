from common.utils.ExcelUtil import ExcelUtil 
import shutil,os                                                                                                                                                                                                                                                                                                            
'''                                                                                                                                                                                                                                                                                                         
@author: dujianxiao                                                                                                                                                                                                                                                                                         
'''                                                                                                                                                                                                                                                                                                         
class InitExcel(ExcelUtil):     
    '''
    @读用例
    '''   
    def getBook(self,path,file):
        book=''
        try:
            book = self.readExcel(path+'/'+file)
            return book
        except Exception as e:
            print(e)
            self.consoleFunc('red', str(e))
            return book                  
    '''
    @获取页签及其行数、列数
    '''    
    def getSheet(self,reportDate,path,sheetName,file,book):
        try:
            book = self.readExcel(path+'/'+file)
        except Exception as e:
            self.consoleFunc('red', str(e))
        try:
            '''
            #生测试报告，历史报告移动到history中
            '''
            if file.endswith('xls'):
                sheet = book.sheet_by_name(sheetName)
                nrows = sheet.nrows
                ncols = sheet.ncols
            elif file.endswith('xlsx'):
                sheet = book.get_sheet_by_name(sheetName)
                nrows = sheet.max_row
                ncols = sheet.max_column
            try:
                isExists=os.path.exists(path+'/result/history')
                if not isExists:
                    os.makedirs(path+'/result/history') 
                fileList=os.listdir(path+'/result/')
                for i in range(len(fileList)):
                    if str(reportDate) not in str(fileList[i]) and str('report') in str(fileList[i]):
                        try:              
                            shutil.move(path+'/result/'+str(fileList[i]), path+'/result/history')
                        except Exception as e:
                            print(e)
                return sheet,nrows,ncols
            except Exception as e:
                print(e)  
        except Exception as e:
            print(e)                   
        
    '''
    @获取sheet页各标志位的列号
    @param file:用例文件
    @param sheet:
    @param ncols:列数   
    '''    
    def getColumn(self,file,sheet): 
        column=[]   
        column.append(self.findStr(file,sheet,'name'))
        column.append(self.findStr(file,sheet,'url'))
        column.append(self.findStr(file,sheet,'method'))
        column.append(self.findStr(file,sheet,'param'))
        column.append(self.findStr(file,sheet,'file'))
        column.append(self.findStr(file,sheet,'header'))
        column.append(self.findStr(file,sheet,'part101'))
        column.append(self.findStr(file,sheet,'part201'))
        column.append(self.findStr(file,sheet,'part301'))
        column.append(self.findStr(file,sheet,'section101'))
        column.append(self.findStr(file,sheet,'section201'))
        column.append(self.findStr(file,sheet,'section301'))
        column.append(self.findStr(file,sheet,'resText'))
        column.append(self.findStr(file,sheet,'resHeader'))
        column.append(self.findStr(file,sheet,'statusCode'))
        column.append(self.findStr(file,sheet,'expression'))
        column.append(self.findStr(file,sheet,'status'))
        column.append(self.findStr(file,sheet,'time'))
        column.append(self.findStr(file,sheet,'init001'))
        column.append(self.findStr(file,sheet,'restore001'))
        column.append(self.findStr(file,sheet,'dyparam001'))
        column.append(self.findStr(file,sheet,'key001'))
        column.append(self.findStr(file,sheet,'value001'))
        column.append(self.findStr(file,sheet,'headerManager'))
        column.append(self.findStr(file,sheet,'数据库'))
        column.append(self.findStr(file,sheet,'Iteration'))
        return column                                                                                                                                                                              