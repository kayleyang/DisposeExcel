# -*- coding: UTF-8 -*-

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import os,re,time
import xlrd
import xlwt
import ConfigParser

class filterExcel(object):
    def __init__(self):
        
        self.inputDir = "."
        self.outputDir = "."
        self.inputFileName = ""
        self.outputFileName = u"刘汝采购.xls"

        self.filterTitle = [ u"商品编号", u"上下柜状态", u"商品名称", u"三级分类名称", u"品牌名称", u"采购员名称", u"全国库存金额", u"全国现货", u"全国预定", u"全国实际库存", u"周转", u"全国昨日销量", u"全国7日销量", u"全国15日销量", u"全国30日销量", u"全国30日销售额"]
        
        self.filterBuyerTitle=u'采购员名称'
        self.filterBuyerName=[u'刘汝']

        self.filterBrandTitle=u'品牌名称'
        self.filterBrandName=[u'华帝',u'亿田']

        self.filterCountTitle=u'实际库存'
        self.filterCountName=[]

        self.conf = ConfigParser.ConfigParser()
        self.data = []

    def initExcel(self):
        # if os.path.
        excelList = testObject.initFileList('.', [], ['xlsx', 'xls'])
        if len(excelList) <=0:
            os.system('cls')
            print u'没有找到要处理的Excel文件，确定配置文件中inputDir所指目录中包含.xls和.xlsx文件'
            print u"0. 退出好了"
            print u"1. 等我创建好了再重试一下吧"
            while True:
                choosed = int(raw_input())
                if choosed < 2:
                    sys.exit()
        elif len(excelList) == 1:
            print u'发现了一个Excel文件：',excelList[0]

        else :
            print u'发现了很多Excel文件，你选哪个呢？'
            for i, excel in enumerate(excelList):
                print i, '. ',excel
            while True:
                choosed = int(raw_input())
                if choosed < len(excelList):
                    self.readExcel(excelList[0])
                    self.execExcel()
                    break
                else :
                    print u'请重新输入:',



    def readExcel(self):
        os.system('cls')
        print u"正在读取",self.inputFileName,u"文件"
        excel = xlrd.open_workbook(self.inputFileName)
        sheet = excel.sheet_by_name(u'Sheet1')
        nrows= sheet.nrows
        ncols = sheet.ncols
        self.data.append(self.filterTitle)
        for i in range(nrows):
            row = []
            for j in range(ncols):
                row.append(sheet.cell(i,j).value)
            print u"正在读取第",str(i),u"行"
            if len(row) !=0 :
                self.data.append(row)

    def execExcel(self):
        nrows= sheet.nrows
        ncols = sheet.ncols
        self.data.append(self.filterTitle)
        titleLen = len(self.filterTitle)
        for k in range(titleLen,ncols):
            if self.filterCountTitle in sheet.cell(0,k).value:
                self.data[0].append(sheet.cell(0,k).value)
        for i in range(nrows):
            row = []
            for j in range(ncols):
                if sheet.cell(i,4).value in self.filterBrandName:
                    if sheet.cell(i,5).value in self.filterBuyerName:
                        if j < titleLen:
                            row.append(sheet.cell(i,j).value)
                        elif self.filterCountTitle in sheet.cell(0,j).value:
                            row.append(sheet.cell(i,j).value)
            print u"正在读取第",str(i),u"行"
            if len(row) !=0 :
                self.data.append(row)
        self.writeExcel(self.data)

    def writeExcel(self, data):
        excel = xlwt.Workbook()
        sheet = excel.add_sheet(u'Sheet1')
        if os.path.exists(self.outputFileName):
            os.remove(self.outputFileName)
        for i, row in enumerate(data):
            # print row
            for j, cell in enumerate(row):
                # print cell
                sheet.write(i, j, cell)
        excel.save(self.outputFileName)

    def readConf(self, path):
        content = open('config.ini').read()
        #Window下用记事本打开配置文件并修改保存后，编码为UNICODE或UTF-8的文件的文件头  
        #会被相应的加上\xff\xfe（\xff\xfe）或\xef\xbb\xbf，然后再传递给ConfigParser解析的时候会出错  
        #，因此解析之前，先替换掉  
        content = re.sub(r"\xfe\xff","", content)  
        content = re.sub(r"\xff\xfe","", content)  
        content = re.sub(r"\xef\xbb\xbf","", content)  
        open('config.ini', 'w').write(content)  

        self.conf.read("config.ini")
        # options = self.conf.options(u"路径配置".encode('utf-8'))
        # print 'options:', options
        # items = self.conf.items(u"路径配置".encode('utf-8'))
        # print 'items:', items
    
    def initFileList(self, dir, fileList, fileType):
        newDir = dir
        # print dir.split('.').pop()
        if os.path.isfile(dir) and (dir.split('.').pop() in fileType):
            # print dir,'is file.'
            fileList.append(dir.decode('gbk'))
        elif os.path.isdir(dir):
            # print dir, 'is dir.'
            for d in os.listdir(dir):
               newDir = os.path.join(dir,d) 
               self.initFileList(newDir, fileList, fileType)
        return fileList

    def initConf(self):
        iniList = testObject.initFileList('.', [], ['ini'])
        if len(iniList) <=0:
            print u"没有找到配置文件啊!!!!创建配置文件（如：config.ini）后继续或者使用默认配置。。。"
            print u"0. 退出好了"
            print u"1. 等我创建好了再重试一下吧"
            print u"2. 我想看一下默认配置是什么样的"
            while True:
                choosed = int(raw_input())
                if choosed < 2:
                    sys.exit()
                elif choosed == 2:
                    os.system('cls')
                    self.createSysConf()
                    self.printSysConf()
                    print u"0. 退出好了"
                    print u"1. 不行，我得改改"
                    print u"2. 好的，就选它了"
                    print u"3. 好的，就选它了，保存起来"
                    while True:
                        choosed0 = int(raw_input())
                        if choosed0 < 2:
                            sys.exit()
                        elif choosed0 == 2:
                            break
                        elif choosed0 == 3:
                            self.saveSysConf()
                            break
                        else :
                            print u'请重新输入:',
                    break
                else :
                    print u'请重新输入:',
            self.loadConf()
            # TODO
        elif len(iniList) == 1:
            print u'发现了一个配置文件：',iniList[0]
            self.readConf(iniList[0])
            self.loadConf()
            # TODO
        else :
            print u'发现了很多配置文件，哪个是你的呢？请选一个吧'
            for i, ini in enumerate(iniList):
                print i, '. ',ini
            while True:
                choosed = int(raw_input())
                if choosed < len(iniList):
                    self.readConf(iniList[0])
                    self.loadConf()
                    break
                else :
                    print u'请重新输入:',
            # TODO

    def createSysConf(self):
        self.conf.add_section(u"路径配置".encode('utf-8'))
        self.conf.set(u"路径配置".encode('utf-8'), 'inputDir', '.')
        self.conf.set(u"路径配置".encode('utf-8'), 'outputDir', '.')
        self.conf.set(u"路径配置".encode('utf-8'), 'inputFileName', '.')
        self.conf.set(u"路径配置".encode('utf-8'), 'outputFileName', '.')

        self.conf.add_section(u"过滤配置".encode('utf-8'))
        self.conf.set(u"过滤配置".encode('utf-8'), 'filterTitle', self.listToString(self.filterTitle))
        self.conf.set(u"过滤配置".encode('utf-8'), 'filterBuyerTitle', self.filterBuyerTitle)
        self.conf.set(u"过滤配置".encode('utf-8'), 'filterBuyerName', self.listToString(self.filterBuyerName))
        self.conf.set(u"过滤配置".encode('utf-8'), 'filterBrandTitle', self.filterBrandTitle)
        self.conf.set(u"过滤配置".encode('utf-8'), 'filterBrandName', self.listToString(self.filterBrandName))
        self.conf.set(u"过滤配置".encode('utf-8'), 'filterCountTitle', self.filterCountTitle)
        self.conf.set(u"过滤配置".encode('utf-8'), 'filterCountName', self.listToString(self.filterCountName))

    def loadConf(self):
        self.inputDir = self.conf.get(u'路径配置'.encode('utf-8'), 'inputDir')
        self.outputDir = self.conf.get(u'路径配置'.encode('utf-8'), 'outputDir')
        self.inputFileName = self.conf.get(u'路径配置'.encode('utf-8'), 'inputFileName')
        self.outputFileName = self.conf.get(u'路径配置'.encode('utf-8'), 'outputFileName')

        self.filterTitle = self.stringToList(self.conf.get(u"过滤配置".encode('utf-8'), 'filterTitle'))
        self.filterBuyerTitle = self.conf.get(u"过滤配置".encode('utf-8'), 'filterBuyerTitle')     
        self.filterBuyerName = self.stringToList(self.conf.get(u"过滤配置".encode('utf-8'), 'filterBuyerName'))
        self.filterBrandTitle = self.conf.get(u"过滤配置".encode('utf-8'), 'filterBrandTitle')
        self.filterBrandName = self.stringToList(self.conf.get(u"过滤配置".encode('utf-8'), 'filterBrandName'))
        self.filterCountTitle = self.conf.get(u"过滤配置".encode('utf-8'), 'filterCountTitle')
        self.filterCountName = self.stringToList(self.conf.get(u"过滤配置".encode('utf-8'), 'filterCountName'))  

    def printSysConf(self):
        print u'[路径配置]'
        print u'inputDir = ',self.conf.get(u'路径配置'.encode('utf-8'), 'inputDir')
        print u'outputDir = ',self.conf.get(u'路径配置'.encode('utf-8'), 'outputDir')
        print u'[过滤配置]'
        print u'filterTitle = ',self.conf.get(u"过滤配置".encode('utf-8'), 'filterTitle')
        print u'filterBuyerTitle = ',self.conf.get(u"过滤配置".encode('utf-8'), 'filterBuyerTitle')     
        print u'filterBuyerName = ',self.conf.get(u"过滤配置".encode('utf-8'), 'filterBuyerName')
        print u'filterBrandTitle = ',self.conf.get(u"过滤配置".encode('utf-8'), 'filterBrandTitle')
        print u'filterBrandName = ',self.conf.get(u"过滤配置".encode('utf-8'), 'filterBrandName')
        print u'filterCountTitle = ',self.conf.get(u"过滤配置".encode('utf-8'), 'filterCountTitle')
        print u'filterCountName = ',self.conf.get(u"过滤配置".encode('utf-8'), 'filterCountName')
        print ''

    def saveSysConf(self):
        self.conf.write(open("config.ini", "w"))

    def listToString(self, list):
        string = ''
        for i,s in enumerate(list):
            string+= s
            if i < len(list)-1:
                string+=','
        return string

    def stringToList(self, string):
        return string.strip().spilt(',')


if __name__=="__main__":
    startTime = time.time()
    testObject=filterExcel()
    # testObject.initConf()
    testObject.initExcel()
    # testObject.readExcel()
    # testObject.readConf()
    # testObject.createConf()

    stopTime = time.time()
    timeTaken = stopTime - startTime
    print '\nruning cost '+ ("%.2f" % timeTaken) +' s'
        