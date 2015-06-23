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
        if os.path.exists(self.inputFileName):
            self.readExcel(self.inputFileName)
            self.execExcel()
        else:
            print u'没有找到要处理的Excel文件，确定配置文件中inputFileName所指文件存在'
            excelList = testObject.initFileList('.', [], ['xlsx', 'xls'])
            if len(excelList) <=0:
                os.system('cls')
                print u'配置的路径下也没有找到要处理的Excel文件，确定配置文件中inputDir所指目录中包含.xls和.xlsx文件'
                print u"0. 退出好了"
                print u"1. 等我创建好了再重试一下吧"
                while True:
                    choosed = int(raw_input())
                    if choosed < 2:
                        sys.exit()
            elif len(excelList) == 1:
                print u'但是发现了另一个Excel文件：',excelList[0]
                print u'是否处理该文件？'
                print u"0. 退出好了"
                print u"1. 就这个文件了"
                while True:
                    choosed0 = int(raw_input())
                    if choosed0 <= 0:
                        sys.exit()
                    elif choosed0 == 1:
                        self.readExcel(excelList[0])
                        self.execExcel()
                        break
                    else :
                        print u'请重新输入:',
            else :
                print u'但是发现了很多Excel文件，你选哪个呢？'
                for i, excel in enumerate(excelList):
                    print i, '. ',excel
                while True:
                    choosed = int(raw_input())
                    if choosed < len(excelList):
                        self.readExcel(excelList[choosed])
                        self.execExcel()
                        break
                    else :
                        print u'请重新输入:',

    def filterControl(self):
        self.execExcel()
        print u'执行完毕！！！！'
        time.sleep(2)
        print u'是否继续按采购员和品牌分别过滤？'
        for buyer in self.filterBuyerName:
            execExcel()

    def readExcel(self, fileName):
        os.system('cls')
        print u"正在读取",fileName,u"文件"
        excel = xlrd.open_workbook(fileName)
        sheet = excel.sheet_by_name(u'Sheet1')
        nrows= sheet.nrows
        ncols = sheet.ncols
        for i in range(nrows):
            row = []
            for j in range(ncols):
                row.append(sheet.cell(i,j).value)
            print u"正在读取第",str(i),u"行"
            if len(row) !=0 :
                self.data.append(row)

    def execExcel(self):
        data = []
        nrows= len(self.data)
        ncols = len(self.data[0])
        data.append(self.filterTitle)
        titleLen = len(self.filterTitle)
        for k in range(titleLen,ncols):
            if (self.filterCountTitle in self.data[0][k]) and ((len(self.filterCountName) <= 0) or (self.data[0][k] in self.filterCountName)):
                data[0].append(self.data[0][k])
        fullTitle = self.data[0]

        indexOfBrandTitle = 0
        for i, title in enumerate(fullTitle):
            if title == self.filterBrandTitle:
                indexOfBrandTitle = i
                break
        # print indexOfBrandTitle

        indexOfBuyerTitle = 0
        for i, title in enumerate(fullTitle):
            if title == self.filterBuyerTitle:
                indexOfBuyerTitle = i
                break
        # print indexOfBuyerTitle

        for i in range(1,nrows):
            row = []
            for j in range(ncols):
                if (len(self.filterBrandName) <= 0) or (self.data[i][indexOfBrandTitle] in self.filterBrandName):
                    if (len(self.filterBuyerName) <=0) or (self.data[i][indexOfBuyerTitle] in self.filterBuyerName):
                        if j < titleLen:
                            row.append(self.data[i][j])
                        elif self.filterCountTitle in self.data[0][j] and ((len(self.filterCountName) <= 0) or (self.data[0][j] in self.filterCountName)):
                            row.append(self.data[i][j])
            # print u"正在读取第",str(i),u"行"
            if len(row) !=0 :
                data.append(row)
        self.writeExcel(data)

    def writeExcel(self, data):
        excel = xlwt.Workbook()
        sheet = excel.add_sheet(u'Sheet1')
        if os.path.exists(self.outputFileName):
            os.remove(self.outputFileName)
        # print data
        for i, row in enumerate(data):
            # print row
            for j, cell in enumerate(row):
                # print cell
                sheet.write(i, j, cell)
        excel.save(self.outputFileName)

    def readConf(self, path):
        content = open(path).read()
        #Window下用记事本打开配置文件并修改保存后，编码为UNICODE或UTF-8的文件的文件头  
        #会被相应的加上\xff\xfe（\xff\xfe）或\xef\xbb\xbf，然后再传递给ConfigParser解析的时候会出错  
        #，因此解析之前，先替换掉  
        content = re.sub(r"\xfe\xff","", content)  
        content = re.sub(r"\xff\xfe","", content)  
        content = re.sub(r"\xef\xbb\xbf","", content)  
        open(path, 'w').write(content)  

        self.conf.read(path)
        # options = self.conf.options('pathConfig')
        # print 'options:', options
        # items = self.conf.items('pathConfig')
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
            print u'发现了很多配置文件，请选一个吧'
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
        self.conf.add_section('pathConfig')
        self.conf.set('pathConfig', 'inputDir', '.')
        self.conf.set('pathConfig', 'outputDir', '.')
        self.conf.set('pathConfig', 'inputFileName', '.')
        self.conf.set('pathConfig', 'outputFileName', '.')

        self.conf.add_section('filterConfig')
        self.conf.set('filterConfig', 'filterTitle', self.listToString(self.filterTitle))
        self.conf.set('filterConfig', 'filterBuyerTitle', self.filterBuyerTitle)
        self.conf.set('filterConfig', 'filterBuyerName', self.listToString(self.filterBuyerName))
        self.conf.set('filterConfig', 'filterBrandTitle', self.filterBrandTitle)
        self.conf.set('filterConfig', 'filterBrandName', self.listToString(self.filterBrandName))
        self.conf.set('filterConfig', 'filterCountTitle', self.filterCountTitle)
        self.conf.set('filterConfig', 'filterCountName', self.listToString(self.filterCountName))

    def loadConf(self):
        self.inputDir = self.conf.get('pathConfig', 'inputDir')
        self.outputDir = self.conf.get('pathConfig', 'outputDir')
        self.inputFileName = self.conf.get('pathConfig', 'inputFileName')
        self.outputFileName = self.conf.get('pathConfig', 'outputFileName')

        self.filterTitle = self.stringToList(self.conf.get('filterConfig', 'filterTitle'))
        self.filterBuyerTitle = self.conf.get('filterConfig', 'filterBuyerTitle')     
        self.filterBuyerName = self.stringToList(self.conf.get('filterConfig', 'filterBuyerName'))
        self.filterBrandTitle = self.conf.get('filterConfig', 'filterBrandTitle')
        self.filterBrandName = self.stringToList(self.conf.get('filterConfig', 'filterBrandName'))
        self.filterCountTitle = self.conf.get('filterConfig', 'filterCountTitle')
        self.filterCountName = self.stringToList(self.conf.get('filterConfig', 'filterCountName'))  

    def printSysConf(self):
        print u'[pathConfig]'
        print u'inputDir = ',self.conf.get('pathConfig', 'inputDir')
        print u'outputDir = ',self.conf.get('pathConfig', 'outputDir')
        print u'inputFileName = ',self.conf.get('pathConfig', 'inputFileName')
        print u'outputFileName = ',self.conf.get('pathConfig', 'outputFileName')
        print u'[filterConfig]'
        print u'filterTitle = ',self.conf.get('filterConfig', 'filterTitle')
        print u'filterBuyerTitle = ',self.conf.get('filterConfig', 'filterBuyerTitle')     
        print u'filterBuyerName = ',self.conf.get('filterConfig', 'filterBuyerName')
        print u'filterBrandTitle = ',self.conf.get('filterConfig', 'filterBrandTitle')
        print u'filterBrandName = ',self.conf.get('filterConfig', 'filterBrandName')
        print u'filterCountTitle = ',self.conf.get('filterConfig', 'filterCountTitle')
        print u'filterCountName = ',self.conf.get('filterConfig', 'filterCountName')
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
        return string.strip().split(',')


if __name__=="__main__":
    # startTime = time.time()
    testObject=filterExcel()
    testObject.initConf()
    testObject.initExcel()
    # testObject.readExcel()
    # testObject.readConf()
    # testObject.createConf()

    # stopTime = time.time()
    # timeTaken = stopTime - startTime
    # print '\nruning cost '+ ("%.2f" % timeTaken) +' s'
        