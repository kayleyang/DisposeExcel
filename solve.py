# -*- coding: UTF-8 -*-

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import xlrd
import os
import xlwt
import time

class filterExcel(object):
	def __init__(self):
		self.inputFileName = u"厨卫库存6-11.xlsx"
		self.outputFileName = u"刘汝采购.xlsx"
		self.filterTitle = [ u"商品编号", u"上下柜状态", u"商品名称", u"三级分类名称", u"品牌名称", u"采购员名称", u"全国库存金额", u"全国现货", u"全国预定", u"全国实际库存", u"周转", u"全国昨日销量", u"全国7日销量", u"全国15日销量", u"全国30日销量", u"全国30日销售额"]
		
		self.filterBuyerKeywords=u'采购员名称'
		self.filterBuyerName=[u'刘汝']

		self.filterBrandKeywords=u'品牌名称'
		self.filterBrandName=[u'华帝',u'亿田']

		self.filterCountKeywords=u'实际库存'
		self.filterCountCity=u'all'

	def readExcel(self):
		print u"正在读取",self.inputFileName,u"文件"
		excel = xlrd.open_workbook(self.inputFileName)
		sheet = excel.sheet_by_name(u'Sheet1')
		nrows= sheet.nrows
		ncols = sheet.ncols
		data = []
		data.append(self.filterTitle)
		for k in range(len(self.filterTitle),ncols):
			if self.filterCountKeywords in sheet.cell(0,k).value:
				data[0].append(sheet.cell(0,k).value)
		for i in range(nrows):
			row = []
			for j in range(ncols):
				if sheet.cell(i,4).value in self.filterBrandName:
					if sheet.cell(i,5).value in self.filterBuyerName:
						if j < len(self.filterTitle):
							row.append(sheet.cell(i,j).value)
						else:
							if self.filterCountKeywords in sheet.cell(0,j).value:
								row.append(sheet.cell(i,j).value)
			print u"正在读取第",str(i),u"行"
			if len(row) !=0 :
				data.append(row)
		self.writeExcel(data)

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

if __name__=="__main__":
	startTime = time.time()
	testObject=filterExcel()
	testObject.readExcel()
	stopTime = time.time()
	timeTaken = stopTime - startTime
	print '\nruning cost '+ ("%.2f" % timeTaken) +' s'
		