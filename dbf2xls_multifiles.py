#!/usr/bin/python
# -*- coding:utf-8 -*-

import os
from dbfpy import dbf
import xlwt

def main():
	#获取当前路径下全部dbf文件名
	fpath = os.path.abspath('.')
	files = os.listdir(fpath)
	dbfiles=[]
	for fname in files:
		extname = os.path.splitext(fname)
		if extname[1] == '.dbf' or extname[1] == '.DBF':
			dbfiles.append(fname)
	if dbfiles == []:
		print ('Empty!')
		exit()
	
	#dbf2xls
	for dbfile in dbfiles:
		fullpath = fpath+os.sep+dbfile
		db = dbf.Dbf(fullpath)
		exportname=fullpath[0:-3]+'xls' #xls文件与原dbf文件同名，保存在相同的路径中
		print dbfile,
		
		#建立Excel工作簿工作表
		book = xlwt.Workbook(encoding='gbk')
		sheet = book.add_sheet('dbf2xls')
		
		#Excel工作表第一行写入dbf字段名
		c = 0
		r = 0
		for fldnm in db.fieldNames:
			sheet.write(0,c,fldnm)
			c = c + 1
		
		#将dbf各项数据写入Excel
		c = 0
		r = 1
		for rec in db:
			for col in rec:
				if isinstance(col,str):
					col = col.decode('gbk','ignore')
				sheet.write(r,c,col)
				c = c + 1
			r = r + 1
			c = 0
		book.save(exportname)
		print '...OK'
	raw_input ('Done. Press any key to continue.')
		
		
if __name__ == '__main__':
	main()

