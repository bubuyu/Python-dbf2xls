#!/usr/bin/python
# -*- coding:utf-8 -*-
# dbf2xls转换器
# 基于dbfpy xlwt
# 由dbf2csv修改而来 (https://github.com/borrob/dbf2csv)

from dbfpy import dbf
import xlwt

def main():
	#原始DBF文件名（含路径）
	filein = raw_input('File Path: ')
	db=dbf.Dbf(filein)
	filename=filein[0:-3]+'xls' #xls文件与原dbf文件同名，保存在相同的路径中

	
	#建立Excel工作簿工作表
	book = xlwt.Workbook(encoding='gbk')
	sheet = book.add_sheet('dbf_convertor')
	c = 0
	r = 0
	#Excel工作表第一行写入dbf字段名
	for fldnm in db.fieldNames:
		sheet.write(0,c,fldnm)
		c = c + 1
	c = 0
	r = 1
	#将dbf各项数据写入Excel
	for rec in db:
		for col in rec:
			sheet.write(r,c,col)
			c = c + 1
		r = r + 1
		c = 0
	book.save(filename)
	print('Done.')

if __name__ == '__main__':
	main()
