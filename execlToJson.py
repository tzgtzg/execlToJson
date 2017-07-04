#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
#import os.path
import json
import xlrd
import math
import types
import glob
import re
from optparse import OptionParser

fileTypeArray = [".xlsx",".xls"]
curfile = r'E:\github\execlToJson\ErrorCodeID.xlsx'

def readAllExecl(_type):
	currentPath = os.getcwd()
	#print "currentPath:" + currentPath
	#for dir in [x for x in os.listdir(CUR_PATH) if os.path.isdir(os.path.join(CUR_PATH, x))]:
	for dir in [x for x in os.listdir(currentPath)]:
		localPath = os.path.join(currentPath, dir)
		if os.path.isfile(localPath):
			filesp = os.path.splitext(localPath)				 
			for k in fileTypeArray:
				if filesp[1] == k:				
					filename = os.path.basename(localPath)
					if _type == "json":
						readExeclToJson(localPath,filename.split('.')[0])
					elif _type == "lua":
						readExeclToLua(localPath,filename.split('.')[0])
						
			#print  localPath
		#print "dir  " + dir
def readExeclToLua(path,name):
	workbook  = xlrd.open_workbook(path)
	
	sheet2_name = workbook.sheet_names()[0]
	sheet=workbook.sheet_by_name(sheet2_name)   # sheet索引从0开始
	# sheet的名称，行数，列数
	#print sheet.name,sheet.nrows,sheet.ncols
	
	adict = {}
	
	for i in range(1,sheet.nrows):
		data = {}
		#print TransformationType(sheet.cell_value(0,0))
		for j in range(0,sheet.ncols):
			 value = TransformationType(sheet.cell_value(i,j))
			 # print  value 
			 key = TransformationType(sheet.cell_value(0,j)) 
			 if  isinstance(value , str):
				
				 if isJsonString(value):
				 	# 可以转化成 json 对象的 					
					data[key] = eval(value)
				 else:
				 	# str 值
					data[key] = value
			 else:
			 	 # 非str 值
				 data[key] = value
		# 转换成 python 字典 对象			
		adict[TransformationType(sheet.cell_value(i,0))]= data

	data = json.dumps(adict,sort_keys=True,indent=1,ensure_ascii=False)
	ssdata = re.sub(r'\[','{',data,flags=re.M)
	ccdata = re.sub(r']','}',ssdata,flags=re.M)
	leftData = re.sub(r'^ *"',' ["',ccdata,flags=re.M)
	rightdata = re.sub(r':',']=',leftData,flags=re.M)
	# print rightdata
	# 添加lua 头
	# \"total_amount\":\""+ price+"\"
	modestr = " module(\"" +  name +"\")"+"\n"
	localStr = "local " + name + "tab =" + "\n"
	f=open(name+'.lua','w') 
	f.write(modestr+localStr+rightdata)
	f.close()
	print "already create  lua :  " + path



	
	

	


def readExeclToJson(path,name):
	workbook  = xlrd.open_workbook(path)
	# [u'sheet1', u'sheet2']
	#print workbook.sheet_names() 
	
	sheet2_name = workbook.sheet_names()[0]
	sheet=workbook.sheet_by_name(sheet2_name)   # sheet索引从0开始
	# sheet的名称，行数，列数
	#print sheet.name,sheet.nrows,sheet.ncols
	
	adict = {}
	mlist =[]
	for i in range(1,sheet.nrows):
		data = {}
		#print TransformationType(sheet.cell_value(0,0))
		for j in range(0,sheet.ncols):
			 value = TransformationType(sheet.cell_value(i,j))
			 #print type(value)
			 if  isinstance(value , str):
				
				 if isJsonString(value):					
					data[TransformationType(sheet.cell_value(0,j))] = eval(value)
				 else:
					data[TransformationType(sheet.cell_value(0,j))] = value
			 else:
			 	 # print TransformationType(sheet.cell_value(0,j))
				 data[TransformationType(sheet.cell_value(0,j))] = value
					
		# adict[TransformationType(sheet.cell_value(i,0))]= data
		mlist.append(data)
		adict[""+name+"config"] = mlist
	
	 
	data = json.dumps(adict,sort_keys=True,indent=1,ensure_ascii=False)
	f=open(name+'.json','w') 
	f.write(data)
	f.close()
	print "already create  json :  " + path
	
	
def isJsonString(str):
	try:
		eval(str)
	except Exception,e :
		return False     
	return True		

def TransformationType(var):
	#print  type(var)
	if isinstance(var ,float) : #type(var) == 'float':
		str1 = int(var)
	elif isinstance(var, unicode): #type(var) == 'unicode':
		 str1 = var.encode("utf-8")
	else:
		raise Exception("type  is  not  deal ")
		str1 = var
	return str1

def Usage():
    print '-h,--help: print help message.'
    print '-l,  -- lua    to lua  table    --json  to json '
   


#parser.add_option("-p", "--pdbk", action="store_true", 
 #                 dest="pdcl", 
 #                 default=False, 
 #                help="write pdbk data to oracle db") 
#add_option用来加入选项，action是有store，store_true，store_false等，dest是存储的变量，default是缺省值，help是帮助提示
	
def main():
	parser = OptionParser(usage="")
	parser.add_option("-l","--language",action="store",
	dest="languages",
	help="-l,  -- lua    to lua  table    --json  to json")

	(options,args) = parser.parse_args()
	#print options.languages
	#print args
	if options.languages == "lua" or options.languages == "json" :
		readAllExecl(options.languages)
	print "create  all success : " 
	os.system("pause")

if __name__ == "__main__":
	main()
