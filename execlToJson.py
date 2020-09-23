#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
reload(sys)
sys.setdefaultencoding('utf8')
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
					elif _type == "js":
						readExeclToJs(localPath,filename.split('.')[0])

def ExeclToJson(path,name):
	workbook  = xlrd.open_workbook(path)
	# [u'sheet1', u'sheet2']
	#print workbook.sheet_names() 
	adict = {}
	
	for k in workbook.sheet_names():
		sheet=workbook.sheet_by_name(k)   # sheet索引从0开始
		# sheet的名称，行数，列数
		#print sheet.name,sheet.nrows,sheet.ncols
	 	mlist =[]
 		adict[k] = {}

 		reservedNum = 1;   #  预留 行数  第一行  预留填备注
		# for i in range(1,sheet.nrows):     第一行  预留填备注
		for i in range(1 + reservedNum,sheet.nrows):  
			data = {}
			#print TransformationType(sheet.cell_value(0,0))
			for j in range(0,sheet.ncols):
				 value = TransformationType(sheet.cell_value(i,j))
				 #print type(value)
				 if  isinstance(value , str):
					
					 if isJsonString(value):					
						data[TransformationType(sheet.cell_value(0+reservedNum,j))] = eval(value)
					 else:
						data[TransformationType(sheet.cell_value(0+reservedNum,j))] = value
				 else:
				 	 # print TransformationType(sheet.cell_value(0,j))
					 data[TransformationType(sheet.cell_value(0+reservedNum,j))] = value
						
			# adict[TransformationType(sheet.cell_value(i,0))]= data
			# mlist.append(data)
			
			# adict[""+k] = mlist
			adict[""+k][TransformationType(sheet.cell_value(i,0))] = data

	
	data = json.dumps(adict,sort_keys=True,indent=1,ensure_ascii=False)
	return data

def readExeclToJs(path,name):
	data = ExeclToJson(path,name)
	moduleStr = "module.exports = " 
	newData = moduleStr + data
	f=open(name+'.js','w') 
	f.write(newData)
	f.close()
	print "already create  js :  " + path

def readExeclToLua(path,name):
	
	data = ExeclToJson(path,name)


	 # 转换成  数组 格式 时  打开 注释 
	# ssdata = re.sub(r'\[','{',data,flags=re.M)
	# ccdata = re.sub(r']','}',ssdata,flags=re.M)
	rightdata = re.sub(r'(")(\w+)(":)','["'+r'\2'+'"]=',data,flags=re.M) 
	
	# print rightdata
	# 添加lua 头
	# \"total_amount\":\""+ price+"\"
	
	localStr = "local " + name + "tab =" + "\n"
	returnstr = "return " + name+"tab"
	f=open(name+'.lua','w') 
	f.write(localStr + rightdata +  "\n" + returnstr)
	f.close()
	print "already create  lua :  " + path


def readExeclToJson(path,name):
	data = ExeclToJson(path,name)
	f=open(name+'.json','w') 
	f.write(data)
	f.close()
	print "already create  json :  " + path
	# [{"size":20,"text":"1.用牌","color":"255,255,0"},{"size":20, "text":"湖南麻将中使用的牌为万字、筒字与条字各36张，共10张","color":"255,255,0"}]
	
def isJsonString(str):
	try:
		eval(str)
	except Exception,e :
		return False     
	return True		

def TransformationType(var):
	#print  type(var)
	if isinstance(var ,float) : #type(var) == 'float':
		# print("var  ", var)
		if var == int(var) : 
			str1 = int(var)
		else:
			str1 = round(var,2)
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
	help="-l,  -- lua    to lua  table file    --json  to json  --js  to js file")

	(options,args) = parser.parse_args()
	#print options.languages
	#print args
	if options.languages == "lua" or options.languages == "json" or options.languages == "js":
		readAllExecl(options.languages)
	print "create  all success : " 
	os.system("pause")

if __name__ == "__main__":
	main()
