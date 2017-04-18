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
fileTypeArray = [".xlsx",".xls"]
curfile = r'E:\github\execlToJson\ErrorCodeID.xlsx'



def readAllExecl():
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
					
					readExecl(localPath,filename.split('.')[0])
			
			#print  localPath
		#print "dir  " + dir
def readExecl(path,name):

	workbook  = xlrd.open_workbook(path)
	# [u'sheet1', u'sheet2']
	#print workbook.sheet_names() 
	
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
			 #print type(value)
			 if  isinstance(value , str):
				
				 if isJsonString(value):					
					data[TransformationType(sheet.cell_value(0,j))] = eval(value)
				 else:
					data[TransformationType(sheet.cell_value(0,j))] = value
			 else:
				 data[TransformationType(sheet.cell_value(0,j))] = value
					
		adict[TransformationType(sheet.cell_value(i,0))]= data
	
	 
	data = json.dumps(adict,indent=1)
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
	
def main():
	readAllExecl()
	print "create  all success : " 
	os.system("pause")

if __name__ == "__main__":
	main()
