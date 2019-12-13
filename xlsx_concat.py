# -*- coding: utf-8 -*-
#1	将xls和csv转化为xlsx		Trans(path)
#2	将xlsx合并（按列）		xlsx_concat(path_new)
#3	1、2合并				xlsx_con()

import win32com.client as win32
import os 
import pandas as pd 
import xlrd

def help():
	help_o=pd.DataFrame(data=[['xls、csv转化xlsx','Trans(path)','win32,os '],
		['xlsx合并(按列)','xlsx_concat(path_new)','pd,os,xlrd'],
		['xls转化合并','xlsx_con','//']],
		columns=['名称','调用','库'])
	print(help_o)
    
def Trans(path):
	'''
	import win32com.client as win32
	import os 
	'''
#1	创建Transfrom-xlsx文件夹	
	path_new=path+'\\Trans'
	if os.path.exists(path_new):
		pass
	else:
		os.mkdir(path_new)
#2	获取xls、csv、xlsx文件名列表
	print('工作目录',path)
	spl_list=['.xls','.csv','.xlsx']
	filelist = []
	for root,dirs,files in os.walk(path):
	    for file in files:
 	       if os.path.splitext(file)[1] in spl_list:
 	           filelist.append(file)
	aa=len(filelist)
	print('需要转化 %d 个文档:'%(aa))
	print('文件列表',filelist)
#3	转化格式为xlsx，并保存到Transfrom_xlsx
	excel = win32.gencache.EnsureDispatch('Excel.Application')
	for filelist_new in filelist:
		fff=path1+'\\'+filelist_new
		wb = excel.Workbooks.Open(fff)
		wb.SaveAs(path_new+'\\'+filelist_new+".xlsx", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
		wb.Close()                               #FileFormat = 56 is for .xls extension
	excel.Application.Quit()
	print('转化结束')
	path2=pa	

def xlsx_concat(path_new):
	'''
	import pandas as pd
	import os 
	import xlrd
	'''
#1	获取xlsx文件名称列表
	print('合并目录',path_new)
	filelist = []
	spl_list=['.xlsx']
	for root,dirs,files in os.walk(path_new):
	    for file in files:
	        if os.path.splitext(file)[1] in spl_list:
 	           filelist.append(file)
	aa=len(filelist)
	print('需合并 %d 个文档:'%(aa))
	print('文件列表',filelist)
#2	获取表及sheet名称列表	
	Sheetlist=[]
	datalist = []
	for num_file in range(aa):
		file_n = path_new + '/' + filelist[num_file]
		wb =xlrd.open_workbook(file_n)
		sheet_name_1= wb.sheet_names()
		sheetsum= len(sheet_name_1)
#3	将excel数据以Datafram装入datalist列表
		for sheet_new in sheet_name_1:
			datalist.append(pd.read_excel(file_n,sheet_name=sheet_new))
			data_last=datalist[-1]
			col_name=data_last.columns.tolist()
			col_name.insert(0,'来源')
			data_last=data_n.reindex(columns=col_name)
			data_last['来源'] = str(filelist[num_file])+'@'+str(sheet_new)
#4	将datalist的所有DataFrame按列合并
	data = pd.concat(datalist,axis=0,sort=False)
	outwriter = path_new  + '/' + '@excel_output.xls'
	data.to_excel(outerwriter)
	print('已输出到Trans/@excel_output.xls')
	print(data.shape)

def xlsx_con（）:
	path=os.getcwd()
	path_new=path+'\\Trans'
	print('=====第一步=====')
	Trans(path)
	print('=====第二部=====')
	xls_concat(path_new)
	print('=====结束=====')
