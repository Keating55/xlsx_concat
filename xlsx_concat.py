# -*- coding: utf-8 -*-
import win32com.client as win32
import os
import pandas as pd
import xlrd

# 1.创建工作目录

dir_path = os.getcwd()
new_path = dir_path + '\\_Workdir\\Data'
if os.path.exists(new_path):
    pass
else:
    os.mkdirs(new_path)

# 2.读取文件名

file_list = []
file_name_list = []
spl_list = ['.xlsx', '.csv', '.xls']
for root, dirs, files in os.walk(new_path):
    for file in files:
        if os.path.splitext(file)[1] in spl_list:
            file_list.append(file)
            file_name_list.append(os.path.splitext(file)[0])
file_sum = len(file_name_list)
print('共有%d个文件')%(file_sum)
print('正在格式化')

# 3.文件格式化

excel = win32.gencache.EnsureDispatch('Excel.Application')
for file_n in file_list:
    file_path = dir_path + '\\' + file_n
    wb = excel.Workbooks.Open(file_path)
    file_name = os.path.splitext(file_n)[0]
    wb.SaveAs(new_path + '\\' + file_name + ".xlsx",
              FileFormat=51)  #FileFormat = 51 is for .xlsx extension
    wb.Close()  #FileFormat = 56 is for .xls extension
excel.Application.Quit()
print('格式化结束')

# 4.合并文件
print('开始合并文件')

Sheet_list = []
data_list = []
for n in range(file_sum):
    file_n = new_path + '/' + file_name_list[n] + '.xlsx'
    wb = xlrd.open_workbook(file_n)
    sheet_name_list = wb.sheet_names()
    sheet_sum = len(sheet_name_list)
    
    for sheet_n in range(sheet_sum):
        data_list.append(
            pd.read_excel(file_n, sheet_name=sheet_name_list[sheet_n]))
        data_last = data_list[-1]
        col_name = data_last.columns.tolist()
        col_name.insert(0, '来源')
        data_last = data_last.reindex(columns=col_name)
        data_last['来源'] = str(n) + '-' + str(sheet_n + 1) + str(
            file_name_list[n]) + '@' + str(sheet_new)

data = pd.concat(data_list, axis=0, sort=False)
output_file = dir_path + '\\_Workdir\\@excel_output.xlsx'
data.to_excel(output_file)
print('合并结束，文件输出到%s')%(output_file)
print('datashape:',data.shape)
