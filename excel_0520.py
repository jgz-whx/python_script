#!/usr/bin/env python3
import sys
from xlrd import open_workbook
input_file = sys.argv[1]
# 加开Excel输入文件
workbook = open_workbook(input_file)
# 打印出工作簿中工作表数量
print("Number of worksheetd:",workbook.nsheets)
# workbook.sheets ：识别出工作簿所有工作表
for worksheet in workbook.sheets():
    print("Worksheet name:",worksheet.name,"\tRows:",worksheet.nrows,"\tColumns:",worksheet.ncols)
    print("Worksheet name:", worksheet.name)
    print("\tRows:",worksheet.nrows)
    print("\tColumns:",worksheet.ncols)
