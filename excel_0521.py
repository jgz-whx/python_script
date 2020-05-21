#! /usr/bin/env python3
# coding=utf-8

"""
行：
nrows = table.nrows  #获取该sheet中的有效行数

table.row(rowx)  #返回由该行中所有的单元格对象组成的列表

table.row_slice(rowx)  #返回由该列中所有的单元格对象组成的列表

table.row_types(rowx, start_colx=0, end_colx=None)    #返回由该行中所有单元格的数据类型组成的列表

table.row_values(rowx, start_colx=0, end_colx=None)   #返回由该行中所有单元格的数据组成的列表

table.row_len(rowx) #返回该列的有效单元格长度


列：
ncols = table.ncols   #获取列表的有效列数

table.col(colx, start_rowx=0, end_rowx=None)  #返回由该列中所有的单元格对象组成的列表

table.col_slice(colx, start_rowx=0, end_rowx=None)  #返回由该列中所有的单元格对象组成的列表

table.col_types(colx, start_rowx=0, end_rowx=None)    #返回由该列中所有单元格的数据类型组成的列表

table.col_values(colx, start_rowx=0, end_rowx=None)   #返回由该列中所有单元格的数据组成的列表

单元格：
table.cell(rowx,colx)   #返回单元格对象

table.cell_type(rowx,colx)    #返回单元格中的数据类型

table.cell_value(rowx,colx)   #返回单元格中的数据

table.cell_xf_index(rowx, colx)   # 暂时还没有搞懂
"""





import sys
from xlrd import open_workbook
from xlwt import Workbook
input_file = sys.argv[1]
output_file = sys.argv[2]

# 实例化Workbook 对象，使可以将结果写入用于输出的Excel 文件
output_workbook = Workbook()

# add_sheet函数 为输出工作簿添加一个工作表
output_worksheet = output_workbook.add_sheet('whx_0521')

# open_workbook: 打开输入工作簿 赋值给workbook
with open_workbook(input_file) as workbook :

    # sheet_by_name:引用工作簿
    worksheet = workbook.sheet_by_name('whx')

    # 创建行 与 列 索引值上的for循环语句，使用range函数 和 nrows，ncols 属性，在工作簿每行和每列直接迭代
    # range(10)        # 从 0 开始到 10
    # [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    for row_index in range(worksheet.nrows) :
        for column_index in range(worksheet.ncols) :

            # write：将每个单元格的值写入输出文件的工作表
            # cell_value : 返回单元格中的数据
            output_worksheet.write(row_index,column_index,worksheet.cell_value(row_index,column_index))
            print(row_index,column_index,worksheet.cell_value(row_index,column_index))
output_workbook.save(output_file)



