#! /usr/bin/env python3
# coding=utf-8

import sys
from datetime import date
from xlrd import open_workbook,xldate_as_tuple
from xlwt import Workbook

input_file = sys.argv[1]
output_file = sys.argv[2]

# 实例化Workbook 对象，使可以将结果写入用于输出的Excel 文件
output_workbook = Workbook()

# add_sheet函数 为输出工作簿添加一个工作表
output_worksheet = output_workbook.add_sheet('')

# open_workbook: 打开输入工作簿 赋值给workbook
with open_workbook(input_file) as workbook:

    # sheet_by_name:引用工作簿
    worksheet = workbook.sheet_by_name('')

    # 创建行 与 列 索引值上的for循环语句，使用range函数 和 nrows，ncols 属性，在工作簿每行和每列直接迭代
    # range(10)        # 从 0 开始到 10
    # [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
    for row_index in range(worksheet.nrows):
        row_list_output = []
        for col_index in range(worksheet.ncols):
            # cell_type：返回单元格中的数据类型
            # 3：日期数据
            if worksheet.cell_type(row_index,col_index) == 3:
                # xldate_as_tuple 将Excel中代表日期，时间的数值转成元组
                # datemode：使函数确定日期是基于1900年还是1904年，并据此将数值转换成正确的元组
                date_cell = xldate_as_tuple(worksheet.cell_value(row_index,col_index),workbook.datemode)
                # strftime ：将data对象转换为一个具有特定格式的字符串
                date_cell = date(*date_cell[0:3]).strftime('%m/%d/%Y')
                # append ：将值追加给输出列表
                row_list_output.append(date_cell)
                # write：将每个单元格的值写入输出文件的工作表
                output_worksheet.write(row_index,col_index,date_cell)
            else:
                non_date_cell = worksheet.cell_value(row_index,col_index)
                # append ：将值追加给输出列表
                row_list_output.append(non_date_cell)
                output_worksheet.write(row_index,col_index,non_date_cell)

output_workbook.save(output_file)






