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


sale_amount_column_index = 3
# open_workbook: 打开输入工作簿 赋值给workbook
with open_workbook(input_file) as workbook:
    # sheet_by_name:引用工作簿
    worksheet = workbook.sheet_by_name('')

    # 创建空列表data,将用输入文件中要写入输出文件中的那些行来填充这个列表
    data = []
    # 提取标题行 ,直接插入data列表中
    header = worksheet.row_value(0)
    # append ：将值追加给输出列表
    data.append(header)

    #创建行 与 列 索引值上的for循环语句，使用range函数 和 nrows，ncols 属性，在工作簿每行和每列直接迭代
    #>> > range(1, 11)  # 从 1 开始到 11
    #[1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
    for row_index in range(1,worksheet.nrows):


        row_list = []

        # cell_value : 返回单元格中的数据
        # sale_amount_column_index = 3
        sale_amount = worksheet.cell_value(row_index,sale_amount_column_index)
        sale = worksheet.cell_value(row_index,sale_amount_column_index)
        sale_amount = float(str(sale).strip('$'),replace(',',''))
        if sale_amount > 1400.0:
            for column_index in range(worksheet.ncols):
                # cell_value : 返回单元格中的数据
                cell_value = worksheet.cell_value(row_index,column_index)

                # cell_type：返回单元格中的数据类型
                # 3：日期数据
                # if worksheet.cell_type(row_index, col_index) == 3:
                cell_type = worksheet.cell_type(row_index,column_index)
                if cell_type == 3:

                    # xldate_as_tuple 将Excel中代表日期，时间的数值转成元组
                    # datemode：使函数确定日期是基于1900年还是1904年，并据此将数值转换成正确的元组
                    date_cell = xldate_as_tuple(cell_value,workbook.datemode)
                    # strftime ：将data对象转换为一个具有特定格式的字符串
                    date_cell = date(*date_cell[0:3].strftime('%m/%d/%Y'))
                    # append ：将值追加给输出列表
                    row_list.append(date_cell)
                else:
                    row_list.append(cell_value)
        if row_list:
            data.append(row_list)
    for list_index,output_list in enumerate(date):
        for element_index,element in enumerate(output_list):
            output_worksheet.write(list_index,element_index,element)
output_workbook.save(output_file)




