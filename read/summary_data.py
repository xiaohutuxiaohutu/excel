#!/usr/bin/env python3
# coding=UTF-8
import openpyxl
import os


def summary_data(path):
    pattern_column = '状态'
    patter_value = '激活'
    work_book = openpyxl.load_workbook(path, read_only=False)  #
    sheetnames = work_book.sheetnames
    print('读取到如下表单，请选择需要汇总的表单，输入表单名或index序列号，用英文,号分割:')
    print(sheetnames)
    input_sheet_name = input()
    summary_sheet_name = input('请选择要汇总到的表单，输入表单名或index序列号:\n')
    sheet_names = input_sheet_name.split(',')
    result_list = []
    # 过滤需要的表单信息
    for name in sheet_names:
        if name.isdigit():
            # print('sheet_name:' + sheetnames[int(name)])
            sheet = work_book.worksheets[int(name) - 1]
        else:
            # print('sheet_name:' + name)
            sheet = work_book[name]
        # 获取当前表单的所有行数
        rows = sheet.rows
        for index, row in enumerate(rows):
            # print(row)
            line = [col.value for col in row]  # 取值
            if index == 0:
                # 第一行表头
                # print(line)
                pattern_column_index = line.index(pattern_column)
                # print(pattern_column_index)
                continue
            else:
                # print(line)
                if patter_value == line[pattern_column_index]:
                    result_list.insert(len(result_list), line)
                # break
    # print(result_list)
    # 读取要合并的表单
    if summary_sheet_name.isdigit():
        print('sheet_name:' + sheetnames[int(summary_sheet_name) - 1])
        sheet1 = work_book.worksheets[int(summary_sheet_name) - 1]
    else:
        print('sheet_name:' + summary_sheet_name)
        sheet1 = work_book[summary_sheet_name]
    max_row = sheet1.max_row
    print('最大行数%i;' % max_row)
    # 先删除原有的表单行数
    for i in range(0, max_row - 1):
        sheet1.delete_rows(max_row - i)
    # 写入
    for row_index in range(0, len(result_list)):
        for col_index in range(0, len(result_list[row_index])):
            sheet1.cell(row=row_index + 2, column=col_index + 1).value = result_list[row_index][col_index]
    work_book.save(path)


if __name__ == '__main__':
    input_file_path = input('请输入需要合并的excel表绝对路径，多个文件以英文,号隔开:\n')
    while input_file_path.strip() == '':
        input_file_path = input('写入的文件路径不能为空，请重新输入：\n')
    file_path = input_file_path.split(',')
    for index, item in enumerate(file_path):
        print('读取第%i个文件:%s' % (index + 1, item))
        while os.path.splitext(item)[1] not in ['.xlsx']:
            item = input('写入的文件类型不支持，目前只支持xlsx，请重新输入：\n')
        try:
            summary_data(item)
        except:
            print('第%i个文件处理失败:%s' % (index + 1, item))
