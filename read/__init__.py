# import xlrd
# import xlutils.copy
from xlrd import open_workbook
# from xlwt import Workbook
# from xlutils.copy import copy
import openpyxl
import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox
import os

root = tk.Tk()
root.withdraw()
# 默认需要匹配的列标签
origin_pattern = '工号,岗位,身份证号,状态,店代码1,店代码2,店代码3,店代码4,店代码5'
dest_pattern = '工号,岗位,身份证号,状态,经销店代码1,经销店代码2,经销店代码3,经销店代码4,经销店代码5'
# 源文件需要读取的表头标签对应的列数集合
# origin_pattern_index = []
# 目标文件表头对应的列数集合
# dest_pattern_index = []


application_window = tk.Tk()

xls_file_types = [('excel文件', '.xls')]
xlsx_file_types = [('excel文件', '.xlsx')]
xls_xlsx_file_types = [('excel文件', '.xls'), ('excel文件', '.xlsx')]


# 打开文件选择窗口
def open_file_win(title, file_type):
    answer = filedialog.askopenfilenames(parent=application_window,
                                         initialdir=os.getcwd(),
                                         title=title,
                                         filetypes=file_type)
    tk.Tk().wm_withdraw()
    if answer:
        return answer
    else:
        tkinter.messagebox.showinfo('提示', '没有选择文件，请重新选择')
        open_file_win(title)


# 读取xls
def read_excel_xls(path):
    read_result_map = {}
    read_excel_map = {}
    wb = open_workbook(path)
    sheets = wb.sheets()
    for sheet in sheets:
        # print(u"表单 %s 共 %d 行 %d 列" % (sheet.name, sheet.nrows, sheet.ncols))
        for row in range(0, sheet.nrows):
            values = sheet.row_values(row)
            if row == 0:
                read_result_map['header'] = values
                continue
            else:
                read_excel_map[values[0]] = values
    read_result_map['rows'] = read_excel_map
    return read_result_map


# 读取xlsx 返回类型{'header':'','rows':'{key:[],key1:[].....}'}
def read_excel_xlsx(path):
    read_excel_map = {}
    read_result_map = {}
    work_book = openpyxl.load_workbook(path)  # 读取xlsx文件
    # names = data.get_sheet_names
    sheetnames = work_book.sheetnames
    print('读取到如下表单，请选择需要读取那几个表单，输入表单名或index序列号，用英文,号分割:')
    print(sheetnames)
    input_sheet_names = input()
    print(input_sheet_names)
    names_split = input_sheet_names.split(',')

    for name in names_split:
        if name.isdigit():
            index_ = int(name) - 1
            sheet = work_book.worksheets[int(name) - 1]

        else:
            sheet = work_book[name]
        rows = sheet.rows
        # print(type(rows))
        # columns = sheet.columns
        for index, row in enumerate(rows):
            # print(row)
            line = [col.value for col in row]  # 取值
            if index == 0:
                read_result_map['header'] = line
                continue
            else:
                # print(line)
                read_excel_map[line[0]] = line
    # print(read_excel_map)
    read_result_map['rows'] = read_excel_map
    return read_result_map


# 获取下标 pattern_list 需要匹配的列表，dest_list 被匹配的列表
def get_index_list(pattern_list, dest_list):
    list = []
    for index_, value in enumerate(pattern_list):
        # print('%i %s' % (index_, value))
        list.insert(index_, dest_list.index(value))
    return list


# 写入 xlsx文件
def write_xlsx(read_path, write_path):
    # 读取xlsx目标文件
    work_book = openpyxl.load_workbook(write_path, read_only=False)  #
    sheetnames = work_book.sheetnames
    print('读取到如下表单，请选择需要修改的表单，输入表单名或index序列号，用英文,号分割:')
    print(sheetnames)
    input_sheet_name = input()
    # 先读取源文件
    if os.path.splitext(read_path)[1] == '.xls':
        read_data = read_excel_xls(read_path)
    else:
        read_data = read_excel_xlsx(read_path)
    # 获取源文件的表头
    origin_header = read_data['header']
    origin_rows = read_data['rows']
    # print(origin_header)
    origin_index_list = get_index_list(origin_pattern.split(','), origin_header)
    # print(origin_index_list)
    # 需要操作的sheet集合
    sheet_names = input_sheet_name.split(',')
    # 循环打开sheet表单操作
    for name in sheet_names:
        if name.isdigit():
            sheet = work_book.worksheets[int(name) - 1]
        else:
            sheet = work_book[name]
        # 获取当前表单的所有行数
        rows = sheet.rows
        # columns = sheet.columns
        # 需要修改的当前sheet的表头，用来对比
        cur_sheet_header = []
        for index2, row in enumerate(rows):
            # print(row)
            line = [col.value for col in row]  # 取值

            if index2 == 0:
                # 第一行表头
                dest_index_list = get_index_list(dest_pattern.split(','), line)
                continue
            else:
                # print(line)
                # 匹配行数修改内容
                # 获取dest_index_list[0]
                dest_index = dest_index_list[0]
                key_value = line[dest_index]
                # print(key_value)
                # 获取元数据匹配的行
                origin_row_value = origin_rows[key_value]
                # print(origin_row_value)
                # print(origin_row_value)
                for index1, value in enumerate(origin_index_list):
                    value_ = origin_row_value[value]
                    col_num = dest_index_list[index1]
                    sheet.cell(row=index2 + 1, column=col_num + 1).value = value_
                    # print(sheet.cell(row=index2 + 1, column=col_num + 1).value)
                # break
    # work_book.close()
    work_book.save(write_path)


# 写入xls文件
def write_xls(read_path, write_path):
    print()


if __name__ == '__main__':
    print()
    # write_xlsx(dest_file_path_xlsx)


def filter_excel(workbook, column_name=0, by_name='Sheet0'):
    """

    :param workbook:
    :param column_name:
    :param by_name: 对应的Sheet页
    :return:
    """
    table = workbook.sheet_by_name(by_name)  # 获得表格
    total_rows = table.nrows  # 拿到总共行数
    columns = table.row_values(column_name)  # 某一行数据 ['姓名', '用户名', '联系方式', '密码']
    excel_list = []
    for one_row in range(1, total_rows):  # 也就是从Excel第二行开始，第一行表头不算

        row = table.row_values(one_row)
        # if row:
        #     row_object = {}
        #     for i in range(0, len(columns)):
        #         key = table_header[columns[i]]
        #         row_object[key] = row[i]  # 表头与数据对应
        #
        #     excel_list.append(row_object)

    return excel_list
