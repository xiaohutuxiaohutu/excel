# import xlrd
# import xlutils.copy
from xlrd import open_workbook
from xlwt import Workbook
from xlutils.copy import copy
import openpyxl
import os

file_path = 'C:\\Users\\23948\\Desktop\\excel\\人员列表汇理.xls'
dest_file_path = 'C:\\Users\\23948\\Desktop\\excel\\L1认证考试数据.xlsx'
# 工号     岗位	身份证号	状态	店代码1	    店代码2	    店代码3	    店代码4	    店代码5
# 工号     岗位	身份证号 状态     经销店代码1	经销店代码2	经销店代码3	经销店代码4	经销店代码5

dest_file_path_xlsx = 'C:\\Users\\23948\\Desktop\\excel\\L1认证考试数据.xlsx'
# 默认需要匹配的列标签
default_pattern = {
    '工号': '工号'
    , '岗位': '岗位'
    , '身份证号': '身份证号'
    , '状态': '状态'
    , '店代码1': '经销店代码1'
    , '店代码2': '经销店代码2'
    , '店代码3': '经销店代码3'
    , '店代码4': '经销店代码4'
    , '店代码5': '经销店代码5'
}
origin_pattern = '工号,岗位,身份证号,状态,店代码1,店代码2,店代码3,店代码4,店代码5'
dest_pattern = '工号,岗位,身份证号,状态,经销店代码1,经销店代码2,经销店代码3,经销店代码4,经销店代码5'
origin_pattern_index = []
dest_pattern_index = []
# excel_list = []
read_excel_map = {}

read_result_map = {}


# 读取xls
def read_excel_xls(path):
    wb = open_workbook(file_path)
    sheets = wb.sheets()
    for sheet in sheets:
        # print(u"表单 %s 共 %d 行 %d 列" % (sheet.name, sheet.nrows, sheet.ncols))
        # print(sheet.name)
        for row in range(0, sheet.nrows):
            # print(row)
            values = sheet.row_values(row)
            if row == 0:
                read_result_map['header'] = values
                continue
            else:

                read_excel_map[values[0]] = values
                # print(values[0])
                # excel_list.insert(row, values)
    # print(read_excel_map)
    read_result_map['rows'] = read_excel_map
    # print(read_result_map)
    return read_result_map


# read_excel_xls(file_path)


# 读取xlsx
def read_excel_xlsx(path):
    excel_list = []
    # 输入文件路径
    '''
    read_file_path = input("请输入读取文件路径:\n")
    print(read_file_path)
    '''
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
    print(read_excel_map)
    read_result_map['rows'] = read_excel_map
    return read_result_map
    # excel_list.insert(index, line)
    # print(excel_list)
    # table = data.get_sheet_by_name(names[0])  # 获得指定名称的页
    # nrows = table.rows  # 获得行数 类型为迭代器
    # ncols = table.columns  # 获得列数 类型为迭代器
    # print(type(nrows))
    # for row in nrows:
    #     print(row)  # 包含了页名，cell，值
    #     line = [col.value for col in row]  # 取值
    #     print(line)
    # # 读取单元格
    # print(table.cell(1, 1).value)

    # read_excel_xlsx(dest_file_path_xlsx)


# 获取下标 pattern_list 需要匹配的列表，dest_list 被匹配的列表

def get_index_list(pattern_list, dest_list):
    list = []
    for index_, value in enumerate(pattern_list):
        # print('%i %s' % (index_, value))
        list.insert(index_, dest_list.index(value))
    return list


# 写入 xlsx文件
def write_xlsx(path):
    # 读取xlsx目标文件
    work_book = openpyxl.load_workbook(path, read_only=False)  #
    sheetnames = work_book.sheetnames
    print('读取到如下表单，请选择需要修改的表单，输入表单名或index序列号，用英文,号分割:')
    print(sheetnames)
    input_sheet_name = input()
    # print(input_sheet_names)

    # 先读取源文件
    if os.path.splitext(file_path)[1] == '.xls':
        read_data = read_excel_xls(file_path)
    else:
        read_data = read_excel_xlsx(file_path)
    # 获取源文件的表头
    origin_header = read_data['header']
    origin_rows = read_data['rows']
    # print(origin_header)
    origin_index_list = get_index_list(origin_pattern.split(','), origin_header)
    print(origin_index_list)
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
        for index, row in enumerate(rows):
            # print(row)
            line = [col.value for col in row]  # 取值

            if index == 0:
                # 第一行表头
                dest_index_list = get_index_list(dest_pattern.split(','), line)
                continue
            else:
                print(line)
                # 匹配行数修改内容
                # 获取dest_index_list[0]
                dest_index = dest_index_list[0]
                key_value = line[dest_index]
                # print(key_value)
                # 获取元数据匹配的行
                origin_row_value = origin_rows[key_value]
                print(origin_row_value)
                # print(origin_row_value)
                for index1, value in enumerate(origin_index_list):
                    value_ = origin_row_value[value]
                    col_num = dest_index_list[index1]
                    sheet.cell(row=index + 1, column=col_num + 1).value = value_
                    print(sheet.cell(row=index + 1, column=col_num + 1).value)
                # break
    # work_book.close()
    work_book.save(path)


if __name__ == '__main__':
    write_xlsx(dest_file_path_xlsx)

'''

def read_file(file_url):
    try:
        data = open_workbook(file_url)
        sheet = data.get_sheet('人员列表')
        print(sheet)
        return data
    except Exception as e:
        print(str(e))

'''


# data = read_file(file_path)
# rb = open_workbook(file_path)
# index = rb.sheet_by_index(0)
# wb = copy(rb)
# active_sheet = Workbook.get_active_sheet
# print(active_sheet)
# sheets = wb.sheets()
# for sheet in sheets:
#     value = sheet.cell(0, 0).value
#     print(value)
# print(len(sheets))
# index = wb.sheet_by_index(0)
# book = index.book
# rows = book.ragged_rows
# print(rows)
# print(index)
# print(book)


# s = wb.sheet_by_index(0)

# sheets = wb.get_sheets()
# print(sheets)
# print(s.cell(0, 0).value)


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
