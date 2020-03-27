from xlrd import open_workbook
from xlwt import Workbook
from xlutils.copy import copy

file_path = 'C:\\Users\\23948\\Desktop\\excel\\人员列表汇理.xls'
dest_file_path = 'C:\\Users\\23948\\Desktop\\excel\\L1认证考试数据（20200326）.xlsx'
rb = open_workbook(file_path)
names = rb.sheet_names()
nsheets = rb.nsheets
# print(nsheets)
# print(names)
sheets = rb.sheets()


# for i in len(sheets):
#     sh = rb.get_sheet(0)
#     print( u"表单 %s 共 %d 行 %d 列" % (sh.name, sh.nrows, sh.ncols))
# 工号	姓名	岗位	身份证号	状态	店代码1	店代码2	店代码3	店代码4	店代码5

# value = sheet.cell(0, 0).value
# print(value)


def get_org_excel_list(path):
    list = []
    workbook = open_workbook(file_path)
    sheets = workbook.sheets()
    for sheet in sheets:
        # print(u"表单 %s 共 %d 行 %d 列" % (sheet.name, sheet.nrows, sheet.ncols))
        # print(sheet.name)
        for row in range(0, sheet.nrows):
            # print(row)
            values = sheet.row_values(row)
            print(values)
            list.insert(row, values)
    return list;


excel_list = get_org_excel_list(file_path)
print(excel_list)
