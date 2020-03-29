#!/usr/bin/env python3
# coding=UTF-8
import os
import sys

curDir = os.getcwd()  # 获取当前文件路径
rootDir = curDir[:curDir.find("excel\\") + len("excel\\")]  # 获取myProject，也就是项目的根路径
sys.path.append(rootDir)
import read

if __name__ == '__main__':
    read_file_path = input("请输入读取文件路径:\n")
    while read_file_path.strip() == '':
        read_file_path = input('读取文件路径不能为空，请重新输入：\n')
    while os.path.splitext(read_file_path)[1] not in ['.xls', '.xlsx']:
        read_file_path = input('文件类型不支持，目前只支持 xls xlsx，请重新输入：\n')
    read.read_excel_xls()
    '''
    
    write_file_path = input("请输入需要写入的文件路径：\n")
    while write_file_path.strip() == '':
        write_file_path = input('写入的文件路径不能为空，请重新输入：\n')
    # while os.path.splitext(write_file_path)[1] not in ['.xlsx']:
    #     write_file_path = input('写入的文件类型不支持，目前只支持xlsx，请重新输入：\n')

    input_org_match = input("请输入源文件需要匹配的列名，直接enter默认使用 工号,岗位,身份证号,状态,店代码1,店代码2,店代码3,店代码4,店代码5：\n")
    if input_org_match.strip() == '':
        input_org_match = '工号,岗位,身份证号,状态,店代码1,店代码2,店代码3,店代码4,店代码5'

    input_dest_match = input("请输入目标文件需要匹配的列名，直接enter默认使用 工号,岗位,身份证号,状态,经销店代码1,经销店代码2,经销店代码3,经销店代码4,经销店代码5：\n")
    if input_dest_match.strip() == '':
        input_dest_match = '工号,岗位,身份证号,状态,经销店代码1,经销店代码2,经销店代码3,经销店代码4,经销店代码5'
    while len(input_org_match.split(',')) != len(input_dest_match.split(',')):
        input_org_match = input('请重新输入源文件列名')
        if input_org_match.strip() == '':
            input_org_match = '工号,岗位,身份证号,状态,店代码1,店代码2,店代码3,店代码4,店代码5'
        input_dest_match = input("请重新输入目标文件需要匹配的列名，直接enter默认使用 工号,岗位,身份证号,状态,经销店代码1,经销店代码2,经销店代码3,经销店代码4,经销店代码5：\n")
        if input_dest_match.strip() == '':
            input_dest_match = '工号,岗位,身份证号,状态,经销店代码1,经销店代码2,经销店代码3,经销店代码4,经销店代码5'
    origin_pattern = input_org_match
    dest_pattern = input_dest_match
    # 校验输入的参数
    # 文件路径是否合法
    # 写入文件
    
    
    file_path = write_file_path.split(',')
    for index, item in enumerate(file_path):
        print('读取第%i个文件:%s' % (index + 1, item))
        read.write_xlsx(read_file_path, item)
    
    '''
