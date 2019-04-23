import xlrd
import json
import re

from datetime import date, datetime

from KeyVo import KeyVo

file = 'G-goto跳转表.xlsx'

# 类型常量
TYPE_INT = 'Integer'
TYPE_STRING = 'String'


def read_excel():
    wb = xlrd.open_workbook(filename=file)  # 打开文件
    print(wb.sheet_names())  # 获取所有表格名字
    sheet1 = wb.sheet_by_index(0)  # 通过索引获取表格
    sheet2 = wb.sheet_by_name('Sheet1')  # 通过名字获取表格
    print(sheet1, sheet2)
    print(sheet1.name, sheet1.nrows, sheet1.ncols)
    rows = sheet1.row_values(2)  # 获取一整行内容
    cols = sheet1.col_values(3)  # 获取一整列内容
    print(rows)
    print(cols)
    print(sheet1.cell(0, 0))
    print(sheet1.cell(0, 0).value)  # 获取指定行列的内容
    print(sheet1.cell_value(0, 0))
    print(sheet1.col(0)[0].value)
    print(sheet1.cell(4, 1))
    print(sheet1.cell(4, 1).value)
    print(sheet1.cell(4, 2).value)
    print(sheet1.row_values(4))
    # 表格数据 ctype： 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error


def create_config_vo(p_filename, p_suffix='ts'):
    # 加载模板配置
    with open('template\config.json', 'r') as f:
        temp_map = json.loads(f.read())
        print('====加载模板文件配置成功')

    # 加载导出模板文件
    if p_suffix not in temp_map:
        print('不存在配置：'+p_suffix)
        return

    temp_cfg = temp_map[p_suffix]
    temp_url = temp_cfg['template']
    with open('template\\' + temp_url, 'r', encoding='utf-8') as f:
        tmp_str = f.read()
        print(tmp_str)

    # 正则表达式官方文档参考
    # https://docs.python.org/zh-cn/3.7/library/re.html
    # re.M 让$生效
    # re.DOTALL 让.可以匹配换行符
    str_no = re.sub('^<<<<\s*$(.+)^>>>>\s*$', '', tmp_str, flags=re.M|re.DOTALL)
    print(str_no)

    wb = xlrd.open_workbook(filename=p_filename)
    sheet = wb.sheet_by_index(0)
    export_name = sheet.cell(0, 0).value

    # 导出的文件名
    export_filename = export_name + 'Config' + '.' + p_suffix
    source_filename = p_filename

    t_row_count = sheet.nrows
    t_col_count = sheet.ncols
    t_vo_list = []
    t_comment_index_r = 0
    t_client_key_index_r = 1
    t_type_index_r = 2
    t_server_key_index_r = 3
    t_comment_rows = sheet.row(t_comment_index_r)
    t_client_key_rows = sheet.row(t_client_key_index_r)
    t_type_rows = sheet.row(t_type_index_r)
    t_server_key_rows = sheet.row(t_server_key_index_r)

    for i in range(1, t_col_count):
        cell_client = t_client_key_rows[i]
        cell_server = t_server_key_rows[i]
        cell_type = t_type_rows[i]
        '''
        表格数据 ctype： 
        0 empty
        1 string
        2 number
        3 date
        4 boolean
        5 error
        '''
        if cell_client.ctype != 1 and cell_server.ctype != 1:  # 跳过非字符串的格子
            continue
        t_type = TYPE_INT if cell_type.value == TYPE_INT else TYPE_STRING
        t_vo = KeyVo(p_index=i, p_type=t_type)
        t_vo_list.append(t_vo)
        if cell_client.ctype == 1:
            t_vo.key_client = cell_client.value
            t_vo.export_client = True
        if cell_server.ctype == 1:
            t_vo.key_server = cell_server.value
            t_vo.export_server = True
        print(i, t_vo.index, t_vo.type, t_vo.key_client, t_vo.key_server)


# read_excel()
create_config_vo(file)
