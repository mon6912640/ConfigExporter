import xlrd
import json
import re
import os.path
import json_minify

from monkey_xls import KeyVo, ExportVo

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


def create_config_vo(p_file_path, p_suffix='ts'):
    export_vo = ExportVo()
    export_vo.source_path = p_file_path
    export_vo.source_filename = os.path.basename(p_file_path)

    # 加载模板配置
    with open('template\config.json', 'r', encoding='utf-8') as f:
        # json_minify库支持json文件里面添加注释
        temp_map = json.loads(json_minify.json_minify(f.read()))
        print('====加载模板文件配置成功')

    # 加载导出模板文件
    if p_suffix not in temp_map:
        print('不存在配置：' + p_suffix)
        return

    temp_cfg = temp_map[p_suffix]
    temp_url = temp_cfg['template']
    with open('template\\' + temp_url, 'r', encoding='utf-8') as f:
        str_tmp = f.read()

    wb = xlrd.open_workbook(filename=p_file_path)
    sheet = wb.sheet_by_index(0)
    export_vo.export_name = sheet.cell(0, 0).value

    # 导出的文件名
    export_vo.export_filename = export_vo.export_name + 'Config' + '.' + p_suffix
    export_vo.export_class_name = export_vo.export_name + 'Config'

    row_count = sheet.nrows
    col_count = sheet.ncols
    vo_list = []
    comment_index_r = 0
    client_key_index_r = 1
    type_index_r = 2
    server_key_index_r = 3
    comment_rows = sheet.row(comment_index_r)
    client_key_rows = sheet.row(client_key_index_r)
    type_rows = sheet.row(type_index_r)
    server_key_rows = sheet.row(server_key_index_r)

    for i in range(1, col_count):
        cell_client = client_key_rows[i]
        cell_server = server_key_rows[i]
        cell_type = type_rows[i]
        cell_comment = comment_rows[i]
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
        vo_list.append(t_vo)
        if cell_client.ctype == 1:
            t_vo.key_client = cell_client.value
            t_vo.export_client = True
        if cell_server.ctype == 1:
            t_vo.key_server = cell_server.value
            t_vo.export_server = True
        t_vo.comment = cell_comment.value
        print(i, t_vo.index, t_vo.type, t_vo.key_client, t_vo.key_server)

    # 根据配置转换类型
    def transform_tye(p_type):
        if p_type in temp_cfg['typeMap']:
            return temp_cfg['typeMap'][p_type]
        else:
            return None

    # 正则表达式官方文档参考
    # https://docs.python.org/zh-cn/3.7/library/re.html
    # re.M 让$生效
    # re.DOTALL 让.可以匹配换行符

    def rpl_loop(m):
        result = ''
        loop_str = m.group(1)
        for v in vo_list:
            def rpl_property(m):
                key_str = m.group(1)
                if key_str == 'property_name':
                    return v.key_client
                elif key_str == 'type':
                    return transform_tye(v.type)
                elif key_str == 'comment':
                    return v.comment

            result += re.sub('<#(.*?)#>', rpl_property, loop_str)
        return result

    output_str = re.sub('^<<<<\s*$(.+)^>>>>\s*$', rpl_loop, str_tmp, flags=re.M | re.DOTALL)

    def rpl_export(m):
        key_str = m.group(1)
        if key_str == 'source_filename':
            return export_vo.source_filename
        elif key_str == 'export_name':
            return export_vo.export_name
        elif key_str == 'export_class_name':
            return export_vo.export_class_name

    output_str = re.sub('<#(.*?)#>', rpl_export, output_str)
    print(output_str)


# read_excel()
create_config_vo(file, 'as')
