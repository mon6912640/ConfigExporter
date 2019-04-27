import time

import xlrd
import json
import re
import os.path
import json_minify

from monkey_xls import KeyVo, ExportVo, TempCfgVo

file = 'G-goto跳转表.xlsx'

# 模板配置文件 config.json
template_config = None
cfg_vo_map = {}

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


# 通过key获取模板配置数据
def get_cfg_by_key(p_key) -> TempCfgVo:
    global cfg_vo_map

    if p_key not in cfg_vo_map:
        global template_config

        if not template_config:
            # 加载模板配置
            with open('template\\config.json', 'r', encoding='utf-8') as f:
                # json_minify库支持json文件里面添加注释
                template_config = json.loads(json_minify.json_minify(f.read()))
                print('====加载模板文件配置成功')

        if p_key not in template_config:
            print('config.json 中不存在 ' + p_key + ' 配置：')
            exit();

        cfg_vo_map[p_key] = TempCfgVo(template_config[p_key])

    return cfg_vo_map[p_key]


def create_config_vo(p_file_path, p_cfg):
    export_vo = ExportVo()
    export_vo.source_path = p_file_path
    export_vo.source_filename = os.path.basename(p_file_path)

    export_vo.cfg = p_cfg
    type_map = export_vo.cfg.type_map
    temp_url = export_vo.cfg.template
    '''
    在Python3，可以通过open函数的newline参数来控制Universal new line mode
    读取时候，不指定newline，则默认开启Universal new line mode，所有\n, \r, or \r\n被默认转换为\n；
    写入时，不指定newline，则换行符为各系统默认的换行符（\n, \r, or \r\n, ），指定为newline='\n'，则都替换为\n（相当于Universal new line mode）；
    不论读或者写时，newline=''都表示不转换。
    参考链接：https://www.zhihu.com/question/19751023
    '''
    with open('template\\' + temp_url, 'r', encoding='utf-8') as f:
        str_tmp = f.read()

    wb = xlrd.open_workbook(filename=p_file_path)
    sheet = wb.sheet_by_index(0)
    export_vo.export_name = sheet.cell(0, 0).value

    # 导出的文件名
    export_vo.export_filename = export_vo.export_name + 'Config' + '.' + export_vo.cfg.suffix
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
        # print(i, t_vo.index, t_vo.type, t_vo.key_client, t_vo.key_server)

    # 跳过无key列表的数据列表
    if len(vo_list) == 0:
        return

    # 根据配置转换类型
    def transform_tye(p_type):
        if p_type in type_map:
            return type_map[p_type]
        else:
            return None

    # 正则表达式官方文档参考
    # https://docs.python.org/zh-cn/3.7/library/re.html
    # re.M 让$生效
    # re.DOTALL 让.可以匹配换行符
    def rpl_loop(m):
        result = ''
        loop_str = str(m.group(1)).lstrip('\n')
        for v in vo_list:
            if not v.key_client:
                continue

            def rpl_property(m):
                key_str = m.group(1)
                if key_str == 'property_name':
                    return v.key_client
                elif key_str == 'type':
                    return transform_tye(v.type)
                elif key_str == 'comment':
                    return v.comment
                elif key_str == 'index':
                    # 这里需要注意，python的数字类型不会自动转换为字符串，这里需要强转一下
                    return str(v.index)

            result += re.sub('<#(.*?)#>', rpl_property, loop_str)
        # 这里需要把前后的换行干掉
        return result.rstrip('\n')

    output_str = re.sub('^<<<<\s*$(.+?)^>>>>\s*$', rpl_loop, str_tmp, flags=re.M | re.DOTALL)

    def rpl_export(m):
        key_str = m.group(1)
        if key_str == 'source_filename':
            return export_vo.source_filename
        elif key_str == 'export_name':
            return export_vo.export_name
        elif key_str == 'export_class_name':
            return export_vo.export_class_name

    output_str = re.sub('<#(.*?)#>', rpl_export, output_str)

    with open(os.path.join(export_vo.cfg.output_path, export_vo.export_filename), 'w', encoding='utf-8') as f:
        f.write(output_str)
        # print('成功导出', export_vo.export_filename)


def run(p_key):
    cfg = get_cfg_by_key(p_key)
    # 遍历文件夹内所有的xlsx文件
    for fpath, dirnames, fnames in os.walk(cfg.source_path):
        # print('fpath', fpath)
        # print('dirname', dirnames)
        # print('fnames', fnames)
        # print('--------------')
        for fname in fnames:
            file_url = os.path.join(fpath, fname)
            name, ext = os.path.splitext(file_url)
            if ext == '.xlsx':
                create_config_vo(file_url, cfg)


# read_excel()
# create_config_vo(file, 'ts3')
start = time.time()
run('lua')
end = time.time()
print('总用时', end - start)
