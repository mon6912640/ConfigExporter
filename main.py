import time

import xlrd
import json
import re
import os.path
import json_minify

from monkey_xls import KeyVo, ExcelVo, TempCfgVo

file = 'G-goto跳转表.xlsx'

# 模板配置文件 template.json
template_config = None
cfg_vo_map = {}
file_count = 0

# 类型常量
TYPE_INT = 'Integer'
TYPE_STRING = 'String'

# 操作枚举
OP_VO = 0b1
OP_DATA = 0b10


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
            with open('template\\template.json', 'r', encoding='utf-8') as f:
                # json_minify库支持json文件里面添加注释
                template_config = json.loads(json_minify.json_minify(f.read()))
                print('====加载模板文件配置成功')

        if p_key not in template_config:
            print('template.json 中不存在 ' + p_key + ' 配置：')
            exit()

        cfg_vo_map[p_key] = TempCfgVo(template_config[p_key])

    return cfg_vo_map[p_key]


def create_config_vo(excel_vo: ExcelVo):
    sheet = excel_vo.sheet

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

            def rpl_property(m1):
                key_str = m1.group(1)
                return replace_key(key_str, p_excel_vo=excel_vo, p_key_vo=v)

            result += re.sub('<#(.*?)#>', rpl_property, loop_str)
        # 这里需要把前后的换行干掉
        return result.rstrip('\n')

    output_str = re.sub('^<<<<\\s*$(.+?)^>>>>\\s*$', rpl_loop, excel_vo.cfg.str_tmp, flags=re.M | re.DOTALL)

    def rpl_export(m):
        key_str = m.group(1)
        return replace_key(key_str, p_excel_vo=excel_vo)

    output_str = re.sub('<#(.*?)#>', rpl_export, output_str)

    with open(os.path.join(excel_vo.cfg.output_path, excel_vo.export_filename), 'w', encoding='utf-8') as f:
        f.write(output_str)
        # print('成功导出', excel_vo.export_filename)


# 替换关键字
def replace_key(p_key: str, p_excel_vo: ExcelVo, p_key_vo: KeyVo = None):
    if p_key == 'source_filename':
        return p_excel_vo.source_filename
    elif p_key == 'export_name':
        return p_excel_vo.export_name
    elif p_key == 'export_class_name':
        return p_excel_vo.export_class_name
    elif p_key_vo is not None:
        if p_key == 'property_name':
            return p_key_vo.key_client
        elif p_key == 'type':
            return transform_tye(p_key_vo.type, p_excel_vo.cfg.type_map)
        elif p_key == 'comment':
            # 这里需要注意，python的数字类型不会自动转换为字符串，这里需要强转一下
            return p_key_vo.comment
        elif p_key == 'index':
            return str(p_key_vo.index)
    else:
        return 'undefinded'


# 根据配置转换类型
def transform_tye(p_type, p_map):
    if p_type in p_map:
        return p_map[p_type]
    else:
        return None


# 导出vo文件
def export_vo_file(p_key, op):
    cfg = get_cfg_by_key(p_key)
    # 遍历文件夹内所有的xlsx文件
    global file_count
    for fpath, dirnames, fnames in os.walk(cfg.source_path):
        # print('fpath', fpath)
        # print('dirname', dirnames)
        # print('fnames', fnames)
        # print('--------------')
        for fname in fnames:
            file_url = os.path.join(fpath, fname)
            name, ext = os.path.splitext(file_url)
            if ext == '.xlsx':
                wb = xlrd.open_workbook(filename=file_url)
                sheet = wb.sheet_by_index(0)
                if sheet.cell_type(0, 0) != 1:
                    print('第一行第一格没有填写表名，无效的xlsx：' + fname)
                    continue
                excel_vo = ExcelVo(cfg=cfg, sheet=sheet, source_path=file_url, filename=fname)
                file_count += 1
                if op & OP_VO == OP_VO:  # 导出vo类
                    create_config_vo(excel_vo)
                if op & OP_DATA == OP_DATA:  # 导出json数据
                    print('fuck')


# 导出配置数据文件
def export_vo_data():
    print('fuck')
    return


start = time.time()
# export_vo_file('as')
end = time.time()
print('输出 %s 个文件' % (file_count))
print('总用时', end - start)
