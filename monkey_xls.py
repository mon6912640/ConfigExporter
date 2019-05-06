from enum import Enum
from typing import List

import xlrd


class KeyTypeEnum(Enum):
    """
    配置数据类型枚举
    """
    # 数字
    TYPE_INT = 'Integer'
    # 字符串
    TYPE_STRING = 'String'


class ExcelIndexEnum(Enum):
    """
    excel关键index枚举
    """
    comment_r = 0
    client_key_r = 1
    type_r = 2
    server_key_r = 3
    data_start_r = 4

    data_start_c = 1


class KeyVo:
    key_client = ''
    key_server = ''
    comment = ''
    type = 0
    index = 0
    export_client = False
    export_server = False

    def __init__(self, p_index, p_type):
        self.index = p_index
        self.type = p_type


class TempCfgVo:
    suffix = ''
    template = ''
    type_map = None
    # 源文件夹路径
    source_path = ''
    # 输出路径
    output_path = ''
    # 数据json输出路径
    json_path = ''
    # 是否打包在同一个json文件中
    json_pack_in_one = False
    # 压缩方式
    json_compress = ''
    # 压缩文件后缀名
    compress_suffix = ''
    # 额外副本路径
    json_copy_path = ''
    # 每次生成输出前是否清除旧文件
    clean = False

    __str_tmp: str = None

    def __init__(self, p_cfg_data):
        self.set_data(p_cfg_data)

    def set_data(self, p_cfg_data):
        if 'suffix' in p_cfg_data:
            self.suffix = p_cfg_data['suffix']
        if 'template' in p_cfg_data:
            self.template = p_cfg_data['template']
        if 'typeMap' in p_cfg_data:
            self.type_map = p_cfg_data['typeMap']
        if 'sourcePath' in p_cfg_data:
            self.source_path = p_cfg_data['sourcePath']
        if 'outputPath' in p_cfg_data:
            self.output_path = p_cfg_data['outputPath']
        if 'jsonPath' in p_cfg_data:
            self.json_path = p_cfg_data['jsonPath']
        if 'jsonPackInOne' in p_cfg_data:
            self.json_pack_in_one = p_cfg_data['jsonPackInOne']
        if 'jsonCompress' in p_cfg_data:
            self.json_compress = p_cfg_data['jsonCompress']
        if 'jsonCopyPath' in p_cfg_data:
            self.json_copy_path = p_cfg_data['jsonCopyPath']
        if 'clean' in p_cfg_data:
            self.clean = p_cfg_data['clean']
        if 'compressSuffix' in p_cfg_data:
            self.compress_suffix = p_cfg_data['compressSuffix']

    @property
    def str_tmp(self):
        """
        模板文本
        :return:
        """
        if self.__str_tmp is None:
            '''
            在Python3，可以通过open函数的newline参数来控制Universal new line mode
            读取时候，不指定newline，则默认开启Universal new line mode，所有\n, \r, or \r\n被默认转换为\n；
            写入时，不指定newline，则换行符为各系统默认的换行符（\n, \r, or \r\n, ），指定为newline='\n'，则都替换为\n（相当于Universal new line mode）；
            不论读或者写时，newline=''都表示不转换。
            参考链接：https://www.zhihu.com/question/19751023
            '''
            with open('template\\' + self.template, 'r', encoding='utf-8') as f:
                self.__str_tmp = f.read()
        return self.__str_tmp


class ExcelVo:
    # xls源文件名
    source_filename = ''
    # xls文件的源路径
    source_path = ''
    # 模板配置引用
    cfg: TempCfgVo = None
    # 表格数据引用
    sheet: xlrd.sheet.Sheet = None
    __key_vo_list: List[KeyVo] = None

    def __init__(self, cfg, sheet, source_path, filename):
        self.cfg = cfg
        self.sheet = sheet
        self.source_path = source_path
        self.source_filename = filename

    @property
    def export_name(self):
        """
        导出的文件名（不包含后缀名）
        :return:
        """
        if self.sheet is not None:
            return self.sheet.cell(0, 0).value

    @property
    def export_filename(self):
        """
        导出的文件名（包含后缀名）
        :return:
        """
        if self.sheet is not None:
            return self.export_name + 'Config' + '.' + self.cfg.suffix

    @property
    def export_class_name(self):
        """
        导出的类名
        :return:
        """
        if self.sheet is not None:
            return self.export_name + 'Config'

    @property
    def key_vo_list(self):
        """
        获取Excel的KeyVo列表
        :return:
        """
        if self.__key_vo_list is None:
            col_count = self.sheet.ncols
            self.__key_vo_list = []
            for i in range(1, col_count):
                comment_index_r = ExcelIndexEnum.comment_r.value
                client_key_index_r = ExcelIndexEnum.client_key_r.value
                type_index_r = ExcelIndexEnum.type_r.value
                server_key_index_r = ExcelIndexEnum.server_key_r.value

                comment_rows = self.sheet.row(comment_index_r)
                client_key_rows = self.sheet.row(client_key_index_r)
                type_rows = self.sheet.row(type_index_r)
                server_key_rows = self.sheet.row(server_key_index_r)
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
                t_type = KeyTypeEnum.TYPE_INT.value if cell_type.value == KeyTypeEnum.TYPE_INT.value else KeyTypeEnum.TYPE_STRING.value
                t_vo = KeyVo(p_index=i, p_type=t_type)
                self.__key_vo_list.append(t_vo)
                if cell_client.ctype == 1:
                    t_vo.key_client = cell_client.value
                    t_vo.export_client = True
                if cell_server.ctype == 1:
                    t_vo.key_server = cell_server.value
                    t_vo.export_server = True
                t_vo.comment = cell_comment.value
        return self.__key_vo_list
