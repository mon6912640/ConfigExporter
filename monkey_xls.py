from enum import Enum
from pathlib import Path
from typing import List

import openpyxl.cell.cell as opencell
import openpyxl.worksheet.worksheet
from openpyxl.cell.cell import Cell


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
    # 字段注释所在的排
    comment_r = 0
    # 前端key所在的排
    client_key_r = 1
    # 类型所在的排
    type_r = 2
    # 后端key所在的排
    server_key_r = 3
    # 数据起始排
    data_start_r = 4

    # 数据起始行
    data_start_c = 1


class KeyVo:
    key_client = ''
    key_server = ''
    comment = ''
    # 批注
    pz = ''
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
    # json打包文件名
    json_pack_name = ''
    # 压缩方式
    json_compress = ''
    # 压缩文件后缀名
    compress_suffix = ''
    # 额外副本路径
    json_copy_path = ''
    # 每次生成输出前是否清除旧文件
    clean = False
    # 生成的结构体是否在同一个文件内
    struct_in_one = False
    # 枚举类模板
    enum_tmp = ''
    # 导出枚举类名称
    enum_class_name = ''
    # 导出类文件命名格式 {0}表示表名
    export_vo_name = ''

    # 工具的路径
    app_dir: Path = None

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
        if 'structInOne' in p_cfg_data:
            self.struct_in_one = p_cfg_data['structInOne']
        if 'enumTmp' in p_cfg_data:
            self.enum_tmp = p_cfg_data['enumTmp']
        if 'enumClassName' in p_cfg_data:
            self.enum_class_name = p_cfg_data['enumClassName']
        if 'exportVoName' in p_cfg_data:
            self.export_vo_name = p_cfg_data['exportVoName']
        if 'jsonPackName' in p_cfg_data:
            self.json_pack_name = p_cfg_data['jsonPackName']

    @property
    def str_tmp(self):
        """
        模板文本
        :return:
        """
        if self.__str_tmp is None:
            self.__str_tmp = (self.app_dir / 'template' / self.template).read_text(encoding='utf-8')
        return self.__str_tmp


class ExcelVo:
    # xls源文件名
    source_filename = ''
    # xls文件的源路径
    source_path = ''
    # 模板配置引用
    cfg: TempCfgVo = None
    # 表格数据引用
    # sheet: xlrd.sheet.Sheet = None
    sheet: openpyxl.worksheet.worksheet.Worksheet = None

    __key_vo_list: List[KeyVo] = None
    __has_id_in_client = None
    __has_id_in_server = None
    __export_name = None

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
        if self.__export_name is None:
            # openpyxl的row和col起始是1
            if self.sheet is not None:
                self.__export_name = self.sheet.cell(1, 1).value
            else:
                self.__export_name = 'null'
        return self.__export_name

    @property
    def export_filename(self):
        """
        导出的文件名（包含后缀名）
        :return:
        """
        if self.sheet is not None:
            return self.cfg.export_vo_name.format(self.export_name) + '.' + self.cfg.suffix

    @property
    def export_class_name(self):
        """
        导出的类名
        :return:
        """
        if self.sheet is not None:
            return self.cfg.export_vo_name.format(self.export_name)

    @property
    def key_vo_list(self):
        """
        获取Excel的KeyVo列表
        :return:
        """
        if self.__key_vo_list is None:
            cells_row = list(self.sheet.rows)
            # col_count = self.sheet.max_column
            columns = self.sheet.columns

            comment_index_r = ExcelIndexEnum.comment_r.value
            client_key_index_r = ExcelIndexEnum.client_key_r.value
            type_index_r = ExcelIndexEnum.type_r.value
            server_key_index_r = ExcelIndexEnum.server_key_r.value

            comment_rows = cells_row[comment_index_r]
            client_key_rows = cells_row[client_key_index_r]
            type_rows = cells_row[type_index_r]
            server_key_rows = cells_row[server_key_index_r]

            self.__key_vo_list = []
            # for i in range(1, col_count):
            for cols in columns:
                col_index = cols[0].column - 1
                i = col_index
                if i < 1:
                    continue
                # print('i = ', i)
                cell_client: Cell = client_key_rows[i]
                cell_server: Cell = server_key_rows[i]
                cell_type: Cell = type_rows[i]
                cell_comment: Cell = comment_rows[i]

                if cell_client.data_type != opencell.TYPE_STRING and cell_server.data_type != opencell.TYPE_STRING:
                    # 跳过非字符串的格子
                    continue
                t_vo = KeyVo(p_index=i, p_type=cell_type.value)
                self.__key_vo_list.append(t_vo)
                if cell_client.data_type == opencell.TYPE_STRING:
                    t_vo.key_client = cell_client.value
                    t_vo.export_client = True
                if cell_server.data_type == opencell.TYPE_STRING:
                    t_vo.key_server = cell_server.value
                    t_vo.export_server = True
                t_vo.comment = cell_comment.value
                if cell_comment.comment:  # 批注
                    t_vo.pz = cell_comment.comment.content
                    # 把批注的“*/”去掉
                    if '*/' in t_vo.pz:
                        t_vo.pz = t_vo.pz.replace('*/', '')
                # if t_vo.pz:
                #     print('--', t_vo.pz)
        return self.__key_vo_list

    def has_id_in_client(self) -> bool:
        """
        是否有id主键（前端）
        :return:
        """
        if self.__has_id_in_client is None:
            key_list = self.key_vo_list
            has_id = False
            for v in key_list:
                if v.key_client == 'id':
                    has_id = True
                    break
            self.__has_id_in_client = has_id
        return self.__has_id_in_client

    def has_id_in_server(self) -> bool:
        """
        是否有id主键（后端）
        :return:
        """
        if self.__has_id_in_server is None:
            key_list = self.key_vo_list
            has_id = False
            for v in key_list:
                if v.key_server == 'id':
                    has_id = True
                    break
            self.__has_id_in_server = has_id
        return self.__has_id_in_server
