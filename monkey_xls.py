import xlrd


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
    __str_tmp: str = None

    def __init__(self, p_cfg_data):
        self.suffix = p_cfg_data['suffix']
        self.template = p_cfg_data['template']
        self.type_map = p_cfg_data['typeMap']
        self.source_path = p_cfg_data['sourcePath']
        self.output_path = p_cfg_data['outputPath']

    @property
    def str_tmp(self):
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
