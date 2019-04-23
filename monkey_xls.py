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


class ExportVo:
    # xls源文件名
    source_filename = ''
    # xls文件的源路径
    source_path = ''
    # 导出的文件名（不包含后缀名）
    export_name = ''
    # 导出的文件名（包含后缀名）
    export_filename = ''
    # 导出的类名
    export_class_name = ''
