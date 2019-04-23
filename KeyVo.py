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