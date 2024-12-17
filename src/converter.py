import os

from src.docx2json import docx_to_json
from src.file_parser import doc2docx_by_pywin32, save_json_to_file


class Converter:
    def __init__(self, data_path, output_file='output/output.json'):
        self.data_path = data_path
        self.output_file = output_file

    def convert(self, is_save=False):
        # 判断文件类型，调用相应的转换函数
        filename = os.path.basename(self.data_path)
        filedir = os.path.dirname(self.data_path)
        if self.data_path.endswith('.doc') and not filename.startswith('~$'):
            self.data_path = doc2docx_by_pywin32(self.data_path)
            json_data = docx_to_json(self.data_path, self.output_file)
        elif self.data_path.endswith('.docx') and not filename.startswith('~$'):
            json_data = docx_to_json(self.data_path, self.output_file)
        else:
            return {'status': 'error', 'message': 'Unsupported file type'}
        
        if is_save:
            # 创建输出目录
            output_path = os.path.dirname(self.output_file)
            if output_path not in os.listdir():
                os.mkdir(output_path)
            # 将JSON字符串写入文件
            save_json_to_file(json_data, self.output_file)
        
        return json_data

    def md2json(self):
        pass
    
    def excel2json(self):
        pass
    
    def csv2json(self):
        pass