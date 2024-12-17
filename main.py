
import os
from src.converter import Converter

file_path = r'data\需求格式大纲.doc'
# 相对路径 转为 绝对路径
absolute_path = os.path.abspath(file_path)

cvt = Converter(absolute_path)
cvt.convert(is_save=True)