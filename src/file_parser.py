import os
import subprocess
import win32com.client as win32


def save_json_to_file(json_data, output_file):
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(json_data)
        

def doc2docx_by_soffice(input_file, output_path=None):
    # 使用 LibreOffice 进行转换
    filename = os.path.basename(input_file)
    # 如果 output_path 为 None，则使用与 input_file 相同的目录
    if output_path is None:
        output_path = os.path.dirname(input_file)
    try:
        subprocess.run(['soffice', '--headless', '--convert-to', 'docx', input_file, '--outdir', output_path], check=True)
        print(f"Converted {filename} to {output_path}")
        return os.path.join(output_path, os.path.splitext(filename)[0] + '.docx')
    except subprocess.CalledProcessError as e:
        print(f"Failed to convert {filename}: {e}")

def doc2docx_by_pywin32(input_file, output_path=None):
    if output_path is None:
        output_path = input_file.replace('.doc', '.docx')
    print(f"Converting {input_file} to {output_path}")
    word = win32.Dispatch('Word.Application')
    word.Visible = False  # 不显示Word应用程序
    wb = word.Documents.Open(input_file)
    wb.SaveAs(output_path, FileFormat=16)  # 16 表示 .docx 文件格式
    wb.Close()
    word.Quit()
    return output_path
