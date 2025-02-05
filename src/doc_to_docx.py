import os
import win32com.client as win32
import docx
from openpyxl import Workbook


def convert_doc_to_docx(doc_file_path,save_path):
    """
    将 .doc 文件转换为 .docx 文件
    :param doc_file_path: .doc 文件的路径
    :return: .docx 文件的路径
    """
    # 获取 .doc 文件的目录和文件名
    file_dir, file_name = os.path.split(doc_file_path)
    # 生成 .docx 文件的文件名
    docx_file_name = os.path.splitext(file_name)[0] + '.docx'
    # 生成 .docx 文件的完整路径
    docx_file_path = os.path.join(save_path, docx_file_name)
    print(f'docx_file_path: {docx_file_path}')

    word = win32.gencache.EnsureDispatch('Word.Application')
    file_path_abs = os.path.abspath(doc_file_path)
    doc = word.Documents.Open(file_path_abs)

    # 将文档保存为 .docx 格式
    doc.SaveAs2(docx_file_path, FileFormat=16)
    doc.Close()
    word.Quit()
    return docx_file_path

if __name__ == '__main__':
    doc_file_path = r'./data/doc/'
    save_path = r'./data/newdocx/'
    all_files = os.listdir(doc_file_path)
    doc_files = [file for file in all_files if file.endswith(('.doc'))]
    os.makedirs(os.path.dirname(save_path), exist_ok=True)
    for doc_file in doc_files:
        doc_file_path = os.path.join(doc_file_path, doc_file)
        docx_file_path = convert_doc_to_docx(doc_file_path, save_path)
        print(f'{doc_file} 转换为 {docx_file_path}')

