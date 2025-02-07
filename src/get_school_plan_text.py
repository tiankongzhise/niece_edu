import os
import re
from docx import Document
import PyPDF2
import pandas as pd
from win32com import client as wc  # 用于处理 .doc 文件
from tqdm import tqdm  # 用于显示进度条
# 定义正则表达式
# pattern = re.compile(r"2025年单招总计划为.*?省考试院最终公布的为准.*?公布为准")
pattern = re.compile(r"2025年单招总计划数为.*?为准")

# 提取文字段的函数
def extract_text_from_file(file_path):
    if file_path.endswith('.docx'):
        # 处理 .docx 文件
        doc = Document(file_path)
        text = '\n'.join([para.text for para in doc.paragraphs])
    elif file_path.endswith('.doc'):
        # 处理 .doc 文件
        word = wc.Dispatch("Word.Application")
        doc = word.Documents.Open(os.path.abspath(file_path))
        text = doc.Content.Text
        doc.Close()
        word.Quit()
    elif file_path.endswith('.pdf'):
        # 处理 .pdf 文件
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = '\n'.join([page.extract_text() for page in reader.pages])
    else:
        return '-'

    # 查找匹配的文字段
    match = pattern.search(text)
    if match:
        return match.group(0)  # 返回匹配的文字段
    else:
        return '-'  # 没有匹配到则返回 '-'

# 获取文件列表
file_list = [f for f in os.listdir('./data/doc') if f.endswith(('.docx', '.doc', '.pdf'))]

# 遍历目录并提取数据
data = []
for file_name in tqdm(file_list, desc="处理文件中", unit="文件"):
# for file_name in os.listdir('./data/doc'):
    file_path = os.path.join('./data/doc', file_name)
    if file_name.endswith(('.docx', '.doc', '.pdf')):
        school_name = os.path.splitext(file_name)[0]  # 去除扩展名作为学校名称
        field_value = extract_text_from_file(file_path)  # 提取文字段
        exc = os.path.splitext(file_name)[1]  # 去扩展名
        data.append([school_name, field_value,exc])

# 将结果保存到 Excel 文件
df = pd.DataFrame(data, columns=['学校名称', '字段','文件类型'])
df.to_excel('./data/2025招生计划字段提取.xlsx', index=False)

print("提取完成，结果已保存到 ./data/2025招生计划字段提取.xlsx")