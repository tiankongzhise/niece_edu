import os
import re
from docx import Document
import PyPDF2
import pandas as pd
from win32com import client as wc  # 用于处理 .doc 文件
from tqdm import tqdm  # 用于显示进度条
from concurrent.futures import ThreadPoolExecutor, as_completed  # 用于多线程

# 定义正则表达式
pattern = re.compile(r"2025年单招总计划数为.*?省教育考试院.*?公布为准")

# 提取文字段的函数
def extract_text_from_file(file_path):
    try:
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
            return None, '-'  # 不支持的文件格式

        # 查找匹配的文字段
        match = pattern.search(text)
        if match:
            return os.path.splitext(os.path.basename(file_path))[0], match.group(0),os.path.splitext(os.path.basename(file_path))[1]  # 返回学校名称和匹配的文字段
        else:
            return os.path.splitext(os.path.basename(file_path))[0], '-' ,os.path.splitext(os.path.basename(file_path))[1]  # 没有匹配到则返回学校名称和 '-'
    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {e}")
        return os.path.splitext(os.path.basename(file_path))[0], '-'  # 出错时返回学校名称和 '-'

# 获取文件列表
file_list = [os.path.join('./data/doc', f) for f in os.listdir('./data/doc') if f.endswith(('.docx', '.doc', '.pdf'))]

# 多线程处理文件
data = []
submit = []
with ThreadPoolExecutor(max_workers=4) as executor:  # 设置最大线程数
    futures = {executor.submit(extract_text_from_file, file_path): file_path for file_path in file_list}
    for future in tqdm(as_completed(futures), total=len(file_list), desc="提交文件中", unit="文件"):
        submit.append(future)
    
    for future in tqdm(submit, total=len(submit), desc="处理文件中", unit="文件"):
        school_name, field_value,exc = future.result()
        data.append([school_name, field_value,exc])

# 将结果保存到 Excel 文件
df = pd.DataFrame(data, columns=['学校名称', '字段','文件类型'])
df.to_excel('./data/2025招生计划字段提取.xlsx', index=False)

print("提取完成，结果已保存到 ./data/2025招生计划字段提取.xlsx")