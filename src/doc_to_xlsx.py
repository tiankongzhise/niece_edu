import os
from pathlib import Path
import pandas as pd
from docx import Document
import comtypes.client
import pdfplumber
def convert_doc_to_docx(doc_path, docx_path):
    try:
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False
        local_doc_path = os.path.abspath(doc_path)
        local_docx_path = os.path.abspath(docx_path)
        doc = word.Documents.Open(str(local_doc_path))
        doc.SaveAs(str(local_docx_path), FileFormat=16)  # 16代表docx格式
        doc.Close()
        word.Quit()
        return True
    except Exception as e:
        print(f"转换DOC到DOCX失败：{e}")
        return False
def extract_tables_from_docx(docx_path):
    try:
        doc = Document(docx_path)
        tables = []
        for table in doc.tables:
            data = []
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                data.append(row_data)
            if data:
                df = pd.DataFrame(data[1:], columns=data[0]) if len(data) > 1 else pd.DataFrame(data)
                tables.append(df)
        return tables
    except Exception as e:
        print(f"从{docx_path}提取表格失败：{e}")
        return []
def extract_tables_from_doc(doc_path):
    try:
        doc_path = Path(doc_path)
        docx_temp = doc_path.with_suffix('.docx')
        if convert_doc_to_docx(doc_path, docx_temp):
            tables = extract_tables_from_docx(docx_temp)
            os.remove(docx_temp)
            return tables
        else:
            return []
    except Exception as e:
        print(f"处理{doc_path}时出错：{e}")
        return []
def extract_tables_from_pdf(pdf_path):
    try:
        tables = []
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                page_tables = page.extract_tables()
                for table_num, table in enumerate(page_tables):
                    df = pd.DataFrame(table[1:], columns=table[0]) if len(table) > 1 else pd.DataFrame(table)
                    tables.append(df)
        return tables
    except Exception as e:
        print(f"从{pdf_path}提取表格失败：{e}")
        return []
def process_file(input_path, output_dir):
    input_path = Path(input_path)
    output_path = output_dir / f"{input_path.stem}.xlsx"
    tables = []
    
    if input_path.suffix.lower() == '.docx':
        tables = extract_tables_from_docx(input_path)
    elif input_path.suffix.lower() == '.doc':
        tables = extract_tables_from_doc(input_path)
    elif input_path.suffix.lower() == '.pdf':
        tables = extract_tables_from_pdf(input_path)
    else:
        print(f"不支持的文件类型：{input_path.suffix}")
        return
    
    if not tables:
        print(f"{input_path}中没有提取到表格。")
        return
    
    # 保存到Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for i, df in enumerate(tables, start=1):
            df.to_excel(writer, sheet_name=f"Table {i}", index=False)
    print(f"已保存表格到：{output_path}")
def main():
    input_dir = Path("./data/doc")
    output_dir = Path("./data/tq")
    output_dir.mkdir(parents=True, exist_ok=True)
    
    supported_extensions = ('.doc', '.docx', '.pdf')
    
    for file_path in input_dir.glob('*'):
        if file_path.suffix.lower() in supported_extensions:
            print(f"处理文件：{file_path}")
            process_file(file_path, output_dir)
        else:
            print(f"跳过不支持的文件：{file_path}")
if __name__ == "__main__":
    main()