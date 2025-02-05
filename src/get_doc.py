import pandas as pd
import requests
import os

def download_doc(doc, save_path,exc='.docx'):
    pre_url = "https://www.hneeb.cn/hnxxg/741/742/gzdzzc25/"

    url = f'{pre_url}{doc}{exc}'
    # 发送HTTP请求下载文件
    response = requests.get(url)
    if response.status_code == 200:
        with open(save_path, 'wb') as file:
            file.write(response.content)
        print(f"文档已成功下载并保存到 {save_path}")
    else:
        print(f"下载失败，状态码: {response.status_code}")

def get_download_list(file_path):
    # 获取目录下所有文件
    all_files = os.listdir(file_path)
    # 过滤出 .docx 和 .doc 文件
    doc_files = [os.path.splitext(file)[0] for file in all_files if file.endswith(('.docx', '.doc'))]
    return doc_files
    




if __name__ == '__main__':
    df = pd.read_excel('./data/2025学校名单.xlsx')
    doc_list = df['学校名称'].to_list()
        # 确保保存目录存在
    os.makedirs(os.path.dirname('./data/doc/'), exist_ok=True)
    # for doc in doc_list:
    #     save_path = f"./data/doc/{doc}.docx"
    #     download_doc(doc, save_path)
    download_list = get_download_list('./data/doc/')
    undownload_list = [doc for doc in doc_list if doc not in download_list]
    # for doc in undownload_list:
    #     save_path = f"./data/doc/{doc}.pdf"
    #     download_doc(doc, save_path,'.pdf')
    print(undownload_list)