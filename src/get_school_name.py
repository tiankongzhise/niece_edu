import os
import pandas as pd

def get_school_name(file_path):
    all_files = os.listdir(file_path)
    school_name = [ os.path.splitext(file)[0] for file in all_files]
    return school_name

if __name__ == '__main__':
    file_path = r'./data/doc/'
    save_path = r'./data/单招学校名称.xlsx'
    school_name = get_school_name(file_path)
    df = pd.DataFrame({'学校名称':school_name})
    df.to_excel(save_path,index = False)