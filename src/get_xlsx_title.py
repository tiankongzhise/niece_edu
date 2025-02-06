import os
import pandas




def get_xlsx_title(file_path):
    df = pandas.read_excel(file_path, engine='openpyxl')
    return df.columns.tolist()

def get_file_list(file_path):
    all_file_name = os.listdir(file_path)
    xlsx_file_name = [file for file in all_file_name if file.endswith('.xlsx')]
    return xlsx_file_name

def main():
    file_path = r'./data/tq/'
    xlsx_file_name = get_file_list(file_path)
    result = []
    for file in xlsx_file_name:
        temp_file_path = os.path.join(file_path,file)
        result.append({'file_name':file,'title':get_xlsx_title(temp_file_path)})
    save_path = r'./data/title.xlsx'
    print(result)
    pandas.DataFrame(result).to_excel(save_path,index=False)

if __name__ == '__main__':
    main()