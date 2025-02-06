import os 

def compare_file(file_path_a,file_path_b):
    file_list_a = os.listdir(file_path_a)
    file_list_b = os.listdir(file_path_b)
    file_name_list_a = [os.path.splitext(file)[0]  for file in file_list_a]
    file_name_list_b = [os.path.splitext(file)[0]  for file in file_list_b]
    return list(set(file_name_list_a).difference(set(file_name_list_b)))

if __name__ == '__main__':
    file_path_a = r'./data/doc/'
    file_path_b = r'./data/tq/'
    file_name_list = compare_file(file_path_a,file_path_b)
    print(file_name_list)
