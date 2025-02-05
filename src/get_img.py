import csv
import requests
import os

# 确保img目录存在
if not os.path.exists('./img'):
    os.makedirs('./img')

# 读取CSV文件
with open('./data/school_list.csv', mode='r', encoding='utf-8') as file:
    reader = csv.reader(file)
    next(reader)  # 跳过标题行
    for row in reader:
        school_name, url = row
        try:
            response = requests.get(url)
            response.raise_for_status()
            with open(f'./img/{school_name}.jpg', 'wb') as img_file:
                img_file.write(response.content)
        except requests.exceptions.RequestException as e:
            print(f"Failed to download image for {school_name}: {e}")
