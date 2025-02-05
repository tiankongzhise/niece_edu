import requests
from bs4 import BeautifulSoup
import re


def get_page(url):
    headers = {
        'User-Agent': ''
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.text
    return None

def save_html(html, file_name):
    with open(file_name, 'w', encoding='utf-8_sig') as f:
        f.write(html)

def parse_html(html):
    soup = BeautifulSoup(html, 'html.parser')
    # divs = soup.find('div', class_=re.compile(r'rich_media_content\s+js_underline_content'))

    result_list = []

    # print(div)
    # 查找所有section标签，且powered-by属性为xiumi.us
    sections = soup.find_all('section', attrs={'powered-by': 'xiumi.us'})
    temp = [section.text.strip() for section in sections]
    school_name = [name for name in temp if '总计划数' not in name and '往期精彩' not in name]
    print(len(school_name))


    # 查找所有class为rich_pages wxw-img的img标签
    img_tags = soup.find_all('img', class_=lambda x: x and x.startswith('rich_pages wxw-img'))

    # 提取每个img标签的data-src属性
    school_list = [img.get('data-src') for img in img_tags if img.get('data-src')]
    print(len(school_list))



    # 确保sections和p_tags的数量一致
    if len(school_name) != len(school_list):
        print("解析错误：学校名称和URL数量不匹配")
        for i in range(len(school_name)):
            # 检查i是否在school_list的有效范围内
            if i < len(school_list):
                url = school_list[i]
            else:
                url = []  # 如果下标越界，则用空列表填充
            
            # 将学校名称和URL添加到字典中，并添加到结果列表中
            result_list.append({
                'name': school_name[i],
                'url': url
            })
    else:
        for x,y in zip(school_name,school_list):
            result_list.append({
                'name': x,
                'url': y
            })

    return result_list
def save_to_csv(data:list[dict], file_name:str)->None:
    with open(file_name, 'w', encoding='utf-8_sig') as f:
        for item in data:
            f.write(f"{item['name']},{item['url']}\n")

if __name__ == '__main__':
    url = 'https://mp.weixin.qq.com/s/Wrkc39LV6LRy4zsl-Aj42g'
    html = get_page(url)
    contents = parse_html(html)
    file_path = './data/school_list.csv'
    save_to_csv(contents, file_path)
    
