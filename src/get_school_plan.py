import pandas as pd
import re
from tqdm import tqdm
def get_school_plan(file_path,save_path):
    # 读取 Excel 文件
    df = pd.read_excel(file_path)

    # 定义正则表达式
    pattern_total = re.compile(r"单招总计划.*?为\s*(\d+)\s*人")
    pattern_veteran = re.compile(r"退役军人(\d+)人")
    pattern_social = re.compile(r"其他社会人员(\d+)人")
    pattern_sports = re.compile(r"体育特长生(\d+)人")
    pattern_arts = re.compile(r"艺术特长生(\d+)人")

    # 初始化结果列
    df['总招生计划'] = '-'
    df['退役军人'] = '-'
    df['其他社会人员'] = '-'
    df['体育特长生'] = '-'
    df['艺术特长生'] = '-'
    df['提及退役军人'] = 0
    df['提及其他社会人员'] = 0
    df['提及体育特长生'] = 0
    df['提及艺术特长生'] = 0

    # 遍历每一行
    for index, row in tqdm(df.iterrows(),desc="处理文件中", unit="行",total=len(df)):
        field_value = row['字段']
        if field_value == '-':
            continue  # 跳过无数据的行

        # 提取总招生计划
        match_total = pattern_total.search(field_value)
        if match_total:
            df.at[index, '总招生计划'] = match_total.group(1)

        # 提取退役军人
        match_veteran = pattern_veteran.search(field_value)
        if match_veteran:
            df.at[index, '退役军人'] = match_veteran.group(1)
        if '军人' in field_value:
            df.at[index, '提及退役军人'] = 1

        # 提取其他社会人员
        match_social = pattern_social.search(field_value)
        if match_social:
            df.at[index, '其他社会人员'] = match_social.group(1)
        if '社会' in field_value:
            df.at[index, '提及其他社会人员'] = 1

        # 提取体育特长生
        match_sports = pattern_sports.search(field_value)
        if match_sports:
            df.at[index, '体育特长生'] = match_sports.group(1)
        if '体育' in field_value:
            df.at[index, '提及体育特长生'] = 1

        # 提取艺术特长生
        match_arts = pattern_arts.search(field_value)
        if match_arts:
            df.at[index, '艺术特长生'] = match_arts.group(1)
        if '艺术' in field_value:
            df.at[index, '提及艺术特长生'] = 1

    # 保存结果到新的 Excel 文件
    df.to_excel(save_path, index=False)
    print(f"处理完成，结果已保存到 {save_path}")
if __name__ == '__main__':
    file_path = r'./data/2025招生计划字段提取整理版.xlsx'
    save_path = r'./data/2025招生计划提取.xlsx'
    get_school_plan(file_path,save_path)