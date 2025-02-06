import pandas as pd
import re
import os
from pathlib import Path
from datetime import datetime
# ================= 配置区域（根据实际情况修改）=================
CONFIG = {
    "input_dir": "./data/tq/",      # 原始Excel存放目录
    "output_dir": "./data/processed/",     # 处理结果输出目录
    "standard_columns": [              # 标准列名列表
        "专业组", 
        "专业名称", 
        "计划数", 
        "学费", 
        "备注"
    ],
    "column_mapping": {  # 智能映射规则（支持正则表达式）
        "专业组": [
            r".*组.*",
            "类别", "分组", "组别", "专业群组"
        ],
        "专业名称": [
            r"专业.*名称",
            "招生专业", "拟招生专业", "专 业 名 称"
        ],
        "计划数": [
            r"计划数.*",
            "专业计划", "单招计划", r"\d+人计划"
        ],
        "学费": [
            r"学费[（(].*元.*",
            "学费标准", "专业学费", "学费\n标准",
            r".*费用.*"  # 扩展匹配规则
        ],
        "备注": [
            r"备注|说明",
            "其他信息", "补充说明"
        ]
    },
    "special_handling": {  # 特殊列处理规则
        "代码类": [".*代码.*", "专业代码", "代码"],
        "学院类": [".*学院.*", "二级学院", "院部"]
    }
}
# ================= 核心代码（无需修改）=================
def setup_dirs():
    """创建输出目录结构"""
    Path(CONFIG['output_dir']).mkdir(exist_ok=True)
    Path(f"{CONFIG['output_dir']}/unmatched_columns").mkdir(exist_ok=True)
    Path(f"{CONFIG['output_dir']}/special_columns").mkdir(exist_ok=True)
def clean_column_name(raw_col):
    """清洗列名：去除换行符和特殊字符"""
    return re.sub(r'[\n\\/（）()\s、：，]', '', str(raw_col)).strip().lower()
def match_column_pattern(col_name, patterns):
    """智能匹配列名模式"""
    cleaned = clean_column_name(col_name)
    # 优先匹配正则表达式
    for pattern in patterns:
        if isinstance(pattern, str) and re.search(pattern, col_name, re.IGNORECASE):
            return True
    # 其次匹配关键字
    for pattern in patterns:
        if isinstance(pattern, str) and pattern.lower() in cleaned:
            return True
    return False
def map_columns(df, filename):
    """执行列名映射，返回处理后的DataFrame和日志信息"""
    column_log = []
    matched = set()
    new_columns = {}
    
    # 主映射逻辑
    for std_col, patterns in CONFIG['column_mapping'].items():
        for raw_col in df.columns:
            if raw_col in new_columns.values():
                continue
            if match_column_pattern(raw_col, patterns):
                new_columns[raw_col] = std_col
                matched.add(raw_col)
                column_log.append(f"{raw_col} → {std_col}")
                break
    
    # 特殊列处理
    special_cols = {}
    for category, patterns in CONFIG['special_handling'].items():
        special_cols[category] = [col for col in df.columns 
                                  if match_column_pattern(col, patterns)]
    
    # 未匹配列处理
    unmatched = [col for col in df.columns if col not in matched]
    if unmatched:
        column_log.append(f"未匹配列：{unmatched}")
    
    # 执行重命名
    df = df.rename(columns={v:k for k,v in new_columns.items()})
    
    # 添加来源标记
    df['_源文件'] = filename
    
    return df[list(new_columns.values())], column_log, special_cols, unmatched
def process_all_files():
    """批量处理所有Excel文件"""
    setup_dirs()
    
    all_data = []
    log_records = []
    special_data = {}
    unmatched_data = {}
    for file_path in Path(CONFIG['input_dir']).glob("*.xlsx"):
        try:
            # 读取文件
            df = pd.read_excel(file_path, engine='openpyxl')
            
            # 处理合并单元格
            df = df.ffill()
            
            # 执行列映射
            processed_df, log, special_cols, unmatched = map_columns(df, file_path.name)
            
            # 收集数据
            all_data.append(processed_df)
            log_records.append({
                "文件名": file_path.name,
                "处理时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "日志": "\n".join(log),
                "特殊列": str(special_cols),
                "未匹配列": str(unmatched)
            })
            
            # 保存特殊列数据
            for category, cols in special_cols.items():
                if cols:
                    special_df = df[cols].copy()
                    special_df['_源文件'] = file_path.name
                    output_path = f"{CONFIG['output_dir']}/special_columns/{category}_{file_path.stem}.xlsx"
                    special_df.to_excel(output_path, index=False)
            
            # 保存未匹配列数据
            if unmatched:
                unmatched_df = df[unmatched].copy()
                unmatched_df['_源文件'] = file_path.name
                output_path = f"{CONFIG['output_dir']}/unmatched_columns/unmatched_{file_path.stem}.xlsx"
                unmatched_df.to_excel(output_path, index=False)
                
        except Exception as e:
            print(f"处理文件 {file_path.name} 时出错：{str(e)}")
    
    # 合并数据
    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        
        # 动态生成最终列名（确保列存在）
        final_columns = []
        for col in CONFIG['standard_columns'] + ['_源文件']:
            if col in final_df.columns:
                final_columns.append(col)
            else:
                print(f"警告：列 '{col}' 不存在，已自动跳过")
        
        # 仅保留存在的列
        final_df = final_df[final_columns]
        
        # 添加缺失列（如果需要）
        for col in CONFIG['standard_columns']:
            if col not in final_df.columns:
                final_df[col] = None  # 或填充默认值
        
        final_output = f"{CONFIG['output_dir']}/统一数据_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        final_df.to_excel(final_output, index=False)
        print(f"合并文件已生成：{final_output}")
    
    # 保存日志
    if log_records:
        log_df = pd.DataFrame(log_records)
        log_output = f"{CONFIG['output_dir']}/处理日志_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        log_df.to_excel(log_output, index=False)
        log_df.to_excel(log_output, index=False)
        print(f"处理日志已生成：{log_output}")
    # 保存特殊列汇总
    for category in CONFIG['special_handling'].keys():
        category_files = list(Path(f"{CONFIG['output_dir']}/special_columns").glob(f"{category}_*.xlsx"))
        if category_files:
            category_dfs = [pd.read_excel(f) for f in category_files]
            combined_special = pd.concat(category_dfs, ignore_index=True)
            special_output = f"{CONFIG['output_dir']}/汇总_{category}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            combined_special.to_excel(special_output, index=False)
def validate_data(df):
    """数据校验"""
    # 检查必要列是否存在
    required_columns = ['专业组', '专业名称']
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"缺少必要列：{missing}")
    
    # 检查学费是否为数值
    if '学费' in df.columns:
        df['学费'] = pd.to_numeric(df['学费'].str.replace(r'\D', '', regex=True), errors='coerce')
    
    return df
if __name__ == "__main__":
    print("=== 开始处理 ===")
    process_all_files()
    print("=== 处理完成 ===")
    print(f"请检查输出目录：{CONFIG['output_dir']}")