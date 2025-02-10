import os
import base64

# 目标目录
img_dir = "./pdfimg/new/"
OUTPUT_DIR = "./data/"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "base642.txt")

# 支持的图片扩展名及其对应的MIME类型
IMAGE_MIME_MAP = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".webp": "image/webp",
    ".svg": "image/svg+xml",
}

def img_to_base64(img_dir):
    # 检查目录是否存在
    if not os.path.isdir(img_dir):
        print(f"错误：目录 '{img_dir}' 不存在")
        return {}
    
    base64_data = {}
    
    # 遍历目录中的文件
    for filename in os.listdir(img_dir):
        filepath = os.path.join(img_dir, filename)
        
        # 跳过子目录和非文件项
        if not os.path.isfile(filepath):
            print(f"跳过目录或无效文件: {filename}")
            continue
        
        # 提取扩展名并转为小写
        ext = os.path.splitext(filename)[1].lower()
        
        # 检查是否为支持的图片格式
        if ext not in IMAGE_MIME_MAP:
            print(f"跳过不支持的文件类型: {filename}")
            continue
        
        try:
            # 读取二进制数据
            with open(filepath, "rb") as img_file:
                binary_data = img_file.read()
            
            # 转换为Base64字符串
            base64_str = base64.b64encode(binary_data).decode("utf-8")
            
            # 添加MIME类型前缀（可选）
            mime_type = IMAGE_MIME_MAP[ext]
            full_base64 = f"data:{mime_type};base64,{base64_str}"
            
            # 保存结果
            base64_data[filename] = full_base64
            print(f"转换成功: {filename}")
            
        except Exception as e:
            print(f"处理文件 {filename} 失败: {str(e)}")
    
    return base64_data


def save_base64_to_file(base64_data, output_path):
    """将Base64数据保存到文本文件"""
    try:
        # 确保输出目录存在
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # 写入文件（每行一个Base64字符串）
        with open(output_path, "w", encoding="utf-8") as f:
            # for b64_str in base64_data.values():
            #     f.write(b64_str + "\n")  # 添加换行符分隔
            # 修改保存逻辑为同时记录文件名
            for name, b64_str in base64_data.items():
                file_name_without_ext = os.path.splitext(name)[0]  # 去掉扩展名
                f.write(f"{file_name_without_ext}|{b64_str}\n")  # 使用竖线分隔
                
        print(f"\n保存成功！共写入 {len(base64_data)} 条记录到：{output_path}")
    except PermissionError:
        print(f"错误：无权限写入文件 {output_path}")
    except Exception as e:
        print(f"保存文件失败：{str(e)}")

if __name__ == "__main__":
    result = img_to_base64(img_dir)
    save_base64_to_file(result,OUTPUT_FILE)