import base64
import os
import json
from tencent_sdk import img_to_excel_tencent_sdk




def process_base64_to_excel(input_file="./data/base64.txt", output_dir="./data/excel/"):
    """
    从base64.txt读取数据，调用腾讯云SDK处理，并保存Excel文件
    """
    # 创建输出目录
    os.makedirs(output_dir, exist_ok=True)
    
    # 统计处理结果
    result = {
        "total": 0,
        "success": 0,
        "failed": []
    }

    try:
        with open(input_file, "r", encoding="utf-8") as f:
            for line in f:
                result["total"] += 1
                line = line.strip()
                
                # 解析行数据
                try:
                    filename, img_base64 = line.split("|", 1)
                    if not filename or not img_base64:
                        raise ValueError("数据不完整")
                except Exception as e:
                    result["failed"].append(f"行{result['total']}格式错误: {str(e)}")
                    continue
                
                # 调用腾讯云SDK
                try:
                    response = img_to_excel_tencent_sdk(img_base64)
                    response_json = json.loads(response)
                    
                    # 提取目标数据
                    if "Data" not in response_json:
                        raise KeyError("响应中缺少Data字段")
                        
                    data = response_json.get("Data")
                    if not data:
                        raise ValueError("Data字段为空")
                    
                    # 解码并保存Excel
                    try:
                        excel_data = base64.b64decode(data)
                        output_path = os.path.join(output_dir, f'{filename}.xlsx')
                        
                        with open(output_path, "wb") as excel_file:
                            excel_file.write(excel_data)
                            
                        result["success"] += 1
                        print(f"成功处理: {filename}")
                        
                    except Exception as decode_err:
                        raise ValueError(f"Excel解码失败: {str(decode_err)}") from decode_err
                        
                except Exception as process_err:
                    error_msg = f"文件[{filename}]处理失败: {str(process_err)}"
                    result["failed"].append(error_msg)
                    print(error_msg)
                    
    except FileNotFoundError:
        print(f"错误：输入文件不存在 {input_file}")
        return result
    except Exception as e:
        print(f"系统错误: {str(e)}")
        return result

    # 打印统计结果
    print(f"\n处理完成！成功 {result['success']}/{result['total']}")
    if result["failed"]:
        print("\n失败列表：")
        for msg in result["failed"]:
            print(f" - {msg}")
    
    return result


if __name__ == '__main__':
    process_base64_to_excel()
