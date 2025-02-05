import json
import types
from tencentcloud.common import credential
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.ocr.v20181119 import ocr_client, models
from dotenv import load_dotenv
import os
import base64
load_dotenv()
def img_to_excel_tencent_sdk(img_base64:str)->str:
    try:
        # 实例化一个认证对象，入参需要传入腾讯云账户 SecretId 和 SecretKey，此处还需注意密钥对的保密
        # 代码泄露可能会导致 SecretId 和 SecretKey 泄露，并威胁账号下所有资源的安全性。以下代码示例仅供参考，建议采用更安全的方式来使用密钥，请参见：https://cloud.tencent.com/document/product/1278/85305
        # 密钥可前往官网控制台 https://console.cloud.tencent.com/cam/capi 进行获取
        secret_id = os.getenv("TENCENTCLOUD_SECRET_ID")
        secret_key = os.getenv("TENCENTCLOUD_SECRET_KEY")
        print(f'secret_id:{secret_id},secret_key:{secret_key}')
        cred = credential.Credential(secret_id,
                                    secret_key)
        # 实例化一个http选项，可选的，没有特殊需求可以跳过
        httpProfile = HttpProfile()
        httpProfile.endpoint = "ocr.tencentcloudapi.com"

        # 实例化一个client选项，可选的，没有特殊需求可以跳过
        clientProfile = ClientProfile()
        clientProfile.httpProfile = httpProfile
        # 实例化要请求产品的client对象,clientProfile是可选的
        client = ocr_client.OcrClient(cred, "", clientProfile)

        # 实例化一个请求对象,每个接口都会对应一个request对象
        req = models.RecognizeTableAccurateOCRRequest()
        params = {
            "ImageBase64": img_base64
        }
        req.from_json_string(json.dumps(params))

        # 返回的resp是一个RecognizeTableAccurateOCRResponse的实例，与请求对象对应
        resp = client.RecognizeTableAccurateOCR(req)
        # 输出json格式的字符串回包
        return resp.to_json_string()

    except TencentCloudSDKException as err:
        print(err)

if __name__ == '__main__':
    input_file="./data/base64.txt"
    output_dir="./data/excel/"
    # 统计处理结果
    result = {
        "total": 0,
        "success": 0,
        "failed": []
    }
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
                print(f'filename:{filename},response_json.keys:{response_json.keys()}')
                
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
            break