# src/main/sms_util.py
"""
阿里云短信通知模块（使用官方 Tea SDK）
- 自动加载 ../resource/.env 中的 AccessKey
- 自动读取 ../resource/phonelist.txt 中的接收号码列表
- 支持批量发送（带 1 秒间隔防限流）
"""

import os
import sys
import json
import time
from typing import List, Tuple

from alibabacloud_dysmsapi20170525.client import Client as DysmsapiClient
from alibabacloud_tea_openapi import models as open_api_models
from alibabacloud_dysmsapi20170525 import models as dysmsapi_models
from alibabacloud_tea_util import models as util_models
from dotenv import load_dotenv


# ========================
# 加载 .env 配置文件
# ========================
_script_dir = os.path.dirname(__file__)
_dotenv_path = os.path.join(_script_dir, '..', 'resource', '.env')
if os.path.exists(_dotenv_path):
    load_dotenv(dotenv_path=_dotenv_path)
else:
    print(f"⚠️ .env 文件未找到: {_dotenv_path}", file=sys.stderr)


def _load_phone_numbers() -> List[str]:
    """
    从 ../resource/phonelist.txt 加载有效手机号列表。
    忽略空行、注释行（# 开头），并清洗非数字字符。
    """
    phonelist_path = os.path.join(_script_dir, '..', 'resource', 'phonelist.txt')
    if not os.path.exists(phonelist_path):
        print(f"⚠️ phonelist.txt 不存在: {phonelist_path}", file=sys.stderr)
        return []

    phones = set()
    with open(phonelist_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            clean_num = ''.join(filter(str.isdigit, line))
            if len(clean_num) == 11 and clean_num.startswith(('13', '14', '15', '17', '18', '19')):
                phones.add(clean_num)
            else:
                print(f"⚠️ 跳过无效号码: {line}", file=sys.stderr)
    return sorted(phones)


def _create_client(access_key_id: str, access_key_secret: str) -> DysmsapiClient:
    """创建阿里云短信客户端"""
    config = open_api_models.Config(
        access_key_id=access_key_id,
        access_key_secret=access_key_secret,
        region_id='cn-hangzhou'  # 短信服务仅支持杭州
    )
    return DysmsapiClient(config)


def _send_single_sms(
    client: DysmsapiClient,
    phone_number: str,
    todaycount: int,
    yesterdaycount: int,
    increment: int,
    sign_name: str = "云均信息技术工作室",
    template_code: str = "SMS_498585210"
) -> bool:
    """发送单条短信"""
    try:
        request = dysmsapi_models.SendSmsRequest(
            phone_numbers=phone_number,
            sign_name=sign_name,
            template_code=template_code,
            template_param=json.dumps({
                "todaycount": str(todaycount),
                "yesterdaycount": str(yesterdaycount),
                "increment": str(increment)
            })
        )
        runtime = util_models.RuntimeOptions()
        response = client.send_sms_with_options(request, runtime)
        body = response.body

        if body and body.code == "OK":
            print(f"✅ 短信发送成功！RequestId: {body.request_id} → {phone_number}")
            return True
        else:
            print(f"❌ 发送失败 ({phone_number})：Code={body.code}, Message={body.message}", file=sys.stderr)
            return False

    except Exception as e:
        print(f"💥 向 {phone_number} 发送时异常: {e}", file=sys.stderr)
        return False


def send_mod_count_sms(todaycount: int, yesterdaycount: int, increment: int) -> Tuple[int, int]:
    """
    批量发送 MOD 统计短信通知。
    
    从 phonelist.txt 读取所有有效号码，逐个发送。
    返回 (成功数, 总尝试数)
    
    注意：函数名保留为 send_mod_count_sms 以兼容主脚本调用，
          但行为已改为批量发送。
    """
    # 获取凭证
    access_key_id = os.getenv("ALIBABA_CLOUD_ACCESS_KEY_ID")
    access_key_secret = os.getenv("ALIBABA_CLOUD_ACCESS_KEY_SECRET")

    if not access_key_id or not access_key_secret:
        raise EnvironmentError(
            "环境变量缺失：请确保设置了 ALIBABA_CLOUD_ACCESS_KEY_ID 和 ALIBABA_CLOUD_ACCESS_KEY_SECRET。"
        )

    # 加载号码
    phone_numbers = _load_phone_numbers()
    if not phone_numbers:
        return 0, 0

    # 创建客户端（复用，避免重复初始化）
    client = _create_client(access_key_id, access_key_secret)

    # 批量发送
    success_count = 0
    total = len(phone_numbers)
    for i, phone in enumerate(phone_numbers, 1):
        print(f"📱 [{i}/{total}] 正在向 {phone} 发送通知...")
        if _send_single_sms(client, phone, todaycount, yesterdaycount, increment):
            success_count += 1
        # 防止触发频率限制（阿里云建议 ≥1秒）
        if i < total:
            time.sleep(1)

    return success_count, total