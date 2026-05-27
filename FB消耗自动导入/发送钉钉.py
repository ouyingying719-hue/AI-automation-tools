#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
钉钉通知脚本 - 发送任务结果到钉钉群

用法：
    py 发送钉钉.py 成功 "FB消耗导入完成\n总消耗：1234元"
    py 发送钉钉.py 失败 "FB消耗导入失败\n错误：会话已过期"
"""

import urllib.request
import urllib.error
import json
import os
import sys

ENV_FILE = '.env'
KEYWORD = 'BI'  # 钉钉机器人安全关键词，消息中必须包含


def load_env(key):
    """读.env"""
    if not os.path.exists(ENV_FILE):
        return None
    with open(ENV_FILE, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line.startswith(f'{key}='):
                return line.split('=', 1)[1].strip()
    return None


def send_dingtalk(title, content, at_all=False):
    """发送钉钉消息（markdown格式）"""
    webhook = load_env('DINGTALK_WEBHOOK')
    if not webhook:
        print("未配置 DINGTALK_WEBHOOK，跳过通知")
        return False

    # 确保消息包含安全关键词（钉钉机器人要求）
    if KEYWORD not in title and KEYWORD not in content:
        title = f"{KEYWORD} - {title}"

    payload = {
        "msgtype": "markdown",
        "markdown": {
            "title": title,
            "text": content,
        },
        "at": {
            "isAtAll": at_all,
        },
    }

    data = json.dumps(payload, ensure_ascii=False).encode('utf-8')
    req = urllib.request.Request(
        webhook,
        data=data,
        method='POST',
        headers={'Content-Type': 'application/json; charset=utf-8'},
    )

    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            result = json.loads(resp.read().decode('utf-8'))
            if result.get('errcode') == 0:
                print("钉钉通知发送成功")
                return True
            else:
                print(f"钉钉通知失败：{result}")
                return False
    except Exception as e:
        print(f"钉钉通知发送异常：{e}")
        return False


def main():
    if len(sys.argv) < 3:
        # 默认测试消息
        send_dingtalk("测试", "BI 通知测试 - 配置正常")
        return

    status = sys.argv[1]
    content = sys.argv[2]

    if status == '成功':
        title = "BI消耗导入成功"
        text = f"## ✅ {title}\n\n{content}"
    else:
        title = "BI消耗导入失败"
        text = f"## ❌ {title}\n\n{content}"

    send_dingtalk(title, text, at_all=(status == '失败'))


if __name__ == '__main__':
    main()
