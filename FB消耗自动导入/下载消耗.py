#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从FB Marketing API下载消耗数据
输出格式与"思维全-每日导消耗.xlsx"一致

用法：
    py 下载消耗.py              # 下载昨天的数据
    py 下载消耗.py 2026-05-19   # 下载指定日期
"""

import urllib.request
import urllib.parse
import urllib.error
import json
import sys
from datetime import datetime, timedelta
from openpyxl import Workbook

# ============ 配置 ============
ENV_FILE = '.env'
OUTPUT_FILE = '思维全-每日导消耗.xlsx'
API_VERSION = 'v21.0'
# ==============================

# 需要拉取的广告账户ID
AD_ACCOUNTS = [
    'act_468253789344241',   # FB-飞书7户-思维&美术-港澳-KOL-post
    'act_442457935507062',   # FB-飞书9户-思维&美术-优质-表单
    'act_366864216464093',   # FB-飞书8户-思维&美术-混投-表单
    'act_1816783745480753',  # FB-飞书3户-思维&美术-港澳-WA
    'act_627543006449377',   # FB-飞书18户-思维&美术-台湾-KOL
    'act_1158469106256042',  # FB-飞书27户-思维&美术-台湾-H5
    'act_2107766700052928',  # FB-飞书23户-思维&美术-新加坡-表单
    'act_1416198419836342',  # FB-飞书31户-思维&美术-新加坡-KOL
]


def load_token():
    with open(ENV_FILE, 'r', encoding='utf-8') as f:
        for line in f:
            if line.startswith('FB_ACCESS_TOKEN='):
                return line.split('=', 1)[1].strip()
    return None


def fb_api_get(url, max_retries=3):
    """带重试的GET请求，返回(result, error_msg)"""
    import time
    last_err = None
    for attempt in range(max_retries):
        try:
            with urllib.request.urlopen(url, timeout=60) as response:
                return json.loads(response.read().decode('utf-8')), None
        except urllib.error.HTTPError as e:
            error_body = e.read().decode('utf-8')
            last_err = f"HTTP {e.code}: {error_body[:200]}"
            # 4xx错误不重试
            if 400 <= e.code < 500:
                return None, last_err
        except Exception as e:
            last_err = f"{type(e).__name__}: {e}"
        if attempt < max_retries - 1:
            wait = 2 ** attempt  # 1s, 2s, 4s
            print(f"    重试中（{attempt+1}/{max_retries}），等{wait}秒...")
            time.sleep(wait)
    return None, last_err


def fetch_insights(account_id, since, until, token):
    """获取一个账户在指定日期范围的insights，按country和ad分组
    返回 (rows, error_msg)，error_msg为None表示成功
    """
    base_url = f"https://graph.facebook.com/{API_VERSION}/{account_id}/insights"
    params = {
        'access_token': token,
        'fields': 'account_name,ad_name,spend,impressions,clicks',
        'level': 'ad',
        'breakdowns': 'country',
        'time_range': json.dumps({'since': since, 'until': until}),
        'time_increment': '1',
        'limit': '500',
    }
    url = f"{base_url}?{urllib.parse.urlencode(params)}"

    all_rows = []
    while url:
        result, err = fb_api_get(url)
        if err:
            return all_rows, err
        all_rows.extend(result.get('data', []))
        url = result.get('paging', {}).get('next')
    return all_rows, None


def main():
    print("=" * 60)
    print("FB广告消耗数据下载")
    print("=" * 60)

    token = load_token()
    if not token:
        print("错误：未找到 FB_ACCESS_TOKEN")
        sys.exit(1)

    # 日期：参数指定，否则取昨天
    if len(sys.argv) > 1:
        target_date = sys.argv[1]
    else:
        target_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')

    print(f"\n下载日期：{target_date}")
    print(f"账户数：{len(AD_ACCOUNTS)}")

    all_rows = []
    failed_accounts = []
    empty_accounts = []
    for i, acc_id in enumerate(AD_ACCOUNTS, 1):
        print(f"\n[{i:2d}/{len(AD_ACCOUNTS)}] {acc_id}")
        rows, err = fetch_insights(acc_id, target_date, target_date, token)
        if err:
            print(f"    → 失败：{err}")
            failed_accounts.append((acc_id, err))
        elif rows:
            print(f"    → {len(rows)} 行 (账户：{rows[0].get('account_name', '')})")
        else:
            print(f"    → 无数据")
            empty_accounts.append(acc_id)
        all_rows.extend(rows)

    print(f"\n{'='*60}")
    print(f"总计：{len(all_rows)} 行")

    if failed_accounts:
        print(f"\n[警告] {len(failed_accounts)} 个账户拉取失败，需要手动检查：")
        for acc_id, err in failed_accounts:
            print(f"    - {acc_id}: {err}")

    if empty_accounts:
        print(f"\n[提示] {len(empty_accounts)} 个账户无数据（可能当天确实没投放）：")
        for acc_id in empty_accounts:
            print(f"    - {acc_id}")

    if not all_rows:
        print("错误：没有数据，所有账户都拉取失败或无数据，不生成文件")
        sys.exit(2)

    # 检查 Excel 锁文件
    import os as _os
    lock_file = '~$' + OUTPUT_FILE
    if _os.path.exists(lock_file):
        print(f"\n错误：{OUTPUT_FILE} 被 Excel 打开（检测到锁文件 {lock_file}）")
        print("请关闭 Excel 后重新运行")
        sys.exit(4)

    # 写入xlsx，列名和原表保持一致
    wb = Workbook()
    ws = wb.active
    ws.title = '消耗数据'

    headers = ['单日', '账户名称', '广告名称', '国家/地区',
               '已花费金额 (USD)', '展示次数', '点击量（全部）']
    ws.append(headers)

    for row in all_rows:
        try:
            spend = float(row.get('spend', 0) or 0)
        except (ValueError, TypeError):
            spend = 0.0
        try:
            impressions = int(row.get('impressions', 0) or 0)
        except (ValueError, TypeError):
            impressions = 0
        try:
            clicks = int(row.get('clicks', 0) or 0)
        except (ValueError, TypeError):
            clicks = 0

        ws.append([
            row.get('date_start', ''),
            row.get('account_name', ''),
            row.get('ad_name', ''),
            row.get('country', ''),
            spend,
            impressions,
            clicks,
        ])

    try:
        wb.save(OUTPUT_FILE)
    except PermissionError:
        print(f"\n错误：无法写入 {OUTPUT_FILE}，文件被占用（Excel 打开了？）")
        print("请关闭 Excel 后重新运行")
        sys.exit(4)
    print(f"已保存：{OUTPUT_FILE}")
    print("=" * 60)

    # 部分账户失败时不中止流程（数据是有的），但在输出里留下"失败账户"行
    # run_all.py 会把这条信息带到钉钉里，作为 ⚠️ 警告


if __name__ == '__main__':
    main()
