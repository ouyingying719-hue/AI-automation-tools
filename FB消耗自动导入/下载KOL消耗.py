#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从 FB Marketing API 下载KOL广告消耗数据（扩展字段）
输出格式与"思维美术广告消耗-KOL-导入BI.xlsx"一致

用法：
    py 下载KOL消耗.py              # 下载昨天的数据
    py 下载KOL消耗.py 2026-05-22   # 指定日期
"""

import urllib.request
import urllib.parse
import urllib.error
import json
import sys
import time
import os
from datetime import datetime, timedelta
from openpyxl import Workbook

ENV_FILE = '.env'
OUTPUT_FILE = '思维美术广告消耗-KOL-导入BI_输出.xlsx'
API_VERSION = 'v21.0'

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

# action_type 关键字映射
LINK_CLICK_KEYS = {'link_click'}
LEAD_KEYS = {'lead', 'leadgen.other', 'offsite_conversion.fb_pixel_lead'}
MESSAGE_KEYS = {
    'onsite_conversion.messaging_conversation_started_7d',
    'onsite_conversion.total_messaging_connection',
}
PURCHASE_KEYS = {
    'purchase', 'omni_purchase', 'offsite_conversion.fb_pixel_purchase',
}


def load_token():
    with open(ENV_FILE, 'r', encoding='utf-8') as f:
        for line in f:
            if line.startswith('FB_ACCESS_TOKEN='):
                return line.split('=', 1)[1].strip()
    return None


def fb_api_get(url, max_retries=3):
    last_err = None
    for attempt in range(max_retries):
        try:
            with urllib.request.urlopen(url, timeout=60) as response:
                return json.loads(response.read().decode('utf-8')), None
        except urllib.error.HTTPError as e:
            error_body = e.read().decode('utf-8')
            last_err = f"HTTP {e.code}: {error_body[:200]}"
            if 400 <= e.code < 500:
                return None, last_err
        except Exception as e:
            last_err = f"{type(e).__name__}: {e}"
        if attempt < max_retries - 1:
            wait = 2 ** attempt
            print(f"    重试中（{attempt+1}/{max_retries}），等{wait}秒...")
            time.sleep(wait)
    return None, last_err


def sum_actions(actions, key_set):
    """从actions数组里把符合关键字的求和"""
    if not actions:
        return 0
    total = 0
    for a in actions:
        if a.get('action_type') in key_set:
            try:
                total += int(float(a.get('value', 0)))
            except (ValueError, TypeError):
                pass
    return total


def first_action_value(actions, key_set):
    """从 video_pXX_watched_actions 数组里取第一个匹配项的值"""
    if not actions:
        return 0
    for a in actions:
        if a.get('action_type') == 'video_view':
            try:
                return int(float(a.get('value', 0)))
            except (ValueError, TypeError):
                return 0
    return 0


def fetch_insights(account_id, since, until, token):
    """拉某账户的扩展字段"""
    fields = [
        'account_name',
        'campaign_id', 'campaign_name',
        'adset_id', 'adset_name',
        'ad_id', 'ad_name',
        'reach', 'impressions', 'clicks',
        'spend',
        'actions',
        'video_p25_watched_actions',
        'video_p50_watched_actions',
        'video_p75_watched_actions',
        'video_p95_watched_actions',
        'video_p100_watched_actions',
    ]
    base_url = f"https://graph.facebook.com/{API_VERSION}/{account_id}/insights"
    params = {
        'access_token': token,
        'fields': ','.join(fields),
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


def fetch_adset_start_dates(adset_ids, token):
    """批量获取广告组的 start_time，返回 {adset_id: 'YYYY-MM-DD'}"""
    result = {}
    ids = list(adset_ids)
    BATCH = 50  # FB API批量请求上限
    for i in range(0, len(ids), BATCH):
        batch = ids[i:i+BATCH]
        params = {
            'access_token': token,
            'ids': ','.join(batch),
            'fields': 'start_time',
        }
        url = f"https://graph.facebook.com/{API_VERSION}/?{urllib.parse.urlencode(params)}"
        data, err = fb_api_get(url)
        if err:
            print(f"    取广告组开始日期失败：{err}")
            continue
        for adset_id, info in data.items():
            start_time = info.get('start_time', '')
            if start_time:
                # ISO格式 "2026-05-18T07:00:00-0700" → "2026-05-18"
                result[adset_id] = start_time[:10]
    return result


def main():
    print("=" * 60)
    print("FB广告KOL消耗数据下载（扩展字段）")
    print("=" * 60)

    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    token = load_token()
    if not token:
        print("错误：未找到 FB_ACCESS_TOKEN")
        sys.exit(1)

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
    print(f"insights总计：{len(all_rows)} 行")

    if failed_accounts:
        print(f"\n[警告] {len(failed_accounts)} 个账户拉取失败：")
        for acc_id, err in failed_accounts:
            print(f"    - {acc_id}: {err}")

    if not all_rows:
        print("没有数据，不生成文件")
        sys.exit(2)

    # 取所有 adset_id 的 start_time
    print("\n获取广告组开始日期...")
    unique_adsets = {r.get('adset_id') for r in all_rows if r.get('adset_id')}
    print(f"  广告组数：{len(unique_adsets)}")
    adset_start = fetch_adset_start_dates(unique_adsets, token)
    print(f"  取到开始日期：{len(adset_start)} 个")

    # 检查锁文件
    lock_file = '~$' + OUTPUT_FILE
    if os.path.exists(lock_file):
        print(f"\n错误：{OUTPUT_FILE} 被 Excel 打开（{lock_file}）")
        print("请关闭 Excel 后重新运行")
        sys.exit(4)

    # 写入xlsx，列名跟样板表保持一致
    wb = Workbook()
    ws = wb.active
    ws.title = 'Raw Data Report'

    headers = [
        '单日', '国家/地区', '账户名称',
        '广告系列编号', '广告系列名称',
        '广告组编号', '广告组名称',
        '广告编号', '广告名称',
        '覆盖人数', '展示次数', '点击量（全部）',
        '货币', '已花费金额 (USD)',
        '链接点击量', '潜在客户人数', '消息对话发起次数', '购物次数',
        '视频播放进度达 25% 的次数', '视频播放进度达 50% 的次数',
        '视频播放进度达 75% 的次数', '视频播放进度达 95% 的次数',
        '视频播放进度达 100% 的次数',
        '开始日期', '结束日期', '报告开始日期', '报告结束日期',
    ]
    ws.append(headers)

    def to_int(v):
        try:
            return int(float(v or 0))
        except (ValueError, TypeError):
            return 0

    def to_float(v):
        try:
            return float(v or 0)
        except (ValueError, TypeError):
            return 0.0

    for r in all_rows:
        actions = r.get('actions', [])
        link_click = sum_actions(actions, LINK_CLICK_KEYS)
        leads = sum_actions(actions, LEAD_KEYS)
        msgs = sum_actions(actions, MESSAGE_KEYS)
        purchases = sum_actions(actions, PURCHASE_KEYS)

        v25 = first_action_value(r.get('video_p25_watched_actions'), None)
        v50 = first_action_value(r.get('video_p50_watched_actions'), None)
        v75 = first_action_value(r.get('video_p75_watched_actions'), None)
        v95 = first_action_value(r.get('video_p95_watched_actions'), None)
        v100 = first_action_value(r.get('video_p100_watched_actions'), None)

        adset_id = r.get('adset_id', '')
        start_date = adset_start.get(adset_id, '')

        ws.append([
            r.get('date_start', ''),                          # 单日
            r.get('country', ''),                             # 国家/地区
            r.get('account_name', ''),                        # 账户名称
            r.get('campaign_id', ''),                         # 广告系列编号
            r.get('campaign_name', ''),                       # 广告系列名称
            adset_id,                                         # 广告组编号
            r.get('adset_name', ''),                          # 广告组名称
            r.get('ad_id', ''),                               # 广告编号
            r.get('ad_name', ''),                             # 广告名称
            to_int(r.get('reach', 0)),                        # 覆盖人数
            to_int(r.get('impressions', 0)),                  # 展示次数
            to_int(r.get('clicks', 0)),                       # 点击量（全部）
            'USD',                                            # 货币
            to_float(r.get('spend', 0)),                      # 已花费金额 (USD)
            link_click,                                       # 链接点击量
            leads,                                            # 潜在客户人数
            msgs,                                             # 消息对话发起次数
            purchases,                                        # 购物次数
            v25, v50, v75, v95, v100,                         # 视频播放进度
            start_date,                                       # 开始日期
            'Ongoing',                                        # 结束日期（FB没有，给Ongoing占位）
            target_date,                                      # 报告开始日期
            target_date,                                      # 报告结束日期
        ])

    try:
        wb.save(OUTPUT_FILE)
    except PermissionError:
        print(f"\n错误：无法写入 {OUTPUT_FILE}（被占用）")
        sys.exit(4)
    print(f"\n已保存：{OUTPUT_FILE}")
    print("=" * 60)


if __name__ == '__main__':
    main()
