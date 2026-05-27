#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FB广告消耗数据转换脚本（读取xlsx版本）
直接读取xlsx文件，使用openpyxl

课包匹配逻辑：
1. 先按账户名称关键字匹配
2. 再按广告名称关键字匹配
3. 顺序优先
"""

from openpyxl import load_workbook
import csv
import sys
from collections import defaultdict

# ============ 配置参数 ============
USD_TO_CNY = 6.8579
TAX_FACTOR = 1.03
OPERATOR = '欧莹莹'
RULES_FILE = 'FB账户消耗底表.xlsx'
SOURCE_FILE = '思维全-每日导消耗.xlsx'
OUTPUT_FILE = '课包城市消耗数据回写_输出.xlsx'
# ==================================

# 课包关键字匹配规则（按顺序匹配，先匹配到的优先）
# 格式：(匹配字段, 关键字, 课包名称)
# 匹配字段：'account'=账户名称, 'ad'=广告名称
PACKAGE_RULES = [
    ('account', '飞书9户', 'FB-思维&美术-0元31节课'),
    ('account', '飞书23户', 'FB-思维&美术-0元31节课'),
    ('account', '飞书8户', 'FB-思维&美术-0元31节课'),
    ('account', '飞书3户', 'FB-思维&美术-0元31节课'),
    ('account', '飞书27户', 'FB-思维&美术-0元31节课'),
    ('ad', '伊', 'KOLTW-【送4节课】-伊萊家-Sansan-思维&无'),
    ('ad', '邓明仪直播', 'KOLHK-香港-鄧明儀直播-Shukei-思维&无'),
    ('ad', '陈琪直播', 'KOLHK-香港-陈琪直播-Shukei-思维&无-原视频'),
    ('ad', '陈琪图片', 'KOLHK-香港-陈琪直播-Shukei-思维&无-原视频'),
    ('ad', '木木', 'KOLTW-【送4节课】-木木-Sansan-思维&无'),
    ('ad', '蔓乐爸', 'KOLTW-【送4节课】-蔓樂爸-Sansan-思维&无'),
    ('ad', '吕慧仪直播混剪', 'KOLHK-香港-呂慧儀直播-Shukei-思维&无-混剪-原贴'),
    ('ad', '邓明仪图片', 'KOLHK-香港-專家鄧明儀图片-Shukei-思维&无'),
    ('ad', '陈琪', 'KOLHK-香港-陈琪-Shukei-思维&美术'),
    ('ad', 'KiFan', 'SEOTW-【送8节课】-KiFanLan-自招-Sansan-思维&无'),
    ('ad', '薇薇', 'KOLTW-【送8节课】-薇薇-Sansan-思维&无'),
    ('ad', 'Bob', 'KOLHK-香港-Bob林盛斌-Shukei-思维&美术01'),
    ('ad', 'Jerilyn', 'KOL本地-SG-Jerilyn Moon-【8节课】-hjy-思维&无'),
    ('ad', '抠妈', 'KOLTW-【送8节课】-摳媽與摳比-Sansan-思维&无'),
    ('ad', '蔓樂爸', 'KOLTW-【送4节课】-蔓樂爸-Sansan-思维&无'),
    ('ad', '陈怡君', 'SEOTW-【送8节课】-陳怡君-自招-Sansan-思维&无'),
    ('ad', '钟嘉欣图片', 'KOLHK-香港-钟嘉欣图片-Shukei-思维&无-自创建'),
    ('ad', '小姐', 'KOLTW-【送8节课】-小姐不熙娣-Sansan-思维&无'),
    ('ad', '钟嘉欣', 'KOLHK-香港-钟嘉欣视频1-Shukei-思维&无'),
    ('ad', '蔡雪莹', 'KOLHK-香港-蔡雪莹-Shukei-思维&无'),
    ('ad', '陈展鹏', 'KOLHK-香港-陳展鵬-Shukei-思维&无'),
    ('ad', '姐妹会', 'KOLTW-【送8节课】-WTO姐妹會-Sansan-思维&无'),
    ('ad', '张棋惠', 'KOLTW-【送百元券加送8节课】-張棋惠-Sansan-思维&无'),
    ('ad', '家蔚', 'KOLHK-香港-周家蔚-Shukei-思维&无'),
    ('ad', '邓明仪', 'KOLHK-香港-專家鄧明儀-Shukei-思维&无'),
    ('account', '飞书31户', 'KOL本地-SG- Melissa Celestine Koh-【8节课】-hjy-思维&无'),
    ('ad', '红豆妹', 'KOLTW-【送百元券加送8节课】-紅豆妹-Sansan-思维&无'),
    ('ad', '马力欧', 'KOLTW-【送百元券加送8节课】-馬力歐-Sansan-思维&无'),
    ('ad', '馬力歐', 'KOLTW-【送百元券加送8节课】-馬力歐-Sansan-思维&无'),
    ('ad', '陈凯琳', 'KOLHK-香港-陳凱琳-Shukei-思维&无'),
    ('ad', '直播黃小柔', 'KOLTW-【送8节课】-直播黃小柔-Sansan-思维&无'),
    ('ad', '小柔直播', 'KOLTW-【送8节课】-直播黃小柔-Sansan-思维&无'),
    ('ad', '刘轩老', 'KOLTW-【送8节课】-劉軒-Sansan-思维&无'),
    ('ad', '刘轩', 'KOLTW-【送8节课】-劉軒-Sansan-思维&无'),
    ('ad', '孙慧雪', 'KOLHK-香港-孙慧雪-Shukei-思维&美术'),
    ('ad', '摳媽', 'KOLTW-【送8节课】-摳媽與摳比-Sansan-思维&无'),
    ('ad', '李新', 'KOLTW-【送百元券加送8节课】-李新-Sansan-思维&无'),
    ('ad', '小朗爸', 'KOLTW-【送百元券加送8节课】-小朗爸-Sansan-思维&无'),
    ('ad', '20260415-卖点(立减)', 'KOL本地-SG- Melissa Celestine Koh-【8节课】-hjy-思维&无'),
    ('ad', '吕慧仪', ' KOLHK-香港-呂慧儀直播-Shukei-思维&无-原视频'),
    ('ad', '呂慧儀', 'KOLHK-香港-呂慧儀直播-Shukei-思维&无-原视频'),
    ('ad', '丧尸', 'KOLTW-【送百元券加送8节课】-喪屍老爸 -Sansan-思维&无'),
    ('ad', 'Cherrie', 'SEOTW-【送8节课】-Cherrie-自招-Sansan-思维&无'),
    ('ad', 'Melissa', 'KOL本地-SG- Melissa Celestine Koh-【8节课】-hjy-思维&无'),
    ('ad', 'genesis', 'KOL本地-SG-The Genesis Family-【8节课】-hjy-思维&无'),
    ('ad', '三位KOL', 'KOL本地-SG- Melissa Celestine Koh-【8节课】-hjy-思维&无'),
    ('ad', 'thatmomo', 'KOL本地-SG-thatmomoffour-【8节课】-hjy-思维&无'),
    ('ad', '米菲', 'KOLTW-【送百元券加送8节课】-米菲和米粒MifaMily -Sansan-思维&无'),
    ('ad', '王思佳', 'KOLTW-【送百元券加送8节课】-王思佳 -Sansan-思维&无'),
    ('ad', '周家蔚', 'KOLHK-香港-周家蔚-Shukei-思维&无'),
    ('ad', '周嘉蔚', 'KOLHK-香港-周家蔚-Shukei-思维&无'),
    ('ad', 'Ermy', 'KOL本地-SG-Ermy Lie-【8节课】-hjy-思维&无'),
    ('ad', 'Edwina', 'KOL本地-SG-Edwina Pariwono-【8节课】-hjy-思维&无'),
    ('ad', '小柔', 'KOLTW-【送8节课】-黃小柔-Sansan-思维&无'),
]

def match_package(account_name, ad_name):
    """根据关键字规则匹配课包"""
    for field, keyword, package in PACKAGE_RULES:
        if field == 'account' and keyword in account_name:
            return package
        elif field == 'ad' and keyword in ad_name:
            return package
    return '-'  # 未匹配到

def load_account_mapping(wb):
    """从'平台&账户&课包映射关系'sheet加载 账户→平台/投放账户 映射"""
    ws = wb['平台&账户&课包映射关系']
    mapping = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        account_name = row[0]
        if account_name:
            mapping[str(account_name).strip()] = {
                '投放平台': str(row[1] or '').strip(),
                '投放账户': str(row[2] or '').strip(),
            }

    # 投放平台覆盖规则（优先级高于xlsx映射表）
    PLATFORM_OVERRIDES = {
        '飞书7户': 'KOLHK',
        '飞书31户': 'KOL本地',
    }
    for key, platform in PLATFORM_OVERRIDES.items():
        for acc_name in list(mapping.keys()):
            if key in acc_name:
                mapping[acc_name]['投放平台'] = platform

    return mapping

def load_country_mapping(wb):
    """从'地区名称对照表'sheet加载国家映射"""
    ws = wb['地区名称对照表']
    mapping = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        abbr = row[1]
        english = row[2]
        if abbr and english and str(abbr).strip() != '-':
            mapping[str(abbr).strip()] = str(english).strip()

    return mapping

def load_source_data(filename):
    """加载源数据xlsx"""
    wb = load_workbook(filename, data_only=True, read_only=True)
    ws = wb.active

    headers = None
    rows = []
    for row in ws.iter_rows(values_only=True):
        if headers is None:
            headers = [str(h).strip() if h else '' for h in row]
        else:
            rows.append(dict(zip(headers, row)))

    wb.close()
    return rows

def transform_and_pivot(source_rows, account_map, country_map):
    """转换+透视"""
    pivot = defaultdict(lambda: {'曝光': 0, '点击': 0, '消耗': 0.0})

    unmapped_accounts = set()
    unmapped_countries = set()
    unmapped_packages = 0
    mapped_count = 0

    for row in source_rows:
        account_name = str(row.get('账户名称', '') or '').strip()
        country_code = str(row.get('国家/地区', '') or '').strip()
        ad_name = str(row.get('广告名称', '') or '').strip()

        # 国家映射
        country_full = country_map.get(country_code, '')
        if not country_full:
            unmapped_countries.add(country_code)
            continue

        # 课包：用关键字匹配
        package = match_package(account_name, ad_name)
        if package == '-':
            unmapped_packages += 1
            continue

        # 平台和投放账户：从映射表获取
        acc_info = account_map.get(account_name, None)
        if acc_info:
            platform = acc_info['投放平台']
            account_std = acc_info['投放账户']
        else:
            unmapped_accounts.add(account_name)
            continue

        # 消耗 = USD * 6.8579 / 1.03
        try:
            usd = float(row.get('已花费金额 (USD)', 0) or 0)
        except (ValueError, TypeError):
            usd = 0.0
        cost = usd * USD_TO_CNY / TAX_FACTOR

        # 曝光和点击
        try:
            impressions = int(float(row.get('展示次数', 0) or 0))
        except (ValueError, TypeError):
            impressions = 0
        try:
            clicks = int(float(row.get('点击量（全部）', 0) or 0))
        except (ValueError, TypeError):
            clicks = 0

        # 透视key
        date_val = str(row.get('单日', '') or '').strip()
        key = (date_val, platform, account_std, '无', package, country_full)

        pivot[key]['曝光'] += impressions
        pivot[key]['点击'] += clicks
        pivot[key]['消耗'] += cost
        mapped_count += 1

    # 打印警告
    if unmapped_accounts:
        print(f"\n  警告：{len(unmapped_accounts)} 个账户未匹配平台映射：")
        for a in sorted(unmapped_accounts)[:5]:
            print(f"    - {a}")
    if unmapped_countries:
        print(f"\n  警告：{len(unmapped_countries)} 个国家代码未匹配：")
        for c in sorted(unmapped_countries):
            print(f"    - {c}")
    if unmapped_packages > 0:
        print(f"\n  警告：{unmapped_packages} 行未匹配到课包")

    print(f"\n  成功匹配：{mapped_count} 行")
    return pivot

def save_result(pivot):
    """保存结果为xlsx，表头从第3行开始"""
    from openpyxl import Workbook
    import os

    # 先检查 Excel 锁文件（~$ 开头）
    lock_file = os.path.join(os.path.dirname(os.path.abspath(OUTPUT_FILE)) or '.',
                             '~$' + os.path.basename(OUTPUT_FILE))
    if os.path.exists(lock_file):
        print(f"\n错误：文件被 Excel 打开（检测到锁文件 {os.path.basename(lock_file)}）")
        print("请关闭 Excel 后重新运行")
        sys.exit(4)

    wb = Workbook()
    ws = wb.active
    ws.title = 'sheet1'

    header = ['日期', '投放平台', '投放账户', '广告位', '课包',
              '国家英文名称', '曝光', '点击', '消耗', '更新操作人']

    # 表头写在第3行
    for col, h in enumerate(header, 1):
        ws.cell(row=3, column=col, value=h)

    # 数据从第4行开始
    row_idx = 4
    for key, values in sorted(pivot.items()):
        data = list(key) + [
            values['曝光'],
            values['点击'],
            round(values['消耗'], 6),
            OPERATOR
        ]
        for col, val in enumerate(data, 1):
            ws.cell(row=row_idx, column=col, value=val)
        row_idx += 1

    try:
        wb.save(OUTPUT_FILE)
    except PermissionError:
        print(f"\n错误：无法写入 {OUTPUT_FILE}，文件被占用（Excel 打开了？）")
        print("请关闭 Excel 后重新运行")
        sys.exit(4)
    return len(pivot)

def main():
    print("=" * 50)
    print("FB广告消耗数据转换工具")
    print(f"公式：消耗 = USD * {USD_TO_CNY} / {TAX_FACTOR}")
    print("=" * 50)

    try:
        # 1. 加载规则表
        print("\n加载规则表...")
        rules_wb = load_workbook(RULES_FILE, data_only=True, read_only=True)

        account_map = load_account_mapping(rules_wb)
        print(f"  账户→平台映射：{len(account_map)} 条")

        country_map = load_country_mapping(rules_wb)
        print(f"  国家映射：{len(country_map)} 条")

        rules_wb.close()

        # 2. 加载源数据
        print(f"\n加载源数据：{SOURCE_FILE}")
        source_rows = load_source_data(SOURCE_FILE)
        print(f"  源数据：{len(source_rows)} 行")

        # 3. 转换+透视
        print("\n转换并透视...")
        pivot = transform_and_pivot(source_rows, account_map, country_map)
        print(f"  输出行数：{len(pivot)}")

        if not pivot:
            print("\n错误：没有数据被转换")
            sys.exit(1)

        # 4. 保存
        count = save_result(pivot)

        # 5. 统计
        total_cost = sum(v['消耗'] for v in pivot.values())
        total_imp = sum(v['曝光'] for v in pivot.values())
        total_click = sum(v['点击'] for v in pivot.values())

        print("\n" + "=" * 50)
        print(f"完成！输出文件：{OUTPUT_FILE}")
        print(f"  输出行数：{count}")
        print(f"  总曝光：{total_imp:,}")
        print(f"  总点击：{total_click:,}")
        print(f"  总消耗(CNY)：{total_cost:,.2f}")
        print("=" * 50)

    except FileNotFoundError as e:
        print(f"\n错误：找不到文件 - {e}")
        sys.exit(2)
    except Exception as e:
        print(f"\n错误：{e}")
        import traceback
        traceback.print_exc()
        sys.exit(3)

if __name__ == '__main__':
    main()
