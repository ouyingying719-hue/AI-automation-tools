#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KOL消耗 → 投放账户数据回写填报模板 格式转换

规则：
- 广告素材id、广告素材名称、笔记ID：固定为 0
- 主投学科：固定 思维
- 日期：取"单日"
- 区域细分：国家代码 → 中文名（来自 FB账户消耗底表.xlsx）
- 消耗：已花费金额(USD) × 6.8579 / 1.03
- 投放平台：
    * 飞书7户 → KOLHK
    * 飞书31户 → KOL本地
    * 飞书18户 → 广告名称含"陈怡君" 返回 SEOTW，否则 直购（FB）
    * 其他 → 直购（FB）
- 上传日期：今天
- 开始投放日期：源表的"开始日期"
- 其他字段直接对应
"""

import os
import sys
from datetime import datetime
from openpyxl import load_workbook, Workbook

USD_TO_CNY = 6.8579
TAX_FACTOR = 1.03
SUBJECT = '思维'

RULES_FILE = 'FB账户消耗底表.xlsx'
SOURCE_FILE = '思维美术广告消耗-KOL-导入BI_输出.xlsx'
OUTPUT_FILE = '投放账户数据回写填报模板_输出.xlsx'


def load_country_map(wb):
    """国家代码 → 中文名"""
    ws = wb['地区名称对照表']
    m = {}
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        cn = r[0]
        code = r[1]
        if code and cn and str(code).strip() != '-':
            m[str(code).strip()] = str(cn).strip()
    return m


def decide_platform(account_name, ad_name):
    """投放平台判断"""
    acc = account_name or ''
    ad = ad_name or ''
    if '飞书7户' in acc:
        return 'KOLHK'
    if '飞书31户' in acc:
        return 'KOL本地'
    if '飞书18户' in acc:
        if '陈怡君' in ad:
            return 'SEOTW'
        return 'KOLTW'
    return '直购（FB）'


def load_source(filename):
    wb = load_workbook(filename, data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return [], []
    headers = [str(h).strip() if h else '' for h in rows[0]]
    data = [dict(zip(headers, r)) for r in rows[1:]]
    return headers, data


def to_int(v):
    if v is None or v == '':
        return 0
    try:
        return int(float(v))
    except (ValueError, TypeError):
        return 0


def to_float(v):
    if v is None or v == '':
        return 0.0
    try:
        return float(v)
    except (ValueError, TypeError):
        return 0.0


def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    print("=" * 50)
    print("KOL消耗 → 投放账户数据回写填报模板")
    print("=" * 50)

    if not os.path.exists(SOURCE_FILE):
        print(f"错误：找不到源文件 {SOURCE_FILE}")
        sys.exit(1)
    if not os.path.exists(RULES_FILE):
        print(f"错误：找不到规则表 {RULES_FILE}")
        sys.exit(1)

    # 锁文件检查
    lock_file = '~$' + OUTPUT_FILE
    if os.path.exists(lock_file):
        print(f"\n错误：{OUTPUT_FILE} 被 Excel 打开（{lock_file}）")
        print("请关闭 Excel 后重新运行")
        sys.exit(4)

    print("\n加载规则表...")
    rules_wb = load_workbook(RULES_FILE, data_only=True, read_only=True)
    country_map = load_country_map(rules_wb)
    print(f"  国家映射：{len(country_map)} 条")
    rules_wb.close()

    print(f"\n加载源数据：{SOURCE_FILE}")
    headers, source_rows = load_source(SOURCE_FILE)
    print(f"  源数据：{len(source_rows)} 行")
    if not source_rows:
        print("错误：源数据为空")
        sys.exit(2)

    upload_date = datetime.now().strftime('%Y-%m-%d')
    unmapped_countries = set()
    out_rows = []
    total_cost = 0.0

    for r in source_rows:
        country_code = str(r.get('国家/地区', '') or '').strip()
        country_cn = country_map.get(country_code, '')
        if not country_cn:
            unmapped_countries.add(country_code)
            country_cn = country_code  # 未映射的保留原代码，方便后续排查

        account_name = str(r.get('账户名称', '') or '')
        ad_name = str(r.get('广告名称', '') or '')
        platform = decide_platform(account_name, ad_name)

        usd = to_float(r.get('已花费金额 (USD)', 0))
        cost = usd * USD_TO_CNY / TAX_FACTOR
        total_cost += cost

        start_date = r.get('开始日期', '') or ''
        # 兼容 datetime
        if hasattr(start_date, 'strftime'):
            start_date = start_date.strftime('%Y-%m-%d')
        else:
            start_date = str(start_date)[:10]

        date_val = r.get('单日', '') or ''
        if hasattr(date_val, 'strftime'):
            date_val = date_val.strftime('%Y-%m-%d')
        else:
            date_val = str(date_val)[:10]

        out_rows.append([
            date_val,                                  # 日期
            SUBJECT,                                   # 主投学科
            platform,                                  # 平台
            country_cn,                                # 区域细分
            account_name,                              # 投放账户
            r.get('广告系列编号', '') or '',           # 广告计划id
            r.get('广告系列名称', '') or '',           # 广告计划名称
            r.get('广告组编号', '') or '',             # 广告组id
            r.get('广告组名称', '') or '',             # 广告组名称
            r.get('广告编号', '') or '',               # 广告id
            r.get('广告名称', '') or '',               # 广告名称
            0,                                         # 广告素材id
            0,                                         # 广告素材名称
            0,                                         # 笔记ID
            to_int(r.get('展示次数', 0)),              # 曝光
            to_int(r.get('点击量（全部）', 0)),        # 点击
            to_int(r.get('覆盖人数', 0)),              # 覆盖人数
            to_int(r.get('链接点击量', 0)),            # 链接点击量
            to_int(r.get('购物次数', 0)),              # 购物
            to_int(r.get('潜在客户人数', 0)),          # 潜在客户数
            to_int(r.get('消息对话发起次数', 0)),      # 消息发起次数
            '',                                        # 消耗类型
            round(cost, 6),                            # 消耗
            '',                                        # 行动按钮点击量
            to_int(r.get('视频播放进度达 25% 的次数', 0)),
            to_int(r.get('视频播放进度达 50% 的次数', 0)),
            to_int(r.get('视频播放进度达 75% 的次数', 0)),
            to_int(r.get('视频播放进度达 95% 的次数', 0)),
            to_int(r.get('视频播放进度达 100% 的次数', 0)),
            start_date,                                # 开始投放日期
            upload_date,                               # 上传日期
        ])

    if unmapped_countries:
        print(f"\n警告：{len(unmapped_countries)} 个国家代码未匹配中文：{sorted(unmapped_countries)}")

    # 输出：表头第4行，数据第5行起
    print(f"\n生成输出文件...")
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'

    # 第1行：填报说明
    ws.cell(row=1, column=1, value=(
        '填报说明：\n'
        '1、导入的数据从第5行开始，将需要导入的数据复制到本模板第5行开始即可\n'
        '2、日期、主投学科、平台、区域细分、投放账户、广告计划id、广告计划名称、广告组id、'
        '广告组名称、广告id、广告名称（ps： 这几个字段都不能为空，否则会导入失败）。\n'
        '3、若无广告计划id、广告组id、广告id、广告素材id、笔记id此类字段时，填入"0"'
    ))

    header = [
        '日期', '主投学科', '平台', '区域细分', '投放账户',
        '广告计划id', '广告计划名称', '广告组id', '广告组名称',
        '广告id', '广告名称', '广告素材id', '广告素材名称', '笔记ID',
        '曝光', '点击', '覆盖人数', '链接点击量', '购物', '潜在客户数', '消息发起次数',
        '消耗类型', '消耗', '行动按钮点击量',
        '视频播放进度达25%的次数', '视频播放进度达50%的次数',
        '视频播放进度达75%的次数', '视频播放进度达95%的次数', '视频播放进度达100%的次数',
        '开始投放日期', '上传日期',
    ]
    for col, h in enumerate(header, 1):
        ws.cell(row=4, column=col, value=h)

    for ri, row in enumerate(out_rows, start=5):
        for ci, val in enumerate(row, 1):
            ws.cell(row=ri, column=ci, value=val)

    try:
        wb.save(OUTPUT_FILE)
    except PermissionError:
        print(f"\n错误：无法写入 {OUTPUT_FILE}（被占用）")
        sys.exit(4)

    print("\n" + "=" * 50)
    print(f"完成！输出文件：{OUTPUT_FILE}")
    print(f"  输出行数：{len(out_rows)}")
    print(f"  总消耗(CNY)：{total_cost:,.2f}")
    print("=" * 50)


if __name__ == '__main__':
    main()
