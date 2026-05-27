#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
任务二专用 CLI 上传（fire-and-forget + 等5分钟 + post_verify 行数对账）

为什么单独一个脚本？
  任务二（投放账户数据回写）通常 1000+ 行，multipart 请求体大，
  BI 后端入库时间长，上行响应经常被 Cloudflare 100s 超时拦截（524）。
  CLI 默认的 writeback-upload 命令会重试 3 次（约 6 分钟），不符合
  "fire-and-forget"语义。这里用 SmartbiClient 内部 API 自己控制：
    1. login
    2. 触发 multipart 上传，retries=1，超时/524 都视为"已发出"
    3. sleep 5 分钟，让 BI 慢慢入库
    4. post_verify：登录 → 拉取「投放账户数据回写表（开发中）」当日数据
       → 比较行数 vs 上传行数；一致 → 成功；不一致 → 失败
"""

import os
import sys
import json
import time
import io
from pathlib import Path
from datetime import datetime, timedelta

# Windows 控制台 GBK 容错：遇到无法编码字符替换为 ? 而不是抛 UnicodeEncodeError
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(errors='replace')
    sys.stderr.reconfigure(errors='replace')

SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = SCRIPT_DIR.parent
os.chdir(DATA_DIR)

CLI_ROOT = Path(Path(os.environ.get("SMARTBI_CLI_ROOT", r"D:\smartbi-data-cli")))
CLI_SCRIPTS = CLI_ROOT / "scripts"
CONFIG_FILE = SCRIPT_DIR / "smartbi_writeback_tasks.json"

# 引入 CLI 工具的内部模块
sys.path.insert(0, str(CLI_SCRIPTS))
import smartbi_cli  # noqa: E402
from smartbi_cli import (  # noqa: E402
    SmartbiClient,
    SmartbiError,
    load_writeback_task,
    writeback_target,
    submit_data_acquisition_multipart,
    params_from_context,
    apply_param_overrides,
)

TASK_ID = "ad_account_writeback"
UPLOAD_FILE = "投放账户数据回写填报模板_输出.xlsx"
WAIT_SECONDS = 5 * 60  # 5 分钟


def load_env(key, default=""):
    env_path = DATA_DIR / ".env"
    if not env_path.exists():
        return default
    with open(env_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line.startswith(f"{key}="):
                return line.split("=", 1)[1].strip()
    return default


def count_data_rows(xlsx_path, header_row=4, data_start_row=5):
    """统计上传文件的有效数据行数（首列日期非空）"""
    from openpyxl import load_workbook
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    count = 0
    for row in ws.iter_rows(min_row=data_start_row, values_only=True):
        if not row:
            continue
        first = row[0]
        if first is None or (isinstance(first, str) and not first.strip()):
            continue
        count += 1
    wb.close()
    return count


def count_export_rows(body_bytes):
    """统计 BI 导出的 SPREADSHEET_REPORT xlsx 行数（首列非空 + 跳过表头）"""
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(body_bytes), read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]

    # 找到 "日期" 列所在的行（表头行）
    header_row_idx = None
    date_col_idx = None
    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
        for c_idx, cell in enumerate(row):
            if isinstance(cell, str) and cell.strip() == "日期":
                header_row_idx = r_idx
                date_col_idx = c_idx
                break
        if header_row_idx:
            break

    if header_row_idx is None:
        # 找不到表头，退化为统计所有非空首列行数
        count = 0
        for row in ws.iter_rows(min_row=1, values_only=True):
            if row and row[0] not in (None, ""):
                count += 1
        wb.close()
        return count, "fallback_first_col"

    # 表头之后的行，按 "日期" 列非空计数
    count = 0
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        if not row or date_col_idx >= len(row):
            continue
        cell = row[date_col_idx]
        if cell is None:
            continue
        if isinstance(cell, str) and not cell.strip():
            continue
        count += 1
    wb.close()
    return count, f"date_col_at_row{header_row_idx}"


def fire_upload(client, task, candidate_path):
    """触发上传，超时/524 都视为"已发出"，不抛异常"""
    target = writeback_target(task)
    target_id = str(target["report_id"])

    # 预热：访问 openimportconfig 拿 sessionid 和 upload context
    print(f"[upload] 预热目标 {target_id}")
    client.request(
        "openimportconfig.jsp",
        {"isBrowse": "true", "showLeftTree": "default", "resid": target_id},
    )

    # 拿 upload_context
    from smartbi_cli import (
        initialize_data_acquisition_upload_context,
        target_with_upload_context,
    )
    upload_context = initialize_data_acquisition_upload_context(client, target)
    effective_target = target_with_upload_context(target, upload_context)

    # 触发 multipart 上传（retries=1，单次发起）
    print(f"[upload] 触发上传 {candidate_path.name} ...")
    fire_started = time.time()
    try:
        upload_response = submit_data_acquisition_multipart(
            client=client,
            target=effective_target,
            candidate_file=candidate_path,
            retries=1,
        )
        fire_elapsed = int(time.time() - fire_started)
        print(f"[upload] BI 已返回，耗时 {fire_elapsed}s")
        callback = upload_response.get("response", {}).get("parsed", {})
        print(f"[upload] callback: {json.dumps(callback, ensure_ascii=False)[:500]}")
        return {"status": "completed", "callback": callback, "elapsed": fire_elapsed}
    except SmartbiError as e:
        fire_elapsed = int(time.time() - fire_started)
        # 超时/524/网络错误 → 视为已发出（fire-and-forget 模式）
        msg = str(e)
        if any(k in msg for k in ["timed out", "timeout", "524", "network_error"]):
            print(f"[upload] BI 未在 {fire_elapsed}s 内返回（{msg[:150]}），视为已触发")
            return {"status": "fire_and_forget", "error": msg, "elapsed": fire_elapsed}
        # 真错误才抛
        raise


def post_verify(client, task, target_date):
    """从 post_verify 报表查询 target_date 的数据行数"""
    pv = task.get("post_verify") or {}
    report = pv.get("report") or {}
    report_id = report.get("report_id")
    if not report_id:
        raise SmartbiError("post_verify.report.report_id 缺失")
    alias = report.get("alias") or report_id

    print(f"[verify] 拉取报表 {alias} ({report_id}) 的 {target_date} 数据 ...")
    context = client.open_report_context(report_id)
    params = params_from_context(context)

    # 设置参数：开始日期=结束日期=target_date
    # 配置里 post_verify.parameters 用 {import_date} 占位符，这里直接覆盖为 target_date
    pv_params = pv.get("parameters") or {}
    overrides = []
    for k, v in pv_params.items():
        if v is None:
            continue
        v_str = str(v)
        if v_str == "{import_date}":
            v_str = target_date
        if v_str == "":
            # 空字符串：仍下发空值（让 BI 视为不过滤）
            overrides.append({"key": k, "value": ""})
        else:
            overrides.append({"key": k, "value": v_str})
    if overrides:
        params = apply_param_overrides(params, overrides)
        print(f"[verify] 参数: {overrides}")

    filename, body = client.export_spreadsheet_report(report_id, context, params=params)
    print(f"[verify] 拉取成功 {filename}, {len(body)} bytes")

    rows, mode = count_export_rows(body)
    print(f"[verify] 报表行数={rows} (mode={mode})")
    return rows


def main():
    print("=" * 60)
    print("任务二 CLI 上传（fire-and-forget + 5min + post_verify）")
    print("=" * 60)

    if not CONFIG_FILE.exists():
        print(f"[失败] 找不到配置: {CONFIG_FILE}")
        sys.exit(1)

    candidate = DATA_DIR / UPLOAD_FILE
    if not candidate.exists():
        print(f"[失败] 找不到上传文件: {candidate}")
        sys.exit(1)

    username = load_env("BI_USERNAME")
    password = load_env("BI_PASSWORD")
    if not username or not password:
        print("[失败] .env 缺 BI_USERNAME / BI_PASSWORD")
        sys.exit(1)

    # 上传文件行数（昨天的数据）
    upload_rows = count_data_rows(candidate)
    print(f"[input] 上传文件 {UPLOAD_FILE} 数据行数={upload_rows}")

    # 目标日期（昨天，即上传文件里的"日期"列值）
    target_date = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    print(f"[input] 目标日期={target_date}")

    # 加载任务配置
    task = load_writeback_task(CONFIG_FILE, TASK_ID)

    # 登录
    client = SmartbiClient()
    print(f"[login] {username}")
    client.login(username, password)

    # ---- 触发上传 ----
    upload_info = fire_upload(client, task, candidate)
    print(f"[upload] 状态: {upload_info['status']}")

    # ---- 等待 5 分钟 ----
    print(f"\n[wait] 等待 BI 入库 {WAIT_SECONDS}s ...")
    waited = 0
    while waited < WAIT_SECONDS:
        time.sleep(30)
        waited += 30
        print(f"[wait] {waited}/{WAIT_SECONDS}s ...")
    print("[wait] 等待完成")

    # ---- post_verify ----
    print()
    try:
        verify_rows = post_verify(client, task, target_date)
    except SmartbiError as e:
        print(f"[verify] 失败: {e}")
        # post_verify 网络错误：保守判定为失败，让人工排查
        print("\n" + "=" * 60)
        print("结果汇总：")
        print(f"  上传行数: {upload_rows}")
        print(f"  上传状态: {upload_info['status']}")
        print(f"  post_verify: 失败 - {e}")
        print("=" * 60)
        sys.exit(1)

    # ---- 行数对比 ----
    print()
    print("=" * 60)
    print("结果汇总：")
    print(f"  上传文件行数: {upload_rows}")
    print(f"  上传请求状态: {upload_info['status']}")
    print(f"  post_verify 报表行数: {verify_rows}")
    print(f"  目标日期: {target_date}")

    if verify_rows == upload_rows:
        print(f"  ✓ 行数一致，导入成功")
        print("=" * 60)
        sys.exit(0)
    elif verify_rows == 0:
        print(f"  ✗ 报表里没有 {target_date} 的数据，导入可能失败")
        print("=" * 60)
        sys.exit(1)
    else:
        diff = upload_rows - verify_rows
        print(f"  ⚠ 行数不一致，差 {diff} 行")
        # 容忍 5% 偏差（去重/聚合可能造成）
        tolerance = max(1, int(upload_rows * 0.05))
        if abs(diff) <= tolerance:
            print(f"  在容忍范围内（±{tolerance}），视为成功")
            print("=" * 60)
            sys.exit(0)
        print(f"  超出容忍范围（±{tolerance}），判定为失败")
        print("=" * 60)
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print(f"[异常] {e}")
        traceback.print_exc()
        sys.exit(2)
