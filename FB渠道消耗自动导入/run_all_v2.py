#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
全自动每日任务（每天 8:00 自动跑）

流程：
  任务一：下载消耗.py → 转换脚本.py → 上传BI.py
        生成「课包城市消耗数据回写_输出.xlsx」并导入 BI

  任务二：下载KOL消耗.py → 转换KOL格式.py → 上传BI_任务2.py
        生成「投放账户数据回写填报模板_输出.xlsx」并导入 BI

跑完无论成败都发钉钉通知。

前置条件：
  - 主浏览器已开（开机自启 + 启动主浏览器.bat），调试端口 9222
  - 浏览器里两个上传页 tab 已固定打开：
      1. 海外投放课包城市维度数据填报 → 课包城市消耗数据回写
      2. 海外投放数据产品维度数据填报 → 投放账户产数据回写填报模板
"""

import os
import sys
import time
import subprocess
from datetime import datetime, timedelta

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)

LOG_FILE = os.path.join(SCRIPT_DIR, 'run_log.txt')


def log(msg):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line)
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except Exception:
        pass


def log_raw(msg):
    print(msg)
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(msg + '\n')
    except Exception:
        pass


def run_step(title, script_name):
    log("=" * 50)
    log(f" {title}")
    log("=" * 50)
    result = subprocess.run(
        [sys.executable, script_name],
        capture_output=True,
        encoding='utf-8',
        errors='replace',
    )
    output = (result.stdout or '') + (result.stderr or '')
    log_raw(output)
    if result.returncode != 0:
        log(f"[失败] {script_name} 返回码：{result.returncode}")
        return False, output
    return True, output


def extract_summary(download_output, convert_output, prefix=''):
    summary = {}
    for line in download_output.split('\n'):
        if '总计：' in line and '行' in line:
            summary[f'{prefix}下载行数'] = line.strip()
        if '失败' in line and '账户' in line:
            summary[f'{prefix}失败账户'] = line.strip()
    for line in convert_output.split('\n'):
        if '总曝光' in line:
            summary[f'{prefix}总曝光'] = line.strip()
        if '总点击' in line:
            summary[f'{prefix}总点击'] = line.strip()
        if '总消耗' in line:
            summary[f'{prefix}总消耗'] = line.strip()
        if '输出行数' in line:
            summary[f'{prefix}输出行数'] = line.strip()
    return summary


def extract_error_tail(output, max_lines=10):
    lines = [l for l in output.split('\n') if l.strip()]
    return '\n'.join(lines[-max_lines:])


def detect_friendly_error(output):
    if 'PermissionError' in output and ('Permission denied' in output or '被占用' in output):
        return "文件被 Excel 打开，无法写入。请关闭 Excel 后重新运行。"
    if 'getaddrinfo failed' in output or 'Name or service not known' in output:
        return "DNS 解析失败，可能是网络/VPN 问题。"
    if 'HTTP 401' in output or 'expired' in output.lower() or 'invalid_token' in output.lower():
        return "FB token 失效，需要更新 .env 里的 FB_ACCESS_TOKEN。"
    if '主浏览器' in output and ('没开' in output or '没运行' in output):
        return "主浏览器没启动。请双击 启动主浏览器.bat 启动 Chrome 后重试。"
    if '没找到上传页' in output:
        return "浏览器里没有上传页 tab。请先手动点开对应的 BI 上传页并保留 tab。"
    return None


def send_notification(success, target_date, details):
    status = '成功' if success else '失败'
    try:
        subprocess.run(
            [sys.executable, '发送钉钉.py', status, details],
            capture_output=True,
            timeout=15,
        )
    except Exception as e:
        log(f"发钉钉异常：{e}")


def run_task(task_name, scripts, target_date):
    """跑一个任务（下载→转换→上传），返回 (是否成功, summary 字典, 错误尾部)"""
    log("\n" + "#" * 50)
    log(f"# {task_name}")
    log("#" * 50)

    download_output = ''
    convert_output = ''

    # 步骤1：下载
    ok, download_output = run_step(f"{task_name} - 步骤1：下载FB数据", scripts[0])
    if not ok:
        return False, {}, (f"{task_name} - 下载失败", download_output)

    # 步骤2：转换
    ok, convert_output = run_step(f"{task_name} - 步骤2：转换格式", scripts[1])
    if not ok:
        return False, {}, (f"{task_name} - 转换失败", convert_output)

    # 步骤3：上传
    ok, upload_output = run_step(f"{task_name} - 步骤3：上传BI", scripts[2])
    if not ok:
        return False, {}, (f"{task_name} - 上传失败", upload_output)

    summary = extract_summary(download_output, convert_output, prefix=f'[{task_name}] ')
    return True, summary, None


def main():
    # 重置日志
    try:
        with open(LOG_FILE, 'w', encoding='utf-8') as f:
            f.write(f"========== 运行开始 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ==========\n")
    except Exception:
        pass

    start_time = time.time()
    target_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
    log(f"目标日期：{target_date}")
    log(f"Python：{sys.executable}")
    log(f"工作目录：{os.getcwd()}")

    all_summary = {}
    failures = []

    # === 任务一：下载 + 转换 ===
    log("\n" + "#" * 50)
    log("# 任务一-课包城市消耗：下载+转换")
    log("#" * 50)
    ok1_download, dl1_output = run_step("任务一 - 步骤1：下载FB数据", "下载消耗.py")
    ok1_convert = False
    cv1_output = ''
    if ok1_download:
        ok1_convert, cv1_output = run_step("任务一 - 步骤2：转换格式", "转换脚本.py")
    if not ok1_download:
        failures.append(("任务一 - 下载失败", dl1_output))
    elif not ok1_convert:
        failures.append(("任务一 - 转换失败", cv1_output))
    all_summary.update(extract_summary(dl1_output, cv1_output, prefix='[任务一] '))

    # === 任务二：下载 + 转换 ===
    log("\n" + "#" * 50)
    log("# 任务二-投放账户数据回写：下载+转换")
    log("#" * 50)
    ok2_download, dl2_output = run_step("任务二 - 步骤1：下载FB数据", "下载KOL消耗.py")
    ok2_convert = False
    cv2_output = ''
    if ok2_download:
        ok2_convert, cv2_output = run_step("任务二 - 步骤2：转换格式", "转换KOL格式.py")
    if not ok2_download:
        failures.append(("任务二 - 下载失败", dl2_output))
    elif not ok2_convert:
        failures.append(("任务二 - 转换失败", cv2_output))
    all_summary.update(extract_summary(dl2_output, cv2_output, prefix='[任务二] '))

    # === 合并上传（一个 Selenium session 处理两个 tab）===
    ok1 = ok1_download and ok1_convert
    ok2 = ok2_download and ok2_convert
    upload_ok = False
    if ok1 or ok2:
        log("\n" + "#" * 50)
        log("# 合并上传 BI")
        log("#" * 50)
        upload_success, upload_output = run_step("合并上传 BI", "上传BI_合并.py")
        if not upload_success:
            failures.append(("合并上传失败", upload_output))
        upload_ok = upload_success
    else:
        log("\n两个任务的下载/转换都失败了，跳过上传")

    # 最终判定
    elapsed = int(time.time() - start_time)
    overall_success = ok1 and ok2 and upload_ok

    if overall_success:
        detail_lines = [
            f"日期：{target_date}",
            f"耗时：{elapsed}秒",
            "",
            "✅ 任务一-课包城市消耗：成功",
            "✅ 任务二-投放账户数据回写：成功",
        ]
        for key in [
            '[任务一] 下载行数',
            '[任务一] 输出行数',
            '[任务一] 总消耗',
            '[任务二] 下载行数',
            '[任务二] 输出行数',
            '[任务二] 总消耗',
        ]:
            if key in all_summary:
                detail_lines.append(all_summary[key])
        send_notification(True, target_date, '\n'.join(detail_lines))
        log("全部完成！通知已发送")
    else:
        msg_lines = [f"日期：{target_date}", f"耗时：{elapsed}秒", ""]
        msg_lines.append("✅ 任务一" if ok1 else "❌ 任务一-课包城市消耗：失败")
        msg_lines.append("✅ 任务二" if ok2 else "❌ 任务二-投放账户数据回写：失败")
        msg_lines.append("")
        for fail_title, fail_output in failures:
            friendly = detect_friendly_error(fail_output)
            msg_lines.append(f"--- {fail_title} ---")
            if friendly:
                msg_lines.append(f"💡 {friendly}")
            msg_lines.append(extract_error_tail(fail_output, max_lines=8))
            msg_lines.append("")
        send_notification(False, target_date, '\n'.join(msg_lines))
        log("有任务失败，通知已发送")

    sys.exit(0 if overall_success else 1)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        import traceback
        err = traceback.format_exc()
        log(f"主流程异常：{err}")
        target_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        send_notification(False, target_date,
            f"日期：{target_date}\n步骤：主流程异常\n\n{err[-500:]}")
        sys.exit(1)
