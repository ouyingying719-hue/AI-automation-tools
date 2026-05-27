#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通过 smartbi-data-cli 工具上传两个 BI 回写表（HTTP API 方案）

依赖：
- 同事的 smartbi-data-cli 工具：%SMARTBI_CLI_ROOT%
- 本目录下 cli_branch/smartbi_writeback_tasks.json 配置

环境变量（自动从 .env 读取）：
- BI_USERNAME → SMARTBI_USERNAME
- BI_PASSWORD → SMARTBI_PASSWORD
"""

import os
import sys
import subprocess
from pathlib import Path

# Windows 控制台 GBK 容错：遇到无法编码字符替换为 ? 而不是抛 UnicodeEncodeError
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(errors='replace')
    sys.stderr.reconfigure(errors='replace')

SCRIPT_DIR = Path(__file__).resolve().parent
DATA_DIR = SCRIPT_DIR.parent  # 父目录是数据目录（.env、xlsx 都在那）
os.chdir(DATA_DIR)

CLI_ROOT = Path(Path(os.environ.get("SMARTBI_CLI_ROOT", r"D:\smartbi-data-cli")))
CLI_SCRIPT = CLI_ROOT / "scripts" / "smartbi_cli.py"
CONFIG_FILE = SCRIPT_DIR / "smartbi_writeback_tasks.json"

# 任务配置（仅任务一，任务二走 Selenium 方案）
TASKS = [
    {
        'name': '任务一-课包城市消耗',
        'task_id': 'course_city_cost_writeback',
        'file': '课包城市消耗数据回写_输出.xlsx',
    },
]


def load_env(key, default=''):
    env_path = DATA_DIR / '.env'
    if not env_path.exists():
        return default
    with open(env_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line.startswith(f'{key}='):
                return line.split('=', 1)[1].strip()
    return default


def run_cli(args, env_extra=None):
    """运行 smartbi_cli.py"""
    cmd = [sys.executable, str(CLI_SCRIPT)] + args
    env = os.environ.copy()
    if env_extra:
        env.update(env_extra)
    result = subprocess.run(
        cmd,
        cwd=str(CLI_ROOT),
        capture_output=True,
        encoding='utf-8',
        errors='replace',
        env=env,
    )
    return result


def writeback_check(task_id, file_path):
    """上传前本地校验"""
    print(f"\n[校验] {task_id}")
    print(f"  文件: {file_path}")
    result = run_cli([
        'writeback-check',
        '--config', str(CONFIG_FILE),
        '--task', task_id,
        '--file', str(file_path),
        '--json',
    ])
    print(result.stdout[:2000] if result.stdout else '')
    if result.returncode != 0:
        print(f"  [失败] 返回码 {result.returncode}")
        if result.stderr:
            print(f"  stderr: {result.stderr[:1000]}")
        return False
    return True


def writeback_upload(task_id, file_path, username, password, operator='auto'):
    """实际上传到 BI"""
    print(f"\n[上传] {task_id}")
    print(f"  文件: {file_path}")
    result = run_cli([
        'writeback-upload',
        '--config', str(CONFIG_FILE),
        '--task', task_id,
        '--file', str(file_path),
        '--operator', operator,
        '--confirm',
        '--json',
    ], env_extra={
        'SMARTBI_USERNAME': username,
        'SMARTBI_PASSWORD': password,
    })
    print(result.stdout[:3000] if result.stdout else '')

    # 兜底判定：CLI 工具有时因 post_verify 失败而返回非0或 status=error，
    # 但 BI 实际已经成功入库（callback 显示 success=true / successCount=N）。
    # 优先看 BI 实际回调，回调成功就算上传成功。
    upload_succeeded = False
    callback_success_count = None
    callback_total_count = None
    try:
        import json
        data = json.loads(result.stdout)
        callback = (data.get('callback') or {}).get('callback_json') or {}
        if callback.get('success') is True:
            sheets = callback.get('sheetResults') or []
            if sheets and all(s.get('success') for s in sheets):
                callback_success_count = sum(s.get('successCount', 0) for s in sheets)
                callback_total_count = sum(s.get('totalCount', 0) for s in sheets)
                upload_succeeded = True
    except Exception:
        pass

    if upload_succeeded:
        print(f"  [成功] BI 回调 success=true, 入库 {callback_success_count}/{callback_total_count} 行")
        if data.get('status') != 'ok':
            print(f"  (CLI status={data.get('status')}，可能是 post_verify 虚警，已忽略)")
        return True

    if result.returncode != 0:
        print(f"  [失败] 返回码 {result.returncode}")
        if result.stderr:
            print(f"  stderr: {result.stderr[:1000]}")
        return False
    try:
        import json
        data = json.loads(result.stdout)
        if data.get('status') != 'ok':
            print(f"  [失败] status={data.get('status')}, error={data.get('error')}")
            return False
    except Exception:
        pass
    return True


def main():
    print("=" * 60)
    print("CLI 方案：通过 HTTP API 上传 BI")
    print("=" * 60)

    # 检查 CLI 工具
    if not CLI_SCRIPT.exists():
        print(f"[失败] 找不到 CLI 工具: {CLI_SCRIPT}")
        sys.exit(1)
    if not CONFIG_FILE.exists():
        print(f"[失败] 找不到配置: {CONFIG_FILE}")
        sys.exit(1)

    # 凭据
    username = load_env('BI_USERNAME')
    password = load_env('BI_PASSWORD')
    if not username or not password:
        print("[失败] .env 里没配置 BI_USERNAME / BI_PASSWORD")
        sys.exit(1)

    # 检查文件
    for t in TASKS:
        if not (DATA_DIR / t['file']).exists():
            print(f"[失败] 找不到 {t['file']}")
            sys.exit(1)

    results = {}
    for t in TASKS:
        file_path = DATA_DIR / t['file']

        # 步骤1：本地校验
        if not writeback_check(t['task_id'], file_path):
            results[t['name']] = False
            continue

        # 步骤2：实际上传
        ok = writeback_upload(t['task_id'], file_path, username, password)
        results[t['name']] = ok

    # 汇总
    print("\n" + "=" * 60)
    print("上传结果汇总：")
    all_ok = True
    for t in TASKS:
        ok = results.get(t['name'], False)
        status = '✓ 成功' if ok else '✗ 失败'
        print(f"  {t['name']}: {status}")
        if not ok:
            all_ok = False
    print("=" * 60)

    sys.exit(0 if all_ok else 1)


if __name__ == '__main__':
    main()
