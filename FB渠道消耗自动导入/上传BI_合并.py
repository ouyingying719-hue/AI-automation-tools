#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
合并上传：在一个 Selenium session 里完成两个任务的 BI 上传

逻辑：
- 连接主浏览器（端口 9222）
- 找到所有含 file input 的 BI tab
- 对每个 tab，进入 iframe 后通过内容区分是哪个任务
- 分别上传对应的文件

前提：浏览器里两个上传页 tab 都打开着
"""

import os
import sys
import time
import urllib.request
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)

DEBUG_PORT = 9222

# 两个任务的配置
TASKS = [
    {
        'name': '任务一-课包城市消耗',
        'file': '课包城市消耗数据回写_输出.xlsx',
        'keyword': '课包城市',  # iframe 内容里的区分关键词
    },
    {
        'name': '任务二-投放账户数据回写',
        'file': '投放账户数据回写填报模板_输出.xlsx',
        'keyword': '投放账户',
    },
]


def check_browser():
    try:
        urllib.request.urlopen(f'http://localhost:{DEBUG_PORT}/json/version', timeout=3)
        return True
    except Exception:
        return False


def connect_browser():
    options = Options()
    options.add_experimental_option("debuggerAddress", f"localhost:{DEBUG_PORT}")
    return webdriver.Chrome(options=options)


def find_first_visible(driver, xpaths):
    for xp in xpaths:
        try:
            for e in driver.find_elements(By.XPATH, xp):
                try:
                    if e.is_displayed():
                        return e
                except Exception:
                    continue
        except Exception:
            continue
    return None


def find_all_upload_tabs(driver):
    """找到所有含 file input 的 BI tab，返回 [(handle, file_input, iframe_content_snippet)]"""
    results = []
    print("扫描所有 tab...")
    for handle in driver.window_handles:
        try:
            driver.switch_to.window(handle)
            try:
                driver.switch_to.default_content()
            except Exception:
                pass
            try:
                url = driver.current_url
                title = driver.title
            except Exception:
                continue
            if 'bi.61info.cn' not in url:
                continue
            print(f"  检查 tab: {title[:30]} | {url[:60]}")

            # 主文档找 file input
            try:
                inputs = driver.find_elements(By.XPATH, "//input[@type='file']")
                if inputs:
                    # 取主文档部分内容用于区分
                    snippet = driver.page_source[:5000]
                    results.append((handle, inputs[0], snippet, -1))
                    print(f"    有 file input（主文档）")
                    continue
            except Exception:
                pass

            # iframe 里找
            try:
                iframes = driver.find_elements(By.TAG_NAME, 'iframe')
                for idx, iframe in enumerate(iframes):
                    try:
                        driver.switch_to.frame(iframe)
                        inputs = driver.find_elements(By.XPATH, "//input[@type='file']")
                        if inputs:
                            snippet = driver.page_source[:5000]
                            results.append((handle, inputs[0], snippet, idx))
                            print(f"    有 file input（iframe {idx}）")
                            driver.switch_to.default_content()
                            break
                        driver.switch_to.default_content()
                    except Exception:
                        try:
                            driver.switch_to.default_content()
                        except Exception:
                            pass
            except Exception:
                pass
        except Exception:
            try:
                driver.switch_to.default_content()
            except Exception:
                pass
    return results


def switch_to_tab(driver, handle, iframe_idx):
    """切换到指定 tab 和 iframe"""
    driver.switch_to.window(handle)
    driver.switch_to.default_content()
    if iframe_idx >= 0:
        iframes = driver.find_elements(By.TAG_NAME, 'iframe')
        if iframe_idx < len(iframes):
            driver.switch_to.frame(iframes[iframe_idx])
    # 重新找 file input
    inputs = driver.find_elements(By.XPATH, "//input[@type='file']")
    return inputs[0] if inputs else None


def upload_file(driver, file_input, file_path, task_name):
    """选文件 → 点导入 → 处理确认弹框"""
    abs_path = os.path.abspath(file_path)
    print(f"\n[{task_name}] 上传文件：{abs_path}")
    file_input.send_keys(abs_path)
    time.sleep(1)

    import_btn = find_first_visible(driver, [
        "//input[@value='导入']",
        "//button[text()='导入']",
        "//a[text()='导入']",
        "//*[@onclick and contains(@onclick, 'import')]",
    ])
    if not import_btn:
        print(f"[{task_name}] [失败] 找不到导入按钮")
        return False

    try:
        driver.execute_script("""
            document.querySelectorAll('.layui-layer-shade, .layui-layer').forEach(e => e.remove());
        """)
    except Exception:
        pass

    print(f"[{task_name}] 点击导入...")
    try:
        import_btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", import_btn)

    print(f"[{task_name}] 等待确认弹框...")
    confirm_clicked = False
    for _ in range(15):
        time.sleep(1)
        for xpath in [
            "//button[contains(text(),'是')]",
            "//button[contains(text(),'确定')]",
            "//input[@value='是' or @value='确定' or starts-with(@value,'是') or starts-with(@value,'确定')]",
            "//a[contains(text(),'是') or contains(text(),'确定')]",
            "//*[contains(@class,'btn-primary') and (contains(text(),'是') or contains(text(),'确定'))]",
        ]:
            try:
                for btn in driver.find_elements(By.XPATH, xpath):
                    try:
                        if btn.is_displayed() and btn.is_enabled():
                            txt = btn.text or btn.get_attribute('value') or ''
                            print(f"[{task_name}] 找到确认按钮：{txt!r}，点击")
                            btn.click()
                            confirm_clicked = True
                            break
                    except Exception:
                        continue
                if confirm_clicked:
                    break
            except Exception:
                continue
        if confirm_clicked:
            break

    if not confirm_clicked:
        print(f"[{task_name}] [警告] 没看到确认弹框")

    print(f"[{task_name}] 等待导入完成...")
    time.sleep(8)

    try:
        # 切回主文档检查结果（弹框可能在主文档层）
        driver.switch_to.default_content()
        page_source = driver.page_source
    except Exception:
        page_source = ''

    fail_keywords = ['导入失败', '上传失败', '错误信息', 'BindingParam']
    success_keywords = ['导入成功', '上传成功']

    for kw in fail_keywords:
        if kw in page_source:
            print(f"[{task_name}] [失败] 页面包含「{kw}」")
            return False

    for kw in success_keywords:
        if kw in page_source:
            print(f"[{task_name}] [成功] 页面包含「{kw}」")
            return True

    if confirm_clicked:
        print(f"[{task_name}] [假定成功] 已点确认按钮，未见错误提示")
        return True

    print(f"[{task_name}] [失败] 没点到确认按钮，也没看到成功提示")
    return False


def main():
    print("=" * 50)
    print("合并上传：任务一 + 任务二")
    print("=" * 50)

    # 检查文件
    for task in TASKS:
        if not os.path.exists(task['file']):
            print(f"错误：找不到 {task['file']}")
            sys.exit(1)

    if not check_browser():
        print(f"[失败] 主浏览器调试端口 {DEBUG_PORT} 没开")
        sys.exit(1)

    driver = connect_browser()
    results = {}

    try:
        # 找到所有上传页 tab
        tabs = find_all_upload_tabs(driver)
        print(f"\n找到 {len(tabs)} 个含 file input 的 tab")

        if len(tabs) < 2:
            print("[失败] 需要 2 个上传页 tab（任务一 + 任务二），请在浏览器里都打开")
            sys.exit(1)

        # 匹配：通过 snippet 里的关键词区分
        # 如果 snippet 区分不了（SPA 导航树都有），就按顺序分配
        task_tab_map = {}  # task_name -> (handle, iframe_idx)

        for task in TASKS:
            for i, (handle, fi, snippet, iframe_idx) in enumerate(tabs):
                if handle in [v[0] for v in task_tab_map.values()]:
                    continue  # 已分配给其他任务
                # 尝试关键词匹配
                if task['keyword'] in snippet:
                    task_tab_map[task['name']] = (handle, iframe_idx)
                    break

        # 如果关键词匹配不了（都含或都不含），按顺序分配
        if len(task_tab_map) < 2:
            print("  关键词无法区分，按 tab 顺序分配（第1个=任务一，第2个=任务二）")
            task_tab_map = {}
            for i, task in enumerate(TASKS):
                if i < len(tabs):
                    task_tab_map[task['name']] = (tabs[i][0], tabs[i][3])

        print(f"\n分配结果：")
        for name, (handle, idx) in task_tab_map.items():
            print(f"  {name} → tab handle={handle[:20]}..., iframe={idx}")

        # 逐个上传
        for task in TASKS:
            if task['name'] not in task_tab_map:
                print(f"\n[{task['name']}] 没有分配到 tab，跳过")
                results[task['name']] = False
                continue

            handle, iframe_idx = task_tab_map[task['name']]
            file_input = switch_to_tab(driver, handle, iframe_idx)
            if not file_input:
                print(f"\n[{task['name']}] 切换 tab 后找不到 file input")
                results[task['name']] = False
                continue

            ok = upload_file(driver, file_input, task['file'], task['name'])
            results[task['name']] = ok

            # 上传完等一下再处理下一个
            time.sleep(3)

    except Exception as e:
        print(f"\n错误：{type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    # 汇总
    print("\n" + "=" * 50)
    print("上传结果汇总：")
    all_ok = True
    for task in TASKS:
        ok = results.get(task['name'], False)
        status = '✓ 成功' if ok else '✗ 失败'
        print(f"  {task['name']}: {status}")
        if not ok:
            all_ok = False
    print("=" * 50)

    sys.exit(0 if all_ok else 1)


if __name__ == '__main__':
    main()
