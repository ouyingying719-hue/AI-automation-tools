#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FB广告预览链接自动记录脚本

逻辑:
1. 读取 素材信息.xlsx 中的"名称"列, 已有"预览链接"的行跳过
2. 维护 tracker.json 记录每个素材首次被发现的时间
3. 对首次发现时间已满 delay_hours (默认24小时) 且仍无预览链接的素材,
   按 routing_rules 关键词路由到优先账户; 调用 FB Marketing API 按广告名精确匹配
4. 优先账户找不到时降级到其他账户; 仍找不到则下次再试

依赖: openpyxl, requests
"""

import argparse
import json
import logging
import os
import sys
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path

import requests
from openpyxl import load_workbook

SCRIPT_DIR = Path(__file__).resolve().parent


def load_config(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def setup_logging(log_path):
    for h in list(logging.getLogger().handlers):
        logging.getLogger().removeHandler(h)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def load_tracker(path):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_tracker(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _request_with_retry(url, params, proxies, timeout, max_retries=4):
    """带指数退避重试的GET, 处理ReadTimeout/ConnectionError/5xx"""
    import time
    for attempt in range(max_retries):
        try:
            resp = requests.get(url, params=params, timeout=timeout, proxies=proxies)
            if resp.status_code >= 500:
                raise requests.exceptions.HTTPError(f"server {resp.status_code}")
            return resp
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError,
                requests.exceptions.HTTPError) as e:
            wait = 2 ** attempt
            logging.warning(f"  请求失败({type(e).__name__}), {wait}s后重试 {attempt+1}/{max_retries}")
            time.sleep(wait)
    return requests.get(url, params=params, timeout=timeout, proxies=proxies)


def fetch_ads_from_account(access_token, ad_account_id, api_version, page_size,
                           proxies=None, timeout=120, since_ts=None):
    """从一个广告账户分页拉取广告, 返回 {广告名称: 预览链接}
    since_ts: ISO时间戳字符串, 仅拉取此时间之后更新过的广告; None则全量"""
    import json as _json
    base = f"https://graph.facebook.com/{api_version}/{ad_account_id}/ads"
    params = {
        "access_token": access_token,
        "fields": "id,name,preview_shareable_link,creative{effective_object_story_id,object_story_id}",
        "limit": page_size,
    }
    if since_ts:
        params["filtering"] = _json.dumps([{
            "field": "ad.updated_time", "operator": "GREATER_THAN", "value": since_ts
        }])
        logging.info(f"    增量模式: 仅拉取 {since_ts} 之后更新的广告")
    result = {}
    total_seen = 0
    url = base
    page = 0
    stat = {"shareable": 0, "story": 0, "library": 0, "noname": 0}
    while url:
        page += 1
        resp = _request_with_retry(url, params, proxies, timeout)
        if resp.status_code != 200:
            logging.error(f"API请求失败 {resp.status_code}: {resp.text[:500]}")
            break
        data = resp.json()
        batch = data.get("data", [])
        total_seen += len(batch)
        for item in batch:
            name = item.get("name")
            if not name:
                stat["noname"] += 1
                continue
            link = None
            # 1) 优先取「获得评论的Facebook帖子」公开链接
            creative = item.get("creative") or {}
            story_id = creative.get("effective_object_story_id") or creative.get("object_story_id")
            if story_id and "_" in story_id:
                page_id, post_id = story_id.split("_", 1)
                link = f"https://www.facebook.com/{page_id}/posts/{post_id}"
                stat["story"] += 1
            # 2) 退而求其次:广告管理工具的分享链接(需登录有权限的账号才能查看)
            elif item.get("preview_shareable_link"):
                link = item["preview_shareable_link"]
                stat["shareable"] += 1
            # 3) 最后兜底:广告库公开链接
            elif item.get("id"):
                link = f"https://www.facebook.com/ads/library/?id={item['id']}"
                stat["library"] += 1
            if link:
                result[name] = link
        if page % 10 == 0:
            logging.info(f"    已拉取 {page} 页 / FB广告 {total_seen} 条 / 已记录 {len(result)} 条")
        url = (data.get("paging") or {}).get("next")
        params = None
    logging.info(f"    链接来源: 公开帖子{stat['story']} 分享预览{stat['shareable']} 广告库兜底{stat['library']} 无名{stat['noname']}")
    return result


def route_priority_accounts(name, rules):
    """按 routing_rules 命中关键词,返回该素材的优先账户列表(去重保序)"""
    priority = []
    for rule in rules or []:
        if any(kw in name for kw in rule.get("keywords", [])):
            for acct in rule.get("accounts", []):
                if acct not in priority:
                    priority.append(acct)
    return priority


def lookup_link(name, priority_accts, all_accts, ad_map_by_acct):
    """先在优先账户找,再在其他账户找。返回(链接, 命中账户) 或 (None, None)"""
    seen = set()
    for acct in priority_accts:
        seen.add(acct)
        link = ad_map_by_acct.get(acct, {}).get(name)
        if link:
            return link, acct
    for acct in all_accts:
        if acct in seen:
            continue
        link = ad_map_by_acct.get(acct, {}).get(name)
        if link:
            return link, acct
    return None, None


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="不调用FB API, 仅验证表格与tracker")
    parser.add_argument("--full-refresh", action="store_true", help="忽略广告缓存, 重新全量拉取")
    parser.add_argument("--offline", action="store_true", help="只用现有 ad_cache.json 匹配, 不调FB API")
    parser.add_argument("--ignore-delay", action="store_true", help="忽略 delay_hours, 立即匹配所有空白行")
    parser.add_argument("--config", default=str(SCRIPT_DIR / "config.json"))
    args = parser.parse_args()

    config = load_config(args.config)
    log_path = SCRIPT_DIR / config.get("log_file", "run.log")
    setup_logging(log_path)

    excel_path = SCRIPT_DIR / config["excel_path"]
    tracker_path = SCRIPT_DIR / config["tracker_file"]

    logging.info(f"=== 任务开始 {datetime.now().isoformat()} ===")
    if not excel_path.exists():
        logging.error(f"Excel文件不存在: {excel_path}")
        return 1

    tracker = load_tracker(tracker_path)
    wb = load_workbook(excel_path)
    sheet_name = config.get("sheet_name") or wb.sheetnames[0]
    ws = wb[sheet_name]

    headers = [cell.value for cell in ws[1]]
    name_col = config["name_column"]
    preview_col = config["preview_column"]
    if name_col not in headers:
        logging.error(f"未找到列'{name_col}', 现有表头: {headers}")
        return 1
    name_idx = headers.index(name_col) + 1
    if preview_col in headers:
        preview_idx = headers.index(preview_col) + 1
    else:
        preview_idx = ws.max_column + 1
        ws.cell(row=1, column=preview_idx, value=preview_col)

    now = datetime.now()
    delay = timedelta(hours=config.get("delay_hours", 24))
    new_seen, pending = 0, []

    for row in range(2, ws.max_row + 1):
        raw = ws.cell(row=row, column=name_idx).value
        if not raw or not str(raw).strip():
            continue
        name = str(raw).strip()
        existing = ws.cell(row=row, column=preview_idx).value
        if existing and str(existing).strip():
            continue
        if name not in tracker:
            tracker[name] = {"first_seen": now.isoformat()}
            new_seen += 1
            continue
        first_seen = datetime.fromisoformat(tracker[name]["first_seen"])
        if args.ignore_delay or now - first_seen >= delay:
            pending.append((row, name))

    logging.info(f"新发现素材: {new_seen} 个")
    logging.info(f"已满{delay.total_seconds()/3600:.0f}h待查询: {len(pending)} 个")

    rules = config.get("routing_rules", [])

    if args.dry_run and pending:
        bucket = Counter()
        for _, name in pending:
            pri = route_priority_accounts(name, rules)
            key = ",".join(pri) if pri else "(无规则命中)"
            bucket[key] += 1
        logging.info("--- 路由命中分布(dry-run) ---")
        for key, n in bucket.most_common():
            logging.info(f"  {n:5d} → {key}")

    if pending and not args.dry_run:
        # 加载 ad_cache: {acct: {"last_sync": iso, "ads": {name: link}}}
        cache_path = SCRIPT_DIR / config.get("ad_cache_file", "ad_cache.json")
        if cache_path.exists() and not args.full_refresh:
            with open(cache_path, "r", encoding="utf-8") as f:
                ad_cache = json.load(f)
            logging.info(f"已加载广告缓存: {cache_path.name}")
        else:
            ad_cache = {}
            if args.full_refresh:
                logging.info("--full-refresh: 忽略缓存, 全量重拉")

        # --offline: 跳过API调用,直接用缓存匹配
        if args.offline:
            logging.info("--offline: 仅使用现有缓存匹配, 不调用FB API")
            if not ad_cache:
                logging.error("ad_cache.json 不存在或为空, 无法离线匹配。请先正常跑一次拉取数据")
                return 1
            ad_map_by_acct = {a: e.get("ads", {}) for a, e in ad_cache.items()}
        else:
            proxy_url = config.get("proxy") or os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy")
            proxies = {"http": proxy_url, "https": proxy_url} if proxy_url else None
            if proxies:
                logging.info(f"使用代理: {proxy_url}")

            ad_map_by_acct = {}
            all_accts = config.get("ad_account_ids", [])
            sync_buffer_min = config.get("incremental_buffer_minutes", 30)
            for acct in all_accts:
                entry = ad_cache.get(acct, {})
                since_ts = entry.get("last_sync")
                mode = "增量" if since_ts else "全量(首次)"
                logging.info(f"拉取广告账户 {acct} ({mode})")
                this_sync = (datetime.now() - timedelta(minutes=sync_buffer_min)).isoformat(timespec="seconds")
                try:
                    m_new = fetch_ads_from_account(
                        config["access_token"], acct,
                        config.get("api_version", "v19.0"),
                        config.get("page_size", 50),
                        proxies=proxies,
                        timeout=config.get("request_timeout", 120),
                        since_ts=since_ts,
                    )
                    merged = dict(entry.get("ads") or {})
                    merged.update(m_new)
                    ad_cache[acct] = {"last_sync": this_sync, "ads": merged}
                    ad_map_by_acct[acct] = merged
                    logging.info(f"账户 {acct}: 本次新增/更新 {len(m_new)} 条, 缓存累计 {len(merged)} 条")
                except Exception as e:
                    logging.exception(f"账户 {acct} 拉取失败: {e}")
                    ad_map_by_acct[acct] = dict(entry.get("ads") or {})

            with open(cache_path, "w", encoding="utf-8") as f:
                json.dump(ad_cache, f, ensure_ascii=False, indent=2)
            logging.info(f"广告缓存已保存: {cache_path.name}")

        all_accts = config.get("ad_account_ids", [])

        updated = 0
        for row, name in pending:
            priority = route_priority_accounts(name, rules)
            link, hit_acct = lookup_link(name, priority, all_accts, ad_map_by_acct)
            if link:
                ws.cell(row=row, column=preview_idx, value=link)
                tracker[name]["preview_link"] = link
                tracker[name]["filled_at"] = now.isoformat()
                tracker[name]["hit_account"] = hit_acct
                fallback = " (降级)" if priority and hit_acct not in priority else ""
                logging.info(f"OK  {hit_acct}{fallback}  {name}")
                updated += 1
            else:
                logging.warning(f"MISS  优先={priority or '无规则命中'}  {name}")
        logging.info(f"本次写入预览链接 {updated} 个")

    save_tracker(tracker_path, tracker)
    if not args.dry_run:
        wb.save(excel_path)
    logging.info(f"=== 任务结束 {datetime.now().isoformat()} ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
