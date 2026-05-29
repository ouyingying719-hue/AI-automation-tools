"""FB 主页定时发帖脚本

读取本地 Excel 表格里的文案 + 素材文件名 + 定时时间，把每条帖子排进 FB 的
scheduled_publish_time 队列，发完后机器关掉也没事。

支持帖子类型：纯文字 / 单图 / 多图（最多 10 张）/ 单视频。
时间分配：表里"定时时间"列填了就以表为准；空着或没列就按 daily_slots 自动排。

视频走 FB 分块上传（resumable upload），公司网络抖动也能传。

用法：
    python fb_scheduler.py --dry-run                    # 只打印计划，不真发
    python fb_scheduler.py                              # 正式跑
    python fb_scheduler.py --only-rows 12,16,17,18,19   # 只重发指定行
"""
from __future__ import annotations

import argparse
import csv
import json
import logging
import os
import sys
import time
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Optional

import pytz
import requests
from openpyxl import load_workbook

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

GRAPH = "https://graph.facebook.com"


def setup_logger(log_path):
    logger = logging.getLogger("fb_scheduler")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(sh)
    return logger


@dataclass
class Config:
    access_token: str
    page: str
    api_version: str
    input_file: str
    sheet_name: str
    caption_col: str
    media_cols: list
    media_root: Path
    start_date_raw: str
    daily_slots: list
    timezone: str
    skip_weekends: bool
    video_exts: set
    image_exts: set
    log_file: str
    result_file: str
    proxy: Optional[str] = None
    schedule_col: Optional[str] = None

    @classmethod
    def load(cls, path):
        with open(path, encoding="utf-8") as f:
            raw = json.load(f)
        cols = raw["columns"]
        proxy = (raw.get("proxy")
                 or os.environ.get("HTTPS_PROXY")
                 or os.environ.get("https_proxy"))
        return cls(
            access_token=raw["access_token"].strip(),
            page=str(raw["page"]).strip(),
            api_version=raw.get("api_version", "v19.0"),
            input_file=raw["input_file"],
            sheet_name=raw.get("sheet_name", "Sheet1"),
            caption_col=cols["caption"],
            media_cols=list(cols["media"]),
            media_root=Path(raw["media_root"]),
            start_date_raw=raw["schedule"]["start_date"],
            daily_slots=list(raw["schedule"]["daily_slots"]),
            timezone=raw["schedule"].get("timezone", "Asia/Shanghai"),
            skip_weekends=bool(raw["schedule"].get("skip_weekends", False)),
            video_exts={e.lower() for e in raw["video_extensions"]},
            image_exts={e.lower() for e in raw["image_extensions"]},
            log_file=raw.get("log_file", "run.log"),
            result_file=raw.get("result_file", "result.csv"),
            proxy=proxy,
            schedule_col=cols.get("schedule"),
        )


class FB:
    def __init__(self, token, version, logger, proxy=None):
        self.token = token
        self.base = f"{GRAPH}/{version}"
        self.log = logger
        self.proxies = {"http": proxy, "https": proxy} if proxy else None
        if proxy:
            logger.info(f"使用代理: {proxy}")

    def _check(self, r):
        try:
            data = r.json()
        except Exception:
            r.raise_for_status()
            return {}
        if not r.ok or "error" in data:
            err = data.get("error", {})
            raise RuntimeError(
                f"FB API error {r.status_code}: "
                f"{err.get('message', data)} (type={err.get('type')}, code={err.get('code')})"
            )
        return data

    def get(self, path, **params):
        params.setdefault("access_token", self.token)
        r = requests.get(f"{self.base}/{path.lstrip('/')}", params=params,
                         timeout=60, proxies=self.proxies, verify=False)
        return self._check(r)

    def post(self, path, files=None, **data):
        data.setdefault("access_token", self.token)
        r = requests.post(f"{self.base}/{path.lstrip('/')}", data=data,
                          files=files, timeout=600, proxies=self.proxies, verify=False)
        return self._check(r)

    def resolve_page(self, page_ref):
        info = self.get(page_ref, fields="id,name,access_token")
        page_id = info["id"]
        page_token = info.get("access_token")
        if not page_token:
            me = self.get("me/accounts", fields="id,name,access_token", limit=200)
            for p in me.get("data", []):
                if p["id"] == page_id:
                    page_token = p.get("access_token")
                    break
        if not page_token:
            raise RuntimeError(
                f"拿不到 Page Access Token。请确认你是主页 {info.get('name')} 的管理员，"
                f"且 token 有 pages_manage_posts / pages_read_engagement 权限。"
            )
        return page_id, page_token


@dataclass
class MediaIndex:
    by_name: dict = field(default_factory=dict)
    duplicates: dict = field(default_factory=dict)

    @classmethod
    def build(cls, root, log):
        if not root.exists():
            raise RuntimeError(f"素材目录不存在或访问不到: {root}")
        idx = cls()
        count = 0
        for p in root.rglob("*"):
            if not p.is_file():
                continue
            key = p.name.lower()
            if key in idx.by_name:
                idx.duplicates.setdefault(key, [idx.by_name[key]]).append(p)
            else:
                idx.by_name[key] = p
                count += 1
        log.info(f"素材索引完成：{count} 个文件，重名 {len(idx.duplicates)} 个")
        for k, paths in idx.duplicates.items():
            log.warning(f"  重名: {k} -> {[str(x) for x in paths]}")
        return idx

    def find(self, name):
        return self.by_name.get(name.strip().lower())


@dataclass
class Post:
    row: int
    caption: str
    media_files: list
    media_kind: str
    scheduled_at: datetime
    note: str = ""


def parse_start_date(raw, tz):
    today = datetime.now(tz).date()
    if raw == "today":
        return today
    if raw == "tomorrow":
        return today + timedelta(days=1)
    return datetime.strptime(raw, "%Y-%m-%d").date()


def gen_slots(start, daily_slots, tz, skip_weekends, n):
    out = []
    d = start
    parsed_slots = [datetime.strptime(s, "%H:%M").time() for s in daily_slots]
    while len(out) < n:
        if skip_weekends and d.weekday() >= 5:
            d += timedelta(days=1)
            continue
        for t in parsed_slots:
            naive = datetime.combine(d, t)
            out.append(tz.localize(naive))
            if len(out) >= n:
                break
        d += timedelta(days=1)
    return out


_TIME_FORMATS = (
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y/%m/%d %H:%M:%S",
    "%Y/%m/%d %H:%M",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%dT%H:%M",
)


def parse_explicit_time(value, tz):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value if value.tzinfo else tz.localize(value)
    s = str(value).strip()
    if not s:
        return None
    for fmt in _TIME_FORMATS:
        try:
            naive = datetime.strptime(s, fmt)
            return tz.localize(naive)
        except ValueError:
            continue
    raise ValueError(f"无法解析定时时间 '{s}'，支持格式如 '2026-06-05 14:30'")


def read_posts(cfg, idx, tz, log):
    """返回 [(row, caption, media_paths, kind, note, explicit_dt), ...]"""
    wb = load_workbook(cfg.input_file, data_only=True)
    if cfg.sheet_name not in wb.sheetnames:
        raise RuntimeError(f"表里没有 sheet '{cfg.sheet_name}'，可选: {wb.sheetnames}")
    ws = wb[cfg.sheet_name]
    headers = [c.value for c in ws[1]]

    def col_idx(name):
        for i, h in enumerate(headers):
            if h is not None and str(h).strip() == name:
                return i
        return None

    cap_i = col_idx(cfg.caption_col)
    if cap_i is None:
        raise RuntimeError(f"表头里找不到文案列 '{cfg.caption_col}'。表头: {headers}")
    media_idx_list = [col_idx(c) for c in cfg.media_cols]
    sched_i = col_idx(cfg.schedule_col) if cfg.schedule_col else None
    if cfg.schedule_col and sched_i is None:
        log.warning(f"表头里没有定时时间列 '{cfg.schedule_col}'，将回退到 daily_slots 自动排期")

    rows = []
    for r_i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        caption = row[cap_i] if cap_i < len(row) else None
        caption = "" if caption is None else str(caption).strip()

        names = []
        for ci in media_idx_list:
            if ci is None:
                continue
            v = row[ci] if ci < len(row) else None
            if v is None:
                continue
            s = str(v).strip()
            if s:
                names.append(s)

        if not caption and not names:
            continue

        paths = []
        missing = []
        for n in names:
            p = idx.find(n)
            if p is None:
                missing.append(n)
            else:
                paths.append(p)

        if not paths:
            kind = "text"
        else:
            exts = {p.suffix.lower() for p in paths}
            if exts & cfg.video_exts:
                if len(paths) > 1:
                    log.warning(f"行 {r_i}: 视频帖只支持单个视频，将只用第一个")
                    paths = [next(p for p in paths if p.suffix.lower() in cfg.video_exts)]
                kind = "video"
            else:
                kind = "photo"

        note = ""
        if missing:
            note = f"未匹配素材: {missing}"
            log.warning(f"行 {r_i}: {note}")

        explicit_dt = None
        if sched_i is not None:
            sval = row[sched_i] if sched_i < len(row) else None
            try:
                explicit_dt = parse_explicit_time(sval, tz)
            except ValueError as e:
                raise RuntimeError(f"行 {r_i}: {e}")

        rows.append((r_i, caption, paths, kind, note, explicit_dt))

    return rows


def schedule_text_post(fb, page_id, page_token, message, ts):
    return fb.post(f"{page_id}/feed", access_token=page_token, message=message,
                   published="false", scheduled_publish_time=str(ts))


def schedule_single_photo(fb, page_id, page_token, image, message, ts):
    with open(image, "rb") as f:
        return fb.post(f"{page_id}/photos", access_token=page_token, caption=message,
                       published="false", scheduled_publish_time=str(ts),
                       files={"source": (image.name, f, "application/octet-stream")})


def upload_unpublished_photo(fb, page_id, page_token, image):
    with open(image, "rb") as f:
        r = fb.post(f"{page_id}/photos", access_token=page_token, published="false",
                    files={"source": (image.name, f, "application/octet-stream")})
    return r["id"]


def schedule_multi_photo(fb, page_id, page_token, images, message, ts):
    media_ids = [upload_unpublished_photo(fb, page_id, page_token, p) for p in images]
    data = {"access_token": page_token, "message": message, "published": "false",
            "scheduled_publish_time": str(ts)}
    for i, mid in enumerate(media_ids):
        data[f"attached_media[{i}]"] = json.dumps({"media_fbid": mid})
    r = requests.post(f"{fb.base}/{page_id}/feed", data=data,
                      timeout=600, proxies=fb.proxies, verify=False)
    return fb._check(r)


def _post_with_retry(url, max_retries=3, **kwargs):
    """对单个 POST 加指数退避重试，专为视频分块上传准备。"""
    last_err = None
    for attempt in range(max_retries):
        try:
            return requests.post(url, **kwargs)
        except (requests.exceptions.ConnectionError,
                requests.exceptions.Timeout,
                requests.exceptions.ChunkedEncodingError) as e:
            last_err = e
            if attempt < max_retries - 1:
                time.sleep(2 ** attempt)
    raise last_err


def schedule_video(fb, page_id, page_token, video, message, ts, log=None):
    """分块上传：start -> transfer*N -> finish。
    每片都重试，公司网关切大流量也能走完。
    """
    file_size = video.stat().st_size
    size_mb = file_size / 1024 / 1024
    if log:
        log.info(f"    视频 {video.name}（{size_mb:.1f} MB）开始分块上传…")

    base_url = f"{fb.base}/{page_id}/videos"
    common = {"timeout": (60, 600), "proxies": fb.proxies, "verify": False}

    # 阶段 1：start
    r = _post_with_retry(base_url, data={
        "access_token": page_token,
        "upload_phase": "start",
        "file_size": str(file_size),
    }, **common)
    start_data = fb._check(r)
    upload_session_id = start_data["upload_session_id"]
    video_id = start_data.get("video_id")
    start_offset = int(start_data["start_offset"])
    end_offset = int(start_data["end_offset"])
    if log:
        chunk_mb = (end_offset - start_offset) / 1024 / 1024
        log.info(f"    会话 id={upload_session_id} video_id={video_id}，分片大小约 {chunk_mb:.1f} MB")

    # 阶段 2：transfer
    t0 = time.time()
    chunk_idx = 0
    with open(video, "rb") as f:
        while start_offset < end_offset:
            chunk_idx += 1
            f.seek(start_offset)
            chunk = f.read(end_offset - start_offset)
            files = {"video_file_chunk": (video.name, chunk, "application/octet-stream")}
            r = _post_with_retry(base_url, data={
                "access_token": page_token,
                "upload_phase": "transfer",
                "upload_session_id": upload_session_id,
                "start_offset": str(start_offset),
            }, files=files, **common)
            t = fb._check(r)
            new_start = int(t["start_offset"])
            new_end = int(t["end_offset"])
            if log:
                pct = new_start / file_size * 100 if file_size else 100
                log.info(f"    分片 {chunk_idx}: {start_offset}-{end_offset} -> 下一片 {new_start}-{new_end}（{pct:.0f}%）")
            if new_start == start_offset:
                raise RuntimeError(f"分片上传卡住：start_offset 没推进（{start_offset}）")
            start_offset, end_offset = new_start, new_end

    elapsed = time.time() - t0
    speed = size_mb / elapsed if elapsed > 0 else 0
    if log:
        log.info(f"    分块全部上传完成，耗时 {elapsed:.1f}s，平均 {speed:.2f} MB/s")

    # 阶段 3：finish
    r = _post_with_retry(base_url, data={
        "access_token": page_token,
        "upload_phase": "finish",
        "upload_session_id": upload_session_id,
        "description": message,
        "published": "false",
        "scheduled_publish_time": str(ts),
    }, **common)
    finish_data = fb._check(r)
    if not finish_data.get("success"):
        raise RuntimeError(f"finish 阶段返回 success=false: {finish_data}")
    if log:
        log.info(f"    视频排期完成，video_id={video_id}")
    return {"id": video_id}


def publish(fb, page_id, page_token, post, log=None):
    ts = int(post.scheduled_at.timestamp())
    if post.media_kind == "text":
        return schedule_text_post(fb, page_id, page_token, post.caption, ts)
    if post.media_kind == "photo":
        if len(post.media_files) == 1:
            return schedule_single_photo(fb, page_id, page_token, post.media_files[0], post.caption, ts)
        return schedule_multi_photo(fb, page_id, page_token, post.media_files, post.caption, ts)
    if post.media_kind == "video":
        return schedule_video(fb, page_id, page_token, post.media_files[0], post.caption, ts, log=log)
    raise ValueError(f"未知类型: {post.media_kind}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default="config.json")
    ap.add_argument("--dry-run", action="store_true", help="只打印计划，不真发")
    ap.add_argument("--start", help="覆盖 start_date（today / tomorrow / YYYY-MM-DD）")
    ap.add_argument("--only-rows", help="只发指定行号（逗号分隔，例如 '12,16,17,18,19'）")
    args = ap.parse_args()

    cfg_path = Path(args.config)
    if not cfg_path.exists():
        print(f"找不到配置: {cfg_path}（先把 config.example.json 复制为 config.json）")
        sys.exit(1)
    cfg = Config.load(cfg_path)

    log = setup_logger(Path(cfg.log_file))
    log.info("=" * 60)
    log.info("FB 主页定时发帖")

    tz = pytz.timezone(cfg.timezone)
    start_date = parse_start_date(args.start or cfg.start_date_raw, tz)
    log.info(f"起始日期: {start_date}  时段: {cfg.daily_slots}  时区: {cfg.timezone}")

    idx = MediaIndex.build(cfg.media_root, log)

    rows = read_posts(cfg, idx, tz, log)
    log.info(f"读到 {len(rows)} 条有效帖子")

    if args.only_rows:
        wanted = {int(x.strip()) for x in args.only_rows.split(",") if x.strip()}
        before = len(rows)
        rows = [r for r in rows if r[0] in wanted]
        log.info(f"--only-rows 过滤：{before} -> {len(rows)} 条（保留行 {sorted(wanted)}）")

    if not rows:
        log.info("没活儿，退出。")
        return

    auto_indices = [i for i, r in enumerate(rows) if r[5] is None]
    auto_slots = gen_slots(start_date, cfg.daily_slots, tz, cfg.skip_weekends, len(auto_indices))
    auto_iter = iter(auto_slots)

    posts = []
    for r_i, cap, paths, kind, note, explicit_dt in rows:
        sched_at = explicit_dt if explicit_dt is not None else next(auto_iter)
        posts.append(Post(row=r_i, caption=cap, media_files=paths,
                          media_kind=kind, scheduled_at=sched_at, note=note))

    now = datetime.now(tz)
    min_t = now + timedelta(minutes=11)
    max_t = now + timedelta(days=179)
    for p in posts:
        if p.scheduled_at < min_t:
            raise RuntimeError(
                f"行 {p.row} 的计划时间 {p.scheduled_at} 距现在不足 10 分钟，FB 不接受。"
            )
        if p.scheduled_at > max_t:
            raise RuntimeError(
                f"行 {p.row} 的计划时间 {p.scheduled_at} 超出 FB 6 个月上限。"
            )

    log.info("-" * 60)
    log.info("计划:")
    for p in posts:
        files = ", ".join(x.name for x in p.media_files) or "(无)"
        cap_short = (p.caption[:30] + "…") if len(p.caption) > 30 else p.caption
        log.info(f"  行{p.row}  {p.scheduled_at.strftime('%Y-%m-%d %H:%M')}  "
                 f"[{p.media_kind}] {files}  | {cap_short}")
        if p.note:
            log.info(f"         {p.note}")
    log.info("-" * 60)

    if args.dry_run:
        log.info("DRY-RUN 完成，未调用 FB API。")
        write_result(cfg.result_file, posts, results=None)
        return

    fb = FB(cfg.access_token, cfg.api_version, log, proxy=cfg.proxy)
    page_id, page_token = fb.resolve_page(cfg.page)
    log.info(f"主页解析成功: id={page_id}")

    results = []
    for p in posts:
        try:
            log.info(f"  → 发布 行{p.row}（{p.media_kind}）@ {p.scheduled_at.strftime('%Y-%m-%d %H:%M')} …")
            r = publish(fb, page_id, page_token, p, log=log)
            results.append((p, r, None))
            log.info(f"  ✓ 行{p.row} -> id={r.get('post_id') or r.get('id')} @ {p.scheduled_at}")
            time.sleep(1)
        except Exception as e:
            results.append((p, None, str(e)))
            log.error(f"  ✗ 行{p.row} 失败: {e}")

    log.info("-" * 60)
    log.info("自检：逐条向 FB 查询排期状态…")
    verified_map = {}
    for p, r, err in results:
        if err is not None or r is None:
            continue
        pid = r.get("post_id") or r.get("id") or ""
        expected = int(p.scheduled_at.timestamp())
        try:
            if p.media_kind == "video":
                # 视频对象没有 published 字段；查 status + scheduled_publish_time
                info = fb.get(pid, fields="id,status,scheduled_publish_time",
                              access_token=page_token)
                sched_ts = info.get("scheduled_publish_time")
                status = info.get("status") or {}
                vstatus = status.get("video_status") or status.get("uploading_phase", {}).get("status")
                # 排期视频在 finish 后立刻查可能还在 processing/uploaded，都算正常
                bad_status = vstatus in ("error", "expired")
                ok = (not bad_status) and (sched_ts is not None) and (abs(int(sched_ts) - expected) <= 60)
                if ok:
                    when = datetime.fromtimestamp(int(sched_ts), tz).strftime("%Y-%m-%d %H:%M")
                    detail = f"将于 {when} 发布（视频状态: {vstatus or 'unknown'}）"
                    log.info(f"  ✓ 行{p.row} {pid}  {detail}")
                else:
                    detail = f"video_status={vstatus} scheduled_publish_time={sched_ts}（期望 {expected}）"
                    log.error(f"  ✗ 行{p.row} {pid} 自检不通过: {detail}")
            else:
                info = fb.get(pid, fields="id,published,scheduled_publish_time",
                              access_token=page_token)
                sched_ts = info.get("scheduled_publish_time")
                is_pub = info.get("published")
                ok = (is_pub is False) and (sched_ts is not None) and (abs(int(sched_ts) - expected) <= 60)
                if ok:
                    when = datetime.fromtimestamp(int(sched_ts), tz).strftime("%Y-%m-%d %H:%M")
                    detail = f"将于 {when} 发布"
                    log.info(f"  ✓ 行{p.row} {pid}  {detail}")
                else:
                    detail = f"published={is_pub} scheduled_publish_time={sched_ts}（期望 {expected}）"
                    log.error(f"  ✗ 行{p.row} {pid} 自检不通过: {detail}")
            verified_map[p.row] = (ok, detail)
        except Exception as e:
            verified_map[p.row] = (False, f"查询失败: {e}")
            log.error(f"  ✗ 行{p.row} {pid} 查询失败: {e}")

    write_result(cfg.result_file, posts, results, page_id=page_id, verified_map=verified_map)
    ok_n = sum(1 for _, r, err in results if err is None)
    all_verified = bool(results) and all(
        verified_map.get(p.row, (False, ""))[0]
        for p, r, err in results if err is None
    )
    log.info(f"完成：{ok_n}/{len(results)} 成功。结果写入 {cfg.result_file}")
    if ok_n == len(results) and all_verified:
        log.info("=" * 60)
        log.info("✓ 成功定时")
        log.info("=" * 60)
    elif ok_n == len(results) and not all_verified:
        log.warning("发布全部成功，但部分自检未通过，请打开 result.csv 里的 preview_url 人工确认。")
    else:
        log.warning(f"有 {len(results) - ok_n} 条失败，详情见 result.csv 的 error 列。")


def write_result(path, posts, results=None, page_id=None, verified_map=None):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.wri