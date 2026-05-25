# FB 广告预览链接自动记录

自动从 Facebook Marketing API 拉取广告，按"广告名称 = 表格中素材名称"匹配，把广告的公开预览链接写入 Excel 的「预览链接」列。新增素材会等满 24 小时再查询（给广告留出审核+投放时间）。

## 特性

- **按账户路由**：根据素材名关键词（如"繁体"、"粤语"、"普通话"）优先去对应广告账户查找
- **公开链接优先**：取广告创意的「获得评论的 Facebook 帖子」公开链接，不需要登录就能打开；如果该字段为空，再退回登录态预览或广告库链接
- **增量缓存**：首次全量拉取后写入 `ad_cache.json`，后续运行只拉"上次同步以来更新过的广告"，从 30 分钟降到 30 秒-2 分钟
- **24 小时延迟**：新发现的素材记录 first_seen，满 24 小时后才尝试查询，避免广告还没上线就 MISS
- **不覆盖已有链接**：表格里手动填好的链接行直接跳过

## 快速上手

### 1. 准备依赖
- Python 3.8+（安装时勾选 Add to PATH）
- 一张待填充「预览链接」列的 Excel 表

### 2. 拿 FB 长期访问令牌
1. 访问 https://developers.facebook.com/tools/explorer/
2. 选择应用，勾选权限：`ads_read`, `ads_management`, `business_management`
3. Generate Access Token，再用 [Access Token Debugger](https://developers.facebook.com/tools/debug/accesstoken/) 的 Extend 把短期换成约 60 天长期
4. 推荐：[Business Settings → System Users](https://business.facebook.com/settings/system-users) 创建系统用户，生成永不过期的令牌

### 3. 配置
```bash
cp config.example.json config.json
```
编辑 `config.json`：
- `access_token` 填上一步拿到的令牌
- `ad_account_ids` 填广告账户 ID（带 `act_` 前缀）
- `routing_rules` 按你的命名规范调整
- `excel_path` / `sheet_name` / `name_column` / `preview_column` 对应你的表格

### 4. 运行
```bash
# Windows
run.bat

# 或直接 Python
python fb_preview_updater.py
```

首次约 15-30 分钟（全量拉取），之后每次 30 秒-2 分钟（增量）。

## 命令行参数

```
python fb_preview_updater.py [--dry-run] [--full-refresh] [--offline] [--ignore-delay] [--config PATH]
```

- `--dry-run`：不调 FB API，只统计待查询数量和路由分布
- `--full-refresh`：忽略缓存重新全量拉取
- `--offline`：仅用现有 `ad_cache.json` 匹配，不调 API
- `--ignore-delay`：忽略 24 小时延迟，立即处理所有空白行

## 路由规则

`config.json` 的 `routing_rules` 决定按素材名关键词去哪些账户优先查找。

```json
"routing_rules": [
  {
    "keywords": ["繁体"],
    "accounts": ["act_xxx", "act_yyy"]
  }
]
```

规则按"包含"匹配，多条命中时账户取并集。优先账户没找到时降级到其他账户。

## 自动化（Windows 任务计划程序）

1. Win+R → `taskschd.msc`
2. 创建基本任务 → 触发器：每天 03:00
3. 操作 → 启动程序 `run.bat`，起始于该项目目录
4. 完成

## 文件说明

| 文件 | 说明 | Git |
| --- | --- | --- |
| `fb_preview_updater.py` | 主脚本 | ✓ |
| `config.example.json` | 配置模板 | ✓ |
| `config.json` | 你的真实配置（含令牌） | **忽略** |
| `run.bat` | Windows 一键运行 | ✓ |
| `素材信息.xlsx` | 你的素材表 | **忽略** |
| `ad_cache.json` | 自动生成的广告缓存 | **忽略** |
| `tracker.json` | 自动生成的素材首次发现时间 | **忽略** |
| `run.log` | 自动生成的运行日志 | **忽略** |

## 安全提示

- `config.json` 含访问令牌，已加入 `.gitignore`，**永远不要提交到公开仓库**
- 如令牌不慎泄露，立刻去 [Access Token Debugger](https://developers.facebook.com/tools/debug/accesstoken/) 点 Invalidate 失效掉，并重新生成

## 常见问题

**Q: 跑完显示"本次写入 0 个"？**
通常是新素材的广告还没在 FB 上线，等几小时再跑；或者那个广告不在配置的账户里。

**Q: 401 / OAuthException 错误？**
令牌过期，按"快速上手 第 2 步"重新生成。

**Q: 网络超时？**
脚本已内置 4 次指数退避重试，仍失败时检查能否在浏览器打开 https://www.facebook.com 。如需走代理，在 `config.json` 加 `"proxy": "http://127.0.0.1:7890"`。

**Q: 想改路由规则？**
直接编辑 `config.json` 的 `routing_rules`，下次运行生效。

## License

MIT
