# FB 主页定时发帖

读本地 Excel 表（文案 + 素材文件名 + 定时时间），把帖子推进 FB 的定时发布队列。
推进队列后 FB 服务器到点自动发，本机关掉也没事。

支持帖子类型：纯文字 / 单图 / 多图（最多 10 张）/ 单视频。视频走 FB 分块上传（resumable upload），网络抖动也能传完。

## 第一次用

```cmd
pip install -r requirements.txt
copy config.example.json config.json
python make_template.py        # 生成 posts_template.xlsx，改名为 posts.xlsx 后开始填
notepad config.json            # 把 access_token、page、media_root 填好
```

## 表格怎么填

打开 `posts.xlsx`：

| 文案 | 定时时间 | 素材1 | 素材2 | … 素材10 |
| --- | --- | --- | --- | --- |
| 今日上新～ | 2026-06-10 10:00 | demo.jpg |  |  |
| 三图轮播 | 2026-06-11 14:30 | a.jpg | b.jpg | c.jpg |
| 视频帖（自动排） |  | intro.mp4 |  |  |

- **文案** 必填或留空（留空就是纯素材帖，但至少得有素材）。
- **定时时间** 想精确指定就填 `YYYY-MM-DD HH:MM`；留空走 `daily_slots` 自动排。
- **素材列** 写文件名（含后缀），脚本去 `media_root` 递归扫，按文件名精确匹配（不区分大小写）。
- 多图按列顺序决定轮播顺序，**最多 10 张**。
- 视频帖只能放 1 个视频；图片视频不能混发。

## 跑

```cmd
run.bat                          # 双击：先 dry-run 给你看计划，确认后正式发
python fb_scheduler.py --dry-run # 单独看计划
python fb_scheduler.py           # 正式发
python fb_scheduler.py --only-rows 12,16,17  # 只重发指定行
```

## 发完后会做什么

1. 每条帖子调一次 FB API 排期，记录返回的 post_id
2. 自检：再调一次 GET 查发布状态和 `scheduled_publish_time`，确认排期生效
   - 文字/图片帖看 `published` 字段
   - 视频帖看 `status.video_status`（视频对象没有 `published` 字段）
3. 把结果写到 `result.csv`，包括：
   - `post_id` — FB 返回的帖子 ID
   - `preview_url` — `https://www.facebook.com/{page_id}/posts/{post_id}`，浏览器打开应该显示「X 小时后」「X 天后」
   - `verified` — `yes` / `no`，自检是否通过
   - `verify_detail` — 自检详情或失败原因
4. 全部成功 + 全部自检通过 → 终端打印 **「✓ 成功定时」**

## FB 的硬性规则（脚本会校验）

- 计划时间必须在 **当前 +10 分钟 ~ +6 个月** 之间。

## token 说明

需要 Page Access Token（脚本会用 User Token 自动换）。token 要有：
- `pages_manage_posts`
- `pages_read_engagement`
- `pages_show_list`

短期 User Token 默认只活 1 小时，建议换成 60 天的长期 token：

```
https://graph.facebook.com/oauth/access_token
  ?grant_type=fb_exchange_token
  &client_id=YOUR_APP_ID
  &client_secret=YOUR_APP_SECRET
  &fb_exchange_token=SHORT_TOKEN
```

> Graph API Explorer 默认的 App 不一定有 `pages_manage_posts` 权限。如果遇到 `(#200)` 错误，去 developers.facebook.com 新建一个带 "Manage Pages" use case 的 App，用它的 ID + Secret 换 token。

## 网络说明

- 脚本默认 `verify=False` 跳过 SSL 校验，方便走有 SSL 中间人代理的企业网络。如果你在普通家庭/公网环境，可以删掉 `verify=False` 让 requests 走默认证书校验。
- 不需要代理就在 `config.json` 里把 `proxy` 设为 `null`；需要代理就填 `http://127.0.0.1:7890` 这种格式，或设环境变量 `HTTPS_PROXY`。

## 排错

| 现象 | 怎么办 |
| --- | --- |
| `素材目录不存在或访问不到` | 检查 `media_root` 路径；UNC 路径需要先在资源管理器里挂载 |
| `未匹配素材: [...]` | 文件名和实际目录上对不上，检查大小写和后缀 |
| `计划时间距现在不足 10 分钟` | 把 `start_date` 改成 `tomorrow` 或具体日期 |
| `(#200) The permission(s) pages_manage_posts...` | token 权限不够，重新生成带 `pages_manage_posts` 的 token |
| `(#190) Error validating access token` | token 过期，重新生成 |
| Business Suite 看不到帖子 | 这是已知 UI bug。点 `result.csv` 里的 `preview_url` 直接看，显示「X 小时后」就稳了 |

## 文件清单

```
fb_scheduler.py        主脚本
make_template.py       生成表格模板
config.example.json    配置示例（复制为 config.json 后填入真实值）
requirements.txt       Python 依赖
run.bat                Windows 双击启动（dry-run + 确认 + 正式发）
.gitignore             忽略 config.json / posts.xlsx / result.csv / *.log
```

运行时会生成（已被 .gitignore 排除）：

```
config.json            真实配置
posts.xlsx             帖子表
result.csv             发布结果 + 预览链接 + 自检状态（每次覆盖）
run.log                日志（追加）
```
