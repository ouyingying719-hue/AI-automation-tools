# FB 广告消耗自动导入 SmartBI

每天定时从 Facebook Marketing API 拉取广告消耗数据，转换格式后通过 SmartBI HTTP API 导入 BI 系统，完成后发送钉钉通知。

## 架构

```
任务一：下载消耗.py → 转换脚本.py → cli_branch/上传BI_cli.py
                                     （CLI HTTP API 同步上传）

任务二：下载KOL消耗.py → 转换KOL格式.py → cli_branch/上传BI_任务二_cli.py
                                          （fire-and-forget + 等5分钟 + post_verify 行数对账）
```

- **任务一**：课包城市消耗数据回写（按城市+课包聚合，~30 行）
- **任务二**：投放账户数据回写（广告级别明细，~1000+ 行）

任务二数据量大，上传响应可能被 CDN 超时拦截，采用 fire-and-forget 模式：
触发上传 → 等 5 分钟 → 查询 BI 验证报表当日行数，与上传文件行数对账。

## 文件结构

```
├── run_all_v2.py                    # 主编排脚本
├── 下载消耗.py                      # 任务一：FB API 下载
├── 下载KOL消耗.py                   # 任务二：FB API 下载
├── 转换脚本.py                      # 任务一：格式转换
├── 转换KOL格式.py                   # 任务二：格式转换
├── 发送钉钉.py                      # 钉钉通知
├── cli_branch/
│   ├── 上传BI_cli.py               # 任务一 CLI 上传
│   ├── 上传BI_任务二_cli.py         # 任务二 CLI 上传（fire-and-forget）
│   └── smartbi_writeback_tasks.json # 上传目标 + 验证报表配置
├── 每日自动运行.bat                  # Windows 计划任务入口
├── 注册每日任务.bat                  # 注册每天 08:00 定时任务
├── 取消每日任务.bat                  # 取消定时任务
├── 手动测试全流程.bat                # 手动触发完整流程
├── 测试CLI上传.bat                  # 单测任务一上传
├── 测试CLI任务二.bat                # 单测任务二上传（约6分钟）
├── .env.example                     # 配置模板
└── .gitignore
```

## 部署

### 1. 环境

- Windows 10/11
- Python 3.8+（PATH 中可用）
- `pip install openpyxl requests`
- SmartBI CLI 工具（内部工具，需单独获取）

### 2. 配置

```bash
cp .env.example .env
# 编辑 .env 填入实际值
```

### 3. 准备规则表

将 `FB账户消耗底表.xlsx` 放在同目录（包含"地区名称对照表"和"账户-平台映射"工作表）。

### 4. 配置 SmartBI 目标

编辑 `cli_branch/smartbi_writeback_tasks.json`，填入实际的 report_id：
- `course_city_cost_writeback.target.report_id`：任务一导入配置 ID
- `course_city_cost_writeback.post_verify.report.report_id`：任务一验证报表 ID
- `ad_account_writeback.target.report_id`：任务二导入配置 ID
- `ad_account_writeback.post_verify.report.report_id`：任务二验证报表 ID

### 5. 测试

```bash
# 单测任务一上传（< 1 分钟）
测试CLI上传.bat

# 单测任务二上传（约 6 分钟）
测试CLI任务二.bat

# 全流程
手动测试全流程.bat
```

### 6. 注册定时任务

右键 `注册每日任务.bat` → 以管理员身份运行。

## 注意事项

- FB Access Token 有效期约 60 天，过期后更新 `.env`
- 任务二 post_verify 容忍 5% 行数偏差（去重/聚合可能造成）
- 钉钉通知包含各任务行数和对账结果
- 如需更换验证报表，修改 `smartbi_writeback_tasks.json` 中对应的 `post_verify.report.report_id`
