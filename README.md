# Reward + Outlook SMS (Lightweight)

中文 | [English](#english)

## 中文

这是一个超轻量 Web 工具，包含两个核心功能：

1. 打赏页面（2张图）
2. Outlook 接码页面（OAuth2，快速读取邮件列表 + 查看全文）

主程序文件：

- `app.py`

### 功能说明

- 首页：`/`
  - 打赏
  - 接码
  - 管理员
- 打赏页：`/reward`
  - 固定显示 2 张图（图1、图2）
- 接码页：`/sms`
  - 输入账号行：`email----password----client_id----token`
  - 点击“连接并读取”显示邮件列表（日期/发件人/主题）
  - 点击“刷新最新邮件”拉取新邮件
  - 点击“查看全文”查看该邮件完整内容
- 管理员页：`/admin`
  - 登录后可分别上传打赏图1/图2

### 轻量设计

- 不使用 Flask/Django 等重框架，仅标准库实现
- 列表页默认只拉取邮件头（更快）
- 查看全文按需加载（点击后才拉正文）
- 支持并发请求（`ThreadingHTTPServer`）
- 短时 token 缓存，刷新更快

### 本地运行

```powershell
cd D:\code\demo\outlook_send\github
python .\app.py
```

默认地址：

- `http://127.0.0.1:2020`

注意：不要使用 `http://0.0.0.0:2020` 访问。

### 服务器部署（2020端口）

如果你已经配置好 `systemd` 服务（如 `lc-pay-sms.service`），更新代码后重启：

```bash
sudo systemctl restart lc-pay-sms.service
sudo systemctl status lc-pay-sms.service
```

公网访问：

- `http://<your-server-ip>:2020`

### 管理员默认账号

- 用户名：`admin`
- 密码：`boji1334`

建议上线后立即改为强密码（可通过环境变量 `ADMIN_USER` / `ADMIN_PASS` 设置）。

---

## English

This is a lightweight web tool with two core modules:

1. Reward page (2 images)
2. Outlook SMS page (OAuth2, fast mailbox list + full message view)

Main program:

- `app.py`

### Features

- Home: `/`
  - Reward
  - SMS
  - Admin
- Reward page: `/reward`
  - Shows exactly 2 images (slot 1 and slot 2)
- SMS page: `/sms`
  - Input account line: `email----password----client_id----token`
  - Click **Connect and Fetch** to load message list (date/from/subject)
  - Click **Refresh Latest** to get new messages
  - Click **View Full** to load full content of one message
- Admin page: `/admin`
  - Upload reward image slot 1 / slot 2

### Lightweight Strategy

- Standard library only (no Flask/Django)
- List view fetches headers only by default (fast)
- Full body is fetched on demand per message
- Concurrent handling with `ThreadingHTTPServer`
- Short-lived token cache for faster refresh

### Run Locally

```powershell
cd D:\code\demo\outlook_send\github
python .\app.py
```

Default URL:

- `http://127.0.0.1:2020`

Do not browse with `http://0.0.0.0:2020`.

### Deploy on Server (Port 2020)

If you already have a `systemd` service (e.g. `lc-pay-sms.service`), restart after updates:

```bash
sudo systemctl restart lc-pay-sms.service
sudo systemctl status lc-pay-sms.service
```

Public URL:

- `http://<your-server-ip>:2020`

### Default Admin Credential

- Username: `admin`
- Password: `boji1334`

For production use, change credentials immediately with `ADMIN_USER` / `ADMIN_PASS`.
