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
- 客户验证码页：`/code`
  - 粘贴账号信息后点击“获取验证码”
  - 默认只返回最近 30 秒内的 OpenAI 验证码
  - 页面通过 `POST /api/v1/code` 调用后端，不把 token 放进地址栏
- 邮件列表页：`/sms` 或 `/sms/list`
  - 输入账号行：`email----password----client_id----token`
  - 点击“读取全部邮件”动态显示邮件列表（日期/发件人/主题）
  - 读取和打开全文时显示转圈和进度条
  - 点击“查看全文”用 HTML 预览打开邮件原貌
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

### 直接获取最新验证码接口

如果只想拿验证码，不需要邮件表格，可以调用：

- `GET /api/sms/code`
- `POST /api/sms/code`

GET 示例（账号行必须 URL 编码）：

```text
http://127.0.0.1:2020/api/sms/code?account_line=<urlencoded email----password----client_id----token>
```

POST 表单示例：

```powershell
Invoke-RestMethod -Method Post `
  -Uri http://127.0.0.1:2020/api/sms/code `
  -Body @{ account_line = "email----password----client_id----refresh_token"; top = "20" }
```

返回 JSON 会把验证码放在顶层 `code` 字段。只要纯文本验证码时，调用 `/api/sms/code.txt` 或加 `format=text`。

正式程序接口：

```http
POST /api/v1/code
Content-Type: application/json
```

请求示例：

```json
{
  "account_line": "email----password----client_id----refresh_token",
  "provider": "openai",
  "within_seconds": 30,
  "top": 20
}
```

成功示例：

```json
{
  "ok": true,
  "code": "123456",
  "provider": "openai",
  "email": "user@outlook.com",
  "age_seconds": 8,
  "subject": "..."
}
```

浏览器直接显示一句话的接口：

```text
http://127.0.0.1:2020/api/text-relay/<urlencoded email----password----client_id----token>
```

这个接口默认只返回邮件时间在最近 30 秒内的验证码。可以用 `recent_seconds` 改窗口：

```text
http://127.0.0.1:2020/api/text-relay/<账号信息>?recent_seconds=60
```

成功示例：

```text
YES|您的 OpenAI 验证代码是: 123456
```

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
- Customer code page: `/code`
  - Paste account info and click **Get Code**
  - Defaults to OpenAI codes from emails dated within the last 30 seconds
  - Uses `POST /api/v1/code` so the token is not placed in the browser address bar
- Mail list page: `/sms` or `/sms/list`
  - Input account line: `email----password----client_id----token`
  - Click **Fetch Mail** to dynamically load the message list (date/from/subject)
  - Shows a spinner and progress bar while loading
  - Click **View Full** to preview the original HTML email
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

### Direct Latest-Code API

Use this when you only need the newest verification code instead of the full mail table:

- `GET /api/sms/code`
- `POST /api/sms/code`

GET example:

```text
http://127.0.0.1:2020/api/sms/code?account_line=<urlencoded email----password----client_id----token>
```

POST form example:

```powershell
Invoke-RestMethod -Method Post `
  -Uri http://127.0.0.1:2020/api/sms/code `
  -Body @{ account_line = "email----password----client_id----refresh_token"; top = "20" }
```

The JSON response puts the verification code at top-level `code`. For code-only text output, call `/api/sms/code.txt` or add `format=text`.

Primary programmatic API:

```http
POST /api/v1/code
Content-Type: application/json
```

Request example:

```json
{
  "account_line": "email----password----client_id----refresh_token",
  "provider": "openai",
  "within_seconds": 30,
  "top": 20
}
```

Success example:

```json
{
  "ok": true,
  "code": "123456",
  "provider": "openai",
  "email": "user@outlook.com",
  "age_seconds": 8,
  "subject": "..."
}
```

Browser-friendly one-line relay:

```text
http://127.0.0.1:2020/api/text-relay/<urlencoded email----password----client_id----token>
```

This endpoint only returns codes from emails dated within the last 30 seconds by default. Use `recent_seconds` to change the window:

```text
http://127.0.0.1:2020/api/text-relay/<account-line>?recent_seconds=60
```

Success example:

```text
YES|您的 OpenAI 验证代码是: 123456
```

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
