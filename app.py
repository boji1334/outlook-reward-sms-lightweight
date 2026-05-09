#!/usr/bin/env python3
"""
Lightweight Reward + Outlook SMS Code Web

Features:
- Home: 打赏 / 接码 / 管理员
- 打赏: 固定 2 张图
- 管理员: 登录后分别上传打赏图1、打赏图2
- 接码: 粘贴 Outlook OAuth2 账号行后，快速读取最近邮件信息（IMAP）

No heavy framework. Standard-library only.
"""

from __future__ import annotations

import html
import hashlib
import imaplib
import json
import os
import re
import secrets
import socket
import ssl
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass
from email import policy
from email.header import decode_header, make_header
from email.parser import BytesParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Iterable

# -----------------------------
# Basic config
# -----------------------------
APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)

HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", "2020"))

ADMIN_USER = os.environ.get("ADMIN_USER", "admin")
ADMIN_PASS = os.environ.get("ADMIN_PASS", "boji1334")

SESSION_NAME = "lc_admin_session"
SESSION_TTL = 24 * 3600
MAX_UPLOAD = 5 * 1024 * 1024
SESSIONS: dict[str, int] = {}
TOKEN_CACHE: dict[str, tuple[str, int, str]] = {}
TOKEN_CACHE_TTL = 900
TOKEN_CACHE_MAX = 64

ALLOWED_EXT = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/jpg": ".jpg",
    "image/gif": ".gif",
    "image/webp": ".webp",
}

# -----------------------------
# Outlook IMAP OAuth2 config
# -----------------------------
DEFAULT_IMAP_HOST = "outlook.office365.com"
DEFAULT_IMAP_PORT = 993
LIVE_TOKEN_URL = "https://login.live.com/oauth20_token.srf"
AAD_TOKEN_URL_TMPL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"


@dataclass
class AccountPayload:
    email: str
    password: str
    client_id: str
    token: str


@dataclass
class MailRow:
    msg_id: str
    date: str
    sender: str
    subject: str


@dataclass
class SmsFetchResult:
    account: str
    token_source: str
    server: str
    used_folder: str
    rows: list[MailRow]
    attempts: list[str]
    raw_line: str
    top: int


@dataclass
class SmsDetailResult:
    account: str
    token_source: str
    server: str
    used_folder: str
    msg_id: str
    subject: str
    sender: str
    receiver: str
    date: str
    body_text: str
    attempts: list[str]
    raw_line: str


def now_ts() -> int:
    return int(time.time())


def cleanup_sessions() -> None:
    t = now_ts()
    expired = [k for k, v in SESSIONS.items() if v < t]
    for k in expired:
        SESSIONS.pop(k, None)


def decode_mime_header(value: str) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value


def token_cache_key(account: AccountPayload) -> str:
    raw = f"{account.email}|{account.client_id}|{account.token}".encode("utf-8", errors="ignore")
    return hashlib.sha1(raw).hexdigest()


def cleanup_token_cache() -> None:
    now = now_ts()
    expired = [k for k, v in TOKEN_CACHE.items() if v[1] <= now]
    for k in expired:
        TOKEN_CACHE.pop(k, None)
    if len(TOKEN_CACHE) > TOKEN_CACHE_MAX:
        # Keep entries with later expiry.
        ordered = sorted(TOKEN_CACHE.items(), key=lambda kv: kv[1][1], reverse=True)
        keep = dict(ordered[:TOKEN_CACHE_MAX])
        TOKEN_CACHE.clear()
        TOKEN_CACHE.update(keep)


def get_cached_access_token(account: AccountPayload) -> tuple[str, str] | None:
    cleanup_token_cache()
    item = TOKEN_CACHE.get(token_cache_key(account))
    if not item:
        return None
    access_token, exp_ts, source = item
    if exp_ts <= now_ts():
        TOKEN_CACHE.pop(token_cache_key(account), None)
        return None
    return (source, access_token)


def set_cached_access_token(account: AccountPayload, access_token: str, source: str) -> None:
    cleanup_token_cache()
    normalized_source = source
    while normalized_source.startswith("cache_"):
        normalized_source = normalized_source[len("cache_") :]
    if not normalized_source:
        normalized_source = "direct"
    TOKEN_CACHE[token_cache_key(account)] = (access_token, now_ts() + TOKEN_CACHE_TTL, normalized_source)


# -----------------------------
# Reward image helpers
# -----------------------------
def image_base_name(slot: int) -> str:
    return f"reward_image_{slot}"


def current_image_path(slot: int):
    base = image_base_name(slot)
    for ext in (".png", ".jpg", ".gif", ".webp"):
        p = DATA_DIR / (base + ext)
        if p.exists():
            return p
    return None


def remove_slot_images(slot: int) -> None:
    base = image_base_name(slot)
    for ext in (".png", ".jpg", ".gif", ".webp"):
        p = DATA_DIR / (base + ext)
        if p.exists():
            try:
                p.unlink()
            except Exception:
                pass


# -----------------------------
# Outlook token + fetch helpers
# -----------------------------
def parse_account_line(raw: str) -> AccountPayload:
    parts = [p.strip() for p in raw.strip().split("----")]
    if len(parts) < 4:
        raise ValueError("账号信息格式错误，应为: email----password----client_id----token")
    email, password, client_id, token = parts[0], parts[1], parts[2], parts[3]
    client_id = "".join(client_id.split())
    token = "".join(token.split())
    if not email or "@" not in email:
        raise ValueError("邮箱格式错误")
    if not token:
        raise ValueError("token 为空")
    return AccountPayload(email=email, password=password, client_id=client_id, token=token)


def build_xoauth2_bytes(email: str, access_token: str) -> bytes:
    return f"user={email}\x01auth=Bearer {access_token}\x01\x01".encode("utf-8")


def looks_like_access_token(token: str) -> bool:
    # Common refresh token prefix for Microsoft consumer accounts.
    if token.startswith("M.C"):
        return False
    # Common JWT-like access token prefix.
    if token.startswith("eyJ") and token.count(".") >= 2:
        return True
    # Some access tokens start with Ew...
    if token.startswith("Ew"):
        return True
    return False


def token_post(url: str, payload: dict[str, str], timeout: int) -> dict:
    data = urllib.parse.urlencode(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        body = resp.read().decode("utf-8", errors="replace")
    result = json.loads(body)
    return result


def refresh_token_via_live(client_id: str, refresh_token: str, timeout: int) -> str:
    payload = {
        "client_id": client_id,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "redirect_uri": "https://login.live.com/oauth20_desktop.srf",
    }
    data = token_post(LIVE_TOKEN_URL, payload, timeout)
    access_token = data.get("access_token", "")
    if not access_token:
        raise RuntimeError(data.get("error_description") or data.get("error") or "refresh_live failed")
    return access_token


def refresh_token_via_aad(client_id: str, refresh_token: str, tenant: str, timeout: int) -> str:
    url = AAD_TOKEN_URL_TMPL.format(tenant=tenant)
    payload = {
        "client_id": client_id,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "scope": "https://outlook.office.com/IMAP.AccessAsUser.All offline_access",
    }
    data = token_post(url, payload, timeout)
    access_token = data.get("access_token", "")
    if not access_token:
        raise RuntimeError(data.get("error_description") or data.get("error") or "refresh_aad failed")
    return access_token


def refresh_candidates(
    account: AccountPayload, refresh_endpoint: str, tenant: str, timeout: int
) -> Iterable[tuple[str, str]]:
    mode = refresh_endpoint
    if mode in ("auto", "live"):
        try:
            yield ("refresh_live", refresh_token_via_live(account.client_id, account.token, timeout))
        except Exception:
            if mode == "live":
                raise

    if mode in ("auto", "aad"):
        try:
            yield ("refresh_aad", refresh_token_via_aad(account.client_id, account.token, tenant, timeout))
        except Exception:
            if mode == "aad":
                raise


def token_candidates(
    account: AccountPayload, token_kind: str, refresh_endpoint: str, tenant: str, timeout: int
) -> Iterable[tuple[str, str]]:
    raw = account.token
    yielded = set()

    def put(source: str, token: str):
        key = (source, token)
        if key in yielded:
            return None
        yielded.add(key)
        return key

    if token_kind == "access":
        item = put("direct", raw)
        if item:
            yield item
        return

    if token_kind == "refresh":
        if not account.client_id:
            raise RuntimeError("refresh 模式必须有 client_id")
        for source, token in refresh_candidates(account, refresh_endpoint, tenant, timeout):
            item = put(source, token)
            if item:
                yield item
        return

    # auto
    direct_first = looks_like_access_token(raw)
    if direct_first:
        item = put("direct", raw)
        if item:
            yield item

    if account.client_id:
        for source, token in refresh_candidates(account, refresh_endpoint, tenant, timeout):
            item = put(source, token)
            if item:
                yield item

    if not direct_first:
        item = put("direct", raw)
        if item:
            yield item


def token_candidates_fast(
    account: AccountPayload, token_kind: str, refresh_endpoint: str, tenant: str, timeout: int
) -> Iterable[tuple[str, str]]:
    seen = set()
    cached = get_cached_access_token(account)
    if cached:
        cache_source, cache_token = cached
        key = ("cache_" + cache_source, cache_token)
        seen.add(key)
        yield key

    for source, token in token_candidates(account, token_kind, refresh_endpoint, tenant, timeout):
        key = (source, token)
        if key in seen:
            continue
        seen.add(key)
        yield key


def fetch_imap_headers(
    email_addr: str, access_token: str, folder: str, top: int, timeout: int, host: str, port: int
) -> list[MailRow]:
    old_timeout = socket.getdefaulttimeout()
    socket.setdefaulttimeout(timeout)
    mail = imaplib.IMAP4_SSL(host=host, port=port)
    try:
        xoauth2 = build_xoauth2_bytes(email_addr, access_token)
        mail.authenticate("XOAUTH2", lambda _: xoauth2)

        typ, _ = mail.select(folder, readonly=True)
        if typ != "OK":
            raise RuntimeError(f"无法选择文件夹: {folder}")

        typ, data = mail.search(None, "ALL")
        if typ != "OK" or not data or not data[0]:
            return []

        ids = data[0].split()
        target = ids[-top:] if top > 0 else ids

        rows: list[MailRow] = []
        for msg_id in reversed(target):
            typ, msg_data = mail.fetch(msg_id, "(BODY.PEEK[HEADER.FIELDS (SUBJECT FROM DATE)])")
            if typ != "OK" or not msg_data:
                continue
            header_bytes = b""
            for part in msg_data:
                if isinstance(part, tuple) and len(part) >= 2 and isinstance(part[1], (bytes, bytearray)):
                    header_bytes = bytes(part[1])
                    break
            if not header_bytes:
                continue
            msg = BytesParser(policy=policy.default).parsebytes(header_bytes)
            rows.append(
                MailRow(
                    msg_id=msg_id.decode("ascii", errors="ignore"),
                    date=decode_mime_header(str(msg.get("Date", "")).strip()),
                    sender=decode_mime_header(str(msg.get("From", "")).strip()),
                    subject=decode_mime_header(str(msg.get("Subject", "")).strip()),
                )
            )
        return rows
    finally:
        try:
            mail.logout()
        except Exception:
            pass
        socket.setdefaulttimeout(old_timeout)


def decode_payload(payload: bytes | str | None, charset: str | None) -> str:
    if payload is None:
        return ""
    if isinstance(payload, str):
        return payload
    for cs in [charset, "utf-8", "gb18030", "latin-1"]:
        if not cs:
            continue
        try:
            return payload.decode(cs)
        except Exception:
            continue
    return payload.decode("utf-8", errors="replace")


def html_to_text(raw_html: str) -> str:
    if not raw_html:
        return ""
    text = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", raw_html)
    text = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", text)
    text = re.sub(r"(?is)<[^>]+>", " ", text)
    text = html.unescape(text)
    text = text.replace("\r", "")
    text = re.sub(r"[ \t\f\v]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_body_text(msg) -> str:
    plain_parts: list[str] = []
    html_parts: list[str] = []
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_disposition() == "attachment":
                continue
            ctype = part.get_content_type()
            if ctype not in ("text/plain", "text/html"):
                continue
            raw = part.get_payload(decode=True)
            text = decode_payload(raw, part.get_content_charset())
            if ctype == "text/plain":
                plain_parts.append(text)
            else:
                html_parts.append(text)
    else:
        ctype = msg.get_content_type()
        raw = msg.get_payload(decode=True)
        text = decode_payload(raw, msg.get_content_charset())
        if ctype == "text/html":
            html_parts.append(text)
        else:
            plain_parts.append(text)

    plain = "\n\n".join(p.strip() for p in plain_parts if p and p.strip()).strip()
    if plain:
        return plain
    html_body = "\n\n".join(p for p in html_parts if p).strip()
    if html_body:
        return html_to_text(html_body)
    return ""


def fetch_imap_message_detail(
    email_addr: str, access_token: str, folder: str, msg_id: str, timeout: int, host: str, port: int
) -> dict:
    old_timeout = socket.getdefaulttimeout()
    socket.setdefaulttimeout(timeout)
    mail = imaplib.IMAP4_SSL(host=host, port=port)
    try:
        xoauth2 = build_xoauth2_bytes(email_addr, access_token)
        mail.authenticate("XOAUTH2", lambda _: xoauth2)
        typ, _ = mail.select(folder, readonly=True)
        if typ != "OK":
            raise RuntimeError(f"无法选择文件夹: {folder}")

        typ, msg_data = mail.fetch(msg_id.encode("ascii", errors="ignore"), "(BODY.PEEK[])")
        if typ != "OK" or not msg_data:
            raise RuntimeError(f"无法读取邮件内容: id={msg_id}")

        raw_bytes = b""
        for part in msg_data:
            if isinstance(part, tuple) and len(part) >= 2 and isinstance(part[1], (bytes, bytearray)):
                raw_bytes = bytes(part[1])
                break
        if not raw_bytes:
            raise RuntimeError("邮件内容为空")

        msg = BytesParser(policy=policy.default).parsebytes(raw_bytes)
        return {
            "subject": decode_mime_header(str(msg.get("Subject", "")).strip()),
            "sender": decode_mime_header(str(msg.get("From", "")).strip()),
            "receiver": decode_mime_header(str(msg.get("To", "")).strip()),
            "date": decode_mime_header(str(msg.get("Date", "")).strip()),
            "body_text": extract_body_text(msg),
        }
    finally:
        try:
            mail.logout()
        except Exception:
            pass
        socket.setdefaulttimeout(old_timeout)


def folder_candidates(primary_folder: str) -> list[str]:
    primary = (primary_folder or "INBOX").strip() or "INBOX"
    ordered = [
        primary,
        "INBOX",
        "Inbox",
        "Junk",
        "Junk Email",
        "Spam",
        "垃圾邮件",
        "收件箱",
        "Archive",
    ]
    seen = set()
    out = []
    for name in ordered:
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(name)
    return out


def perform_sms_fetch(
    account_line: str,
    top: int,
    folder: str,
    token_kind: str,
    refresh_endpoint: str,
    tenant: str,
    timeout: int,
) -> SmsFetchResult:
    account = parse_account_line(account_line)
    attempts: list[str] = []
    last_err: str | None = None

    for source, token in token_candidates_fast(
        account=account,
        token_kind=token_kind,
        refresh_endpoint=refresh_endpoint,
        tenant=tenant,
        timeout=timeout,
    ):
        folder_ok_empty = False
        for folder_name in folder_candidates(folder):
            try:
                rows = fetch_imap_headers(
                    email_addr=account.email,
                    access_token=token,
                    folder=folder_name,
                    top=top,
                    timeout=timeout,
                    host=DEFAULT_IMAP_HOST,
                    port=DEFAULT_IMAP_PORT,
                )
                if rows:
                    attempts.append(f"{source}/{folder_name}: success ({len(rows)})")
                    set_cached_access_token(account, token, source)
                    return SmsFetchResult(
                        account=account.email,
                        token_source=source,
                        server=f"{DEFAULT_IMAP_HOST}:{DEFAULT_IMAP_PORT}",
                        used_folder=folder_name,
                        rows=rows,
                        attempts=attempts,
                        raw_line=account_line,
                        top=top,
                    )
                attempts.append(f"{source}/{folder_name}: empty")
                folder_ok_empty = True
            except Exception as exc:
                last_err = str(exc)
                attempts.append(f"{source}/{folder_name}: failed ({exc})")
                if source.startswith("cache_"):
                    TOKEN_CACHE.pop(token_cache_key(account), None)

        # Token worked but common folders have no messages.
        if folder_ok_empty:
            preferred = folder_candidates(folder)[0]
            return SmsFetchResult(
                account=account.email,
                token_source=source,
                server=f"{DEFAULT_IMAP_HOST}:{DEFAULT_IMAP_PORT}",
                used_folder=preferred,
                rows=[],
                attempts=attempts,
                raw_line=account_line,
                top=top,
            )

    if not attempts:
        raise RuntimeError("没有可用 token 候选")
    raise RuntimeError("全部尝试失败: " + (last_err or "unknown error"))


def perform_sms_view(
    account_line: str,
    msg_id: str,
    preferred_folder: str = "INBOX",
    timeout: int = 20,
) -> SmsDetailResult:
    account = parse_account_line(account_line)
    attempts: list[str] = []
    last_err: str | None = None

    for source, token in token_candidates_fast(
        account=account,
        token_kind="auto",
        refresh_endpoint="auto",
        tenant="consumers",
        timeout=timeout,
    ):
        for folder_name in folder_candidates(preferred_folder):
            try:
                detail = fetch_imap_message_detail(
                    email_addr=account.email,
                    access_token=token,
                    folder=folder_name,
                    msg_id=msg_id,
                    timeout=timeout,
                    host=DEFAULT_IMAP_HOST,
                    port=DEFAULT_IMAP_PORT,
                )
                attempts.append(f"{source}/{folder_name}: success")
                set_cached_access_token(account, token, source)
                return SmsDetailResult(
                    account=account.email,
                    token_source=source,
                    server=f"{DEFAULT_IMAP_HOST}:{DEFAULT_IMAP_PORT}",
                    used_folder=folder_name,
                    msg_id=msg_id,
                    subject=detail["subject"],
                    sender=detail["sender"],
                    receiver=detail["receiver"],
                    date=detail["date"],
                    body_text=detail["body_text"],
                    attempts=attempts,
                    raw_line=account_line,
                )
            except Exception as exc:
                last_err = str(exc)
                attempts.append(f"{source}/{folder_name}: failed ({exc})")
                if source.startswith("cache_"):
                    TOKEN_CACHE.pop(token_cache_key(account), None)

    if not attempts:
        raise RuntimeError("没有可用 token 候选")
    raise RuntimeError("读取邮件内容失败: " + (last_err or "unknown error"))


# -----------------------------
# HTML renderer
# -----------------------------
def render_layout(title: str, body_html: str) -> str:
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>{html.escape(title)}</title>
  <style>
    *{{font-family:"Times New Roman", Times, serif}}
    body{{margin:0;background:#f4f6f8;color:#111}}
    .wrap{{max-width:1100px;margin:20px auto;padding:0 16px}}
    .card{{background:#fff;border:1px solid #ddd;border-radius:10px;padding:16px;margin-bottom:14px}}
    h1{{font-size:24px;margin:0 0 12px}}
    h2{{font-size:18px;margin:0 0 10px}}
    h3{{margin:8px 0 10px}}
    .btn{{display:inline-block;background:#0d6efd;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none;border:none;cursor:pointer}}
    .btn.gray{{background:#6c757d}}
    .btn.green{{background:#198754}}
    .row{{display:flex;gap:10px;flex-wrap:wrap;align-items:flex-start}}
    input[type=text],input[type=password],input[type=number],select{{padding:9px;border:1px solid #bbb;border-radius:8px;min-width:120px}}
    textarea{{width:100%;min-height:78px;padding:9px;border:1px solid #bbb;border-radius:8px}}
    .img-box{{background:#fff;border:1px dashed #aaa;padding:10px;border-radius:8px;display:inline-block;width:260px;min-width:260px}}
    .img-box img{{max-width:240px;max-height:320px;width:auto;height:auto;display:block;margin:0 auto;object-fit:contain}}
    .note{{color:#555;font-size:13px}}
    .ok{{color:#198754}}
    .err{{color:#dc3545}}
    .mono{{font-family:"Times New Roman", Times, serif}}
    table{{width:100%;border-collapse:collapse;margin-top:10px}}
    th,td{{border:1px solid #ddd;padding:8px;font-size:14px;vertical-align:top}}
    th{{background:#f0f2f4}}
    .grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:8px}}
    .small{{font-size:12px;color:#666}}
  </style>
</head>
<body>
  <div class="wrap">{body_html}</div>
</body>
</html>"""


def sms_form_html(
    account_line: str = "",
    top: str = "20",
) -> str:
    return f"""
    <div class="card">
      <h1>接码（Outlook 轻量版）</h1>
      <form method="post" action="/sms/fetch">
        <label>账号信息（email----password----client_id----token）</label>
        <textarea name="account_line" placeholder="粘贴完整账号行" required>{html.escape(account_line)}</textarea>
        <div class="grid" style="margin-top:10px;max-width:280px">
          <label>读取数量
            <input type="number" name="top" min="1" max="200" value="{html.escape(top)}">
          </label>
        </div>
        <div class="row" style="margin-top:10px">
          <button class="btn" type="submit">连接并读取</button>
          <a class="btn gray" href="/">返回首页</a>
        </div>
      </form>
      <p class="note">轻量模式：固定读取 INBOX，自动处理 token，仅读取邮件列表（日期/发件人/主题）。</p>
    </div>
    """


def sms_result_html(result: SmsFetchResult) -> str:
    rows_html = ""
    for i, row in enumerate(result.rows, start=1):
        rows_html += (
            "<tr>"
            f"<td>{i}</td>"
            f"<td>{html.escape(row.date)}</td>"
            f"<td>{html.escape(row.sender)}</td>"
            f"<td>{html.escape(row.subject)}</td>"
            "<td>"
            f"<form method=\"post\" action=\"/sms/view\">"
            f"<input type=\"hidden\" name=\"account_line\" value=\"{html.escape(result.raw_line)}\">"
            f"<input type=\"hidden\" name=\"top\" value=\"{result.top}\">"
            f"<input type=\"hidden\" name=\"msg_id\" value=\"{html.escape(row.msg_id)}\">"
            f"<input type=\"hidden\" name=\"folder\" value=\"{html.escape(result.used_folder)}\">"
            f"<button class=\"btn gray\" type=\"submit\">查看全文</button>"
            "</form>"
            "</td>"
            "</tr>"
        )
    if not rows_html:
        rows_html = '<tr><td colspan="5" class="small">没有读取到邮件</td></tr>'

    attempts_html = "<br>".join(html.escape(x) for x in result.attempts)
    return f"""
    <div class="card">
      <h2>读取结果</h2>
      <p class="ok">连接成功</p>
      <p class="small">账号: <span class="mono">{html.escape(result.account)}</span></p>
      <p class="small">服务器: <span class="mono">{html.escape(result.server)}</span> | token来源: <span class="mono">{html.escape(result.token_source)}</span></p>
      <p class="small">实际读取文件夹: <span class="mono">{html.escape(result.used_folder)}</span></p>
      <p class="small">尝试过程:<br>{attempts_html}</p>
      <form method="post" action="/sms/fetch" style="margin-top:8px">
        <input type="hidden" name="account_line" value="{html.escape(getattr(result, 'raw_line', ''))}">
        <input type="hidden" name="top" value="{html.escape(str(getattr(result, 'top', 20)))}">
        <button class="btn" type="submit">刷新最新邮件</button>
      </form>
      <table>
        <thead><tr><th style="width:60px">编号</th><th style="width:250px">日期</th><th>发件人</th><th>主题</th><th style="width:130px">操作</th></tr></thead>
        <tbody>{rows_html}</tbody>
      </table>
    </div>
    """


def sms_error_html(err: str) -> str:
    return f"""
    <div class="card">
      <h2>读取结果</h2>
      <p class="err">读取失败：{html.escape(err)}</p>
    </div>
    """


def sms_detail_html(detail: SmsDetailResult, top: int) -> str:
    attempts_html = "<br>".join(html.escape(x) for x in detail.attempts)
    body = detail.body_text or "(空内容)"
    return f"""
    <div class="card">
      <h2>邮件完整内容</h2>
      <p class="small">账号: <span class="mono">{html.escape(detail.account)}</span></p>
      <p class="small">服务器: <span class="mono">{html.escape(detail.server)}</span> | token来源: <span class="mono">{html.escape(detail.token_source)}</span></p>
      <p class="small">文件夹: <span class="mono">{html.escape(detail.used_folder)}</span> | 邮件ID: <span class="mono">{html.escape(detail.msg_id)}</span></p>
      <p class="small">尝试过程:<br>{attempts_html}</p>
      <hr>
      <p><b>主题:</b> {html.escape(detail.subject)}</p>
      <p><b>发件人:</b> {html.escape(detail.sender)}</p>
      <p><b>收件人:</b> {html.escape(detail.receiver)}</p>
      <p><b>日期:</b> {html.escape(detail.date)}</p>
      <pre style="white-space:pre-wrap;background:#fafafa;border:1px solid #ddd;border-radius:8px;padding:12px">{html.escape(body)}</pre>
      <form method="post" action="/sms/fetch" style="margin-top:8px">
        <input type="hidden" name="account_line" value="{html.escape(detail.raw_line)}">
        <input type="hidden" name="top" value="{top}">
        <button class="btn" type="submit">返回邮件列表</button>
      </form>
    </div>
    """


def parse_multipart_form(content_type: str, body: bytes) -> tuple[dict[str, str], dict[str, dict]]:
    """Parse multipart/form-data using email parser (stdlib only)."""
    headers = (
        f"Content-Type: {content_type}\r\n"
        "MIME-Version: 1.0\r\n"
        "\r\n"
    ).encode("utf-8")
    msg = BytesParser(policy=policy.default).parsebytes(headers + body)
    fields: dict[str, str] = {}
    files: dict[str, dict] = {}

    if not msg.is_multipart():
        return fields, files

    for part in msg.iter_parts():
        if part.get_content_disposition() != "form-data":
            continue
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue
        filename = part.get_filename()
        payload = part.get_payload(decode=True) or b""
        if filename:
            files[name] = {
                "filename": filename,
                "content_type": part.get_content_type(),
                "data": payload,
            }
        else:
            fields[name] = payload.decode(part.get_content_charset() or "utf-8", errors="ignore").strip()

    return fields, files


class AppHandler(BaseHTTPRequestHandler):
    server_version = "LCMini/1.2"

    def log_message(self, fmt, *args):
        return

    def _send_html(self, html_text: str, code: int = 200):
        data = html_text.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        if not getattr(self, "_head_only", False):
            self.wfile.write(data)

    def _send_bytes(self, payload: bytes, content_type: str, code: int = 200):
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        if not getattr(self, "_head_only", False):
            self.wfile.write(payload)

    def _cookie_value(self, name: str):
        raw = self.headers.get("Cookie", "")
        if not raw:
            return None
        for part in raw.split(";"):
            p = part.strip()
            key = name + "="
            if p.startswith(key):
                return p[len(key) :]
        return None

    def _is_admin(self) -> bool:
        cleanup_sessions()
        token = self._cookie_value(SESSION_NAME)
        if not token:
            return False
        exp = SESSIONS.get(token)
        if not exp:
            return False
        if exp < now_ts():
            SESSIONS.pop(token, None)
            return False
        return True

    def _set_admin_session(self) -> None:
        token = secrets.token_urlsafe(24)
        SESSIONS[token] = now_ts() + SESSION_TTL
        self.send_response(303)
        self.send_header("Location", "/admin")
        self.send_header(
            "Set-Cookie",
            f"{SESSION_NAME}={token}; Path=/; HttpOnly; SameSite=Lax; Max-Age={SESSION_TTL}",
        )
        self.end_headers()

    def _clear_admin_session(self) -> None:
        token = self._cookie_value(SESSION_NAME)
        if token:
            SESSIONS.pop(token, None)
        self.send_response(303)
        self.send_header("Location", "/admin")
        self.send_header(
            "Set-Cookie", f"{SESSION_NAME}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0"
        )
        self.end_headers()

    def _parse_urlencoded(self) -> dict[str, str]:
        length = int(self.headers.get("Content-Length", "0") or "0")
        raw = self.rfile.read(length).decode("utf-8", "ignore")
        parsed = urllib.parse.parse_qs(raw)
        result: dict[str, str] = {}
        for k, v in parsed.items():
            result[k] = (v[0] if v else "").strip()
        return result

    def do_HEAD(self):
        self._head_only = True
        try:
            self.do_GET()
        finally:
            self._head_only = False

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        qs = urllib.parse.parse_qs(parsed.query)

        if path == "/":
            body = """
            <div class="card">
              <h1>功能选择</h1>
              <div class="row">
                <a class="btn" href="/reward">打赏</a>
                <a class="btn gray" href="/sms">接码</a>
                <a class="btn green" href="/admin">管理员</a>
              </div>
            </div>
            """
            self._send_html(render_layout("首页", body))
            return

        if path in ("/reward", "/pay"):
            img1 = current_image_path(1)
            img2 = current_image_path(2)
            box1 = '<p class="note">打赏图1：未上传</p>'
            box2 = '<p class="note">打赏图2：未上传</p>'
            if img1:
                box1 = '<img src="/reward-image?slot=1" alt="打赏图1">'
            if img2:
                box2 = '<img src="/reward-image?slot=2" alt="打赏图2">'
            body = f"""
            <div class="card">
              <h1>打赏页面</h1>
              <div class="row">
                <div class="img-box"><h3>图1</h3>{box1}</div>
                <div class="img-box"><h3>图2</h3>{box2}</div>
              </div>
              <p class="note">管理员可在 /admin 分别替换两张图片。</p>
              <a class="btn gray" href="/">返回首页</a>
            </div>
            """
            self._send_html(render_layout("打赏", body))
            return

        if path == "/reward-image":
            slot = 1
            try:
                slot = int((qs.get("slot", ["1"])[0] or "1").strip())
            except Exception:
                slot = 1
            if slot not in (1, 2):
                slot = 1
            img = current_image_path(slot)
            if not img:
                self._send_html(
                    render_layout("未找到", f'<div class="card"><p>未找到打赏图{slot}。</p></div>'),
                    code=404,
                )
                return
            ctype = {
                ".png": "image/png",
                ".jpg": "image/jpeg",
                ".gif": "image/gif",
                ".webp": "image/webp",
            }.get(img.suffix.lower(), "application/octet-stream")
            self._send_bytes(img.read_bytes(), ctype)
            return

        if path == "/sms":
            self._send_html(render_layout("接码", sms_form_html()))
            return

        if path == "/admin/logout":
            self._clear_admin_session()
            return

        if path == "/admin":
            if not self._is_admin():
                body = """
                <div class="card">
                  <h1>管理员登录</h1>
                  <form method="post" action="/admin/login">
                    <p><input type="text" name="username" placeholder="账号"></p>
                    <p><input type="password" name="password" placeholder="密码"></p>
                    <p><button class="btn" type="submit">登录</button></p>
                  </form>
                  <a class="btn gray" href="/">返回首页</a>
                </div>
                """
                self._send_html(render_layout("管理员登录", body))
                return

            img1 = current_image_path(1)
            img2 = current_image_path(2)
            view1 = '<p class="note">图1未上传</p>'
            view2 = '<p class="note">图2未上传</p>'
            if img1:
                view1 = '<img src="/reward-image?slot=1" alt="图1">'
            if img2:
                view2 = '<img src="/reward-image?slot=2" alt="图2">'
            body = f"""
            <div class="card">
              <h1>管理员面板</h1>
              <p class="note">密码已改为你要求的版本（默认: 1334）。</p>
              <div class="row">
                <div class="img-box"><h3>当前图1</h3>{view1}</div>
                <div class="img-box"><h3>当前图2</h3>{view2}</div>
              </div>
            </div>
            <div class="card">
              <h3>上传打赏图1</h3>
              <form method="post" action="/admin/upload?slot=1" enctype="multipart/form-data">
                <input type="file" name="image" accept="image/png,image/jpeg,image/gif,image/webp" required>
                <button class="btn" type="submit">上传图1</button>
              </form>
              <h3 style="margin-top:16px">上传打赏图2</h3>
              <form method="post" action="/admin/upload?slot=2" enctype="multipart/form-data">
                <input type="file" name="image" accept="image/png,image/jpeg,image/gif,image/webp" required>
                <button class="btn" type="submit">上传图2</button>
              </form>
              <p class="note">支持 png/jpg/gif/webp，单图最大 5MB。</p>
              <div class="row">
                <a class="btn gray" href="/reward">查看打赏页</a>
                <a class="btn gray" href="/admin/logout">退出登录</a>
              </div>
            </div>
            """
            self._send_html(render_layout("管理员面板", body))
            return

        self._send_html(render_layout("404", '<div class="card"><h1>404</h1></div>'), code=404)

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        qs = urllib.parse.parse_qs(parsed.query)

        if path == "/sms/fetch":
            form = self._parse_urlencoded()
            account_line = form.get("account_line", "")
            top = form.get("top", "20")

            try:
                top_i = int(top)
            except Exception:
                top_i = 20

            top_i = min(max(top_i, 1), 200)

            parts = [
                sms_form_html(
                    account_line=account_line,
                    top=str(top_i),
                )
            ]
            try:
                result = perform_sms_fetch(
                    account_line=account_line,
                    top=top_i,
                    folder="INBOX",
                    token_kind="auto",
                    refresh_endpoint="auto",
                    tenant="consumers",
                    timeout=20,
                )
                parts.append(sms_result_html(result))
            except Exception as exc:
                parts.append(sms_error_html(str(exc)))

            self._send_html(render_layout("接码结果", "".join(parts)))
            return

        if path == "/sms/view":
            form = self._parse_urlencoded()
            account_line = form.get("account_line", "")
            msg_id = form.get("msg_id", "")
            folder = form.get("folder", "INBOX") or "INBOX"
            top = form.get("top", "20")

            try:
                top_i = int(top)
            except Exception:
                top_i = 20
            top_i = min(max(top_i, 1), 200)

            parts = [
                sms_form_html(
                    account_line=account_line,
                    top=str(top_i),
                )
            ]
            try:
                detail = perform_sms_view(
                    account_line=account_line,
                    msg_id=msg_id,
                    preferred_folder=folder,
                    timeout=20,
                )
                parts.append(sms_detail_html(detail, top_i))
            except Exception as exc:
                parts.append(sms_error_html(str(exc)))

            self._send_html(render_layout("邮件内容", "".join(parts)))
            return

        if path == "/admin/login":
            form = self._parse_urlencoded()
            user = form.get("username", "")
            password = form.get("password", "")
            if user == ADMIN_USER and password == ADMIN_PASS:
                self._set_admin_session()
                return
            body = """
            <div class="card">
              <h1>登录失败</h1>
              <p class="err">账号或密码不正确。</p>
              <a class="btn gray" href="/admin">返回登录</a>
            </div>
            """
            self._send_html(render_layout("登录失败", body), code=401)
            return

        if path == "/admin/upload":
            if not self._is_admin():
                self._send_html(
                    render_layout(
                        "未授权",
                        '<div class="card"><p>请先登录。</p><a class="btn" href="/admin">去登录</a></div>',
                    ),
                    code=401,
                )
                return

            slot = 1
            try:
                slot = int((qs.get("slot", ["1"])[0] or "1").strip())
            except Exception:
                slot = 1
            if slot not in (1, 2):
                slot = 1

            content_type = self.headers.get("Content-Type", "")
            if "multipart/form-data" not in content_type:
                self._send_html(
                    render_layout("错误", '<div class="card"><p class="err">请求格式错误。</p></div>'),
                    code=400,
                )
                return

            length = int(self.headers.get("Content-Length", "0") or "0")
            if length <= 0 or length > MAX_UPLOAD:
                self._send_html(
                    render_layout("错误", '<div class="card"><p class="err">文件过大或为空（最大5MB）。</p></div>'),
                    code=400,
                )
                return

            body = self.rfile.read(length)
            _, files = parse_multipart_form(content_type, body)
            image_file = files.get("image")
            if not image_file:
                self._send_html(
                    render_layout("错误", '<div class="card"><p class="err">未选择文件。</p></div>'),
                    code=400,
                )
                return

            data = image_file.get("data", b"")
            if not data:
                self._send_html(
                    render_layout("错误", '<div class="card"><p class="err">文件为空。</p></div>'),
                    code=400,
                )
                return
            if len(data) > MAX_UPLOAD:
                self._send_html(
                    render_layout("错误", '<div class="card"><p class="err">文件超过5MB。</p></div>'),
                    code=400,
                )
                return

            mime = str(image_file.get("content_type", "")).lower()
            ext = ALLOWED_EXT.get(mime)
            if not ext:
                fn = str(image_file.get("filename", "")).lower()
                for e in (".png", ".jpg", ".jpeg", ".gif", ".webp"):
                    if fn.endswith(e):
                        ext = ".jpg" if e == ".jpeg" else e
                        break
            if not ext:
                self._send_html(
                    render_layout("错误", '<div class="card"><p class="err">仅支持 png/jpg/gif/webp。</p></div>'),
                    code=400,
                )
                return

            remove_slot_images(slot)
            out = DATA_DIR / (image_base_name(slot) + ext)
            out.write_bytes(data)

            body = f"""
            <div class="card">
              <h1>上传成功</h1>
              <p class="ok">打赏图{slot}已更新。</p>
              <div class="row">
                <a class="btn" href="/reward">查看打赏页</a>
                <a class="btn gray" href="/admin">返回管理</a>
              </div>
            </div>
            """
            self._send_html(render_layout("上传成功", body))
            return

        self._send_html(render_layout("404", '<div class="card"><h1>404</h1></div>'), code=404)


def run() -> None:
    httpd = ThreadingHTTPServer((HOST, PORT), AppHandler)
    print(f"Running on http://{HOST}:{PORT}")
    httpd.serve_forever()


if __name__ == "__main__":
    run()
