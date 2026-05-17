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
from datetime import datetime, timezone
from email import policy
from email.header import decode_header, make_header
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Iterable

# -----------------------------
# Basic config
# -----------------------------
APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
ACCOUNT_STORE_PATH = DATA_DIR / "account_store.json"

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


def env_int(name: str, default: int, minimum: int, maximum: int) -> int:
    try:
        value = int(os.environ.get(name, str(default)))
    except Exception:
        value = default
    return max(minimum, min(maximum, value))


SMS_FETCH_TIMEOUT = env_int("SMS_FETCH_TIMEOUT", 12, 5, 60)
SMS_VIEW_TIMEOUT = env_int("SMS_VIEW_TIMEOUT", 15, 5, 60)
SMS_MAX_FOLDER_TRIES = env_int("SMS_MAX_FOLDER_TRIES", 4, 1, 10)
SMS_CODE_SNIPPET_BYTES = env_int("SMS_CODE_SNIPPET_BYTES", 8192, 1024, 65536)
SMS_CODE_FALLBACK_LIMIT = env_int("SMS_CODE_FALLBACK_LIMIT", 3, 0, 10)

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
    code: str = ""


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
# Stored account helpers
# -----------------------------
def mask_secret(value: str, keep: int = 6) -> str:
    value = (value or "").strip()
    if not value:
        return "-"
    if len(value) <= keep:
        return "*" * len(value)
    return "***" + value[-keep:]


def normalize_account_record(record: dict[str, str]) -> dict[str, str] | None:
    account_line = str(record.get("account_line", "")).strip()
    if not account_line:
        return None
    try:
        account = parse_account_line(account_line)
        account_line = "----".join([account.email, account.password, account.client_id, account.token])
        email = account.email
        client_id = account.client_id
        token_tail = account.token[-10:]
    except Exception:
        email = str(record.get("email", "")).strip()
        client_id = str(record.get("client_id", "")).strip()
        token_tail = str(record.get("token_tail", "")).strip()
    return {
        "account_line": account_line,
        "email": email,
        "client_id": client_id,
        "token_tail": token_tail,
        "updated_at": str(record.get("updated_at", "")).strip(),
    }


def load_account_records() -> list[dict[str, str]]:
    try:
        data = json.loads(ACCOUNT_STORE_PATH.read_text(encoding="utf-8"))
    except Exception:
        return []

    raw_records = data if isinstance(data, list) else data.get("accounts", []) if isinstance(data, dict) else []
    if isinstance(data, dict) and data.get("account_line"):
        raw_records = [data]

    records: list[dict[str, str]] = []
    seen: set[str] = set()
    for item in raw_records:
        if not isinstance(item, dict):
            continue
        record = normalize_account_record({str(k): str(v) for k, v in item.items()})
        if not record:
            continue
        account_line = record["account_line"]
        if account_line in seen:
            continue
        seen.add(account_line)
        records.append(record)
    return records


def write_account_records(records: list[dict[str, str]]) -> None:
    ACCOUNT_STORE_PATH.write_text(
        json.dumps({"accounts": records}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def save_account_record(account_line: str) -> None:
    account = parse_account_line(account_line)
    clean_line = "----".join([account.email, account.password, account.client_id, account.token])
    record = {
        "account_line": clean_line,
        "email": account.email,
        "client_id": account.client_id,
        "token_tail": account.token[-10:],
        "updated_at": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
    }
    try:
        records = [x for x in load_account_records() if x.get("account_line") != clean_line]
        records.append(record)
        write_account_records(records)
    except Exception:
        pass


def add_account_lines(raw: str) -> tuple[int, list[str]]:
    added = 0
    errors: list[str] = []
    records = load_account_records()
    existing = {x.get("account_line", "") for x in records}

    for line_no, line in iter_account_blocks(raw):
        try:
            account = parse_account_line(line)
        except Exception as exc:
            errors.append(f"第 {line_no} 行: {exc}")
            continue
        clean_line = "----".join([account.email, account.password, account.client_id, account.token])
        if clean_line in existing:
            continue
        records.append(
            {
                "account_line": clean_line,
                "email": account.email,
                "client_id": account.client_id,
                "token_tail": account.token[-10:],
                "updated_at": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
            }
        )
        existing.add(clean_line)
        added += 1

    if added:
        write_account_records(records)
    return added, errors


# -----------------------------
# Outlook token + fetch helpers
# -----------------------------
ACCOUNT_DELIMITER_RE = re.compile(r"\s*(?:-\s*){4}\s*")
ACCOUNT_START_RE = re.compile(r"^\s*\S+@\S+\.\S+")
ZERO_WIDTH_RE = re.compile(r"[\u200b-\u200f\u202a-\u202e\u2060\ufeff]")
DASH_TRANSLATION = str.maketrans(
    {
        "\u2010": "-",
        "\u2011": "-",
        "\u2012": "-",
        "\u2013": "-",
        "\u2014": "-",
        "\u2015": "-",
        "\u2212": "-",
        "\ufe63": "-",
        "\uff0d": "-",
    }
)


def normalize_account_text(raw: str) -> str:
    text = str(raw or "").translate(DASH_TRANSLATION)
    text = ZERO_WIDTH_RE.sub("", text)
    return text.replace("\u00a0", " ").strip()


def split_account_fields(raw: str) -> list[str]:
    text = normalize_account_text(raw)
    parts = [p.strip() for p in ACCOUNT_DELIMITER_RE.split(text, maxsplit=3)]
    if len(parts) >= 4:
        return parts

    # Fallback for copy/paste that turns the four fields into separate lines.
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if len(lines) >= 4 and "@" in lines[0]:
        return [lines[0], lines[1], lines[2], "".join(lines[3:])]

    return parts


def iter_account_blocks(raw: str) -> Iterable[tuple[int, str]]:
    lines = [
        (line_no, line.strip())
        for line_no, line in enumerate(normalize_account_text(raw).splitlines(), start=1)
        if line.strip()
    ]
    if not lines:
        return

    start_line, current = lines[0][0], [lines[0][1]]
    for line_no, line in lines[1:]:
        if current and ACCOUNT_START_RE.match(line):
            yield start_line, "\n".join(current)
            start_line, current = line_no, [line]
        else:
            current.append(line)

    if current:
        yield start_line, "\n".join(current)


def parse_account_line(raw: str) -> AccountPayload:
    parts = split_account_fields(raw)
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
    refresh_style = raw.startswith("M.C")
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

    if not direct_first and not refresh_style:
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


def selected_mail_count(select_data) -> int:
    if not select_data:
        return 0
    head = select_data[0]
    if isinstance(head, (bytes, bytearray)):
        raw = head.decode("utf-8", errors="ignore")
    else:
        raw = str(head or "")
    m = re.search(r"\d+", raw)
    if not m:
        return 0
    try:
        return int(m.group(0))
    except Exception:
        return 0


def extract_seq_from_fetch_meta(meta) -> str:
    if isinstance(meta, (bytes, bytearray)):
        raw = bytes(meta).decode("utf-8", errors="ignore")
    else:
        raw = str(meta or "")
    m = re.match(r"\s*(\d+)", raw)
    return m.group(1) if m else ""


def is_auth_or_token_error(err: str) -> bool:
    t = err.lower()
    markers = (
        "authenticate failed",
        "authentication failed",
        "invalid credentials",
        "login failed",
        "invalid_grant",
        "xoauth2",
        "a1 no authenticate",
        "access denied",
        "unauthorized",
    )
    return any(x in t for x in markers)


def is_network_error(err: str) -> bool:
    t = err.lower()
    markers = (
        "timed out",
        "timeout",
        "temporary failure",
        "connection reset",
        "network is unreachable",
        "name or service not known",
        "ssl",
        "eof",
        "broken pipe",
    )
    return any(x in t for x in markers)


def fetch_meta_kind(meta) -> str:
    if isinstance(meta, (bytes, bytearray)):
        raw = bytes(meta).decode("utf-8", errors="ignore")
    else:
        raw = str(meta or "")
    up = raw.upper()
    if "HEADER.FIELDS" in up:
        return "header"
    if "BODY[TEXT]" in up or "BODY.PEEK[TEXT]" in up or "TEXT]<" in up:
        return "text"
    return ""


def normalize_snippet_text(raw: bytes) -> str:
    text = decode_payload(raw, None)
    if "<" in text and ">" in text:
        text = html_to_text(text)
    text = text.replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def extract_verification_code(subject: str, snippet: str) -> str:
    s = subject or ""
    b = snippet or ""
    sources = [s, b, s + "\n" + b]

    key_patterns = (
        r"(?is)(?:验证码|临时验证码|verification code|otp|openai\s*代码|code)\D{0,40}(?<!\d)(\d{6})(?!\d)",
        r"(?is)(?<!\d)(\d{6})(?!\d)\D{0,20}(?:验证码|verification code|otp|code)",
    )
    for src in sources:
        for pattern in key_patterns:
            m = re.search(pattern, src)
            if m:
                return m.group(1)

    for src in (s, b):
        all_codes = re.findall(r"(?<!\d)(\d{6})(?!\d)", src)
        if all_codes:
            return all_codes[0]
    return ""


def likely_code_mail(subject: str, sender: str) -> bool:
    text = f"{subject} {sender}".lower()
    hints = (
        "code",
        "otp",
        "verification",
        "verify",
        "login",
        "sign in",
        "openai",
        "chatgpt",
        "\u9a8c\u8bc1\u7801",
        "\u4e34\u65f6",
        "\u767b\u5f55",
    )
    return any(h in text for h in hints)


def fetch_imap_headers(
    email_addr: str, access_token: str, folder: str, top: int, timeout: int, host: str, port: int
) -> list[MailRow]:
    old_timeout = socket.getdefaulttimeout()
    socket.setdefaulttimeout(timeout)
    mail = imaplib.IMAP4_SSL(host=host, port=port)
    try:
        xoauth2 = build_xoauth2_bytes(email_addr, access_token)
        mail.authenticate("XOAUTH2", lambda _: xoauth2)

        typ, select_data = mail.select(folder, readonly=True)
        if typ != "OK":
            raise RuntimeError(f"无法选择文件夹: {folder}")

        total = selected_mail_count(select_data)
        if total <= 0:
            return []

        fetch_count = top if top > 0 else total
        start = max(1, total - fetch_count + 1)
        seq_set = f"{start}:{total}"
        typ, msg_data = mail.fetch(
            seq_set,
            f"(BODY.PEEK[HEADER.FIELDS (SUBJECT FROM DATE)] BODY.PEEK[TEXT]<0.{SMS_CODE_SNIPPET_BYTES}>)",
        )
        if typ != "OK" or not msg_data:
            return []

        grouped: dict[str, dict[str, bytes]] = {}
        for part in msg_data:
            if not (isinstance(part, tuple) and len(part) >= 2):
                continue
            seq = extract_seq_from_fetch_meta(part[0])
            if not seq:
                continue
            kind = fetch_meta_kind(part[0])
            if not kind:
                continue
            payload = part[1]
            if not isinstance(payload, (bytes, bytearray)):
                continue
            raw = bytes(payload)
            if not raw:
                continue

            entry = grouped.setdefault(seq, {"header": b"", "text": b""})
            if kind == "header":
                entry["header"] = raw
            elif kind == "text":
                entry["text"] += raw

        rows: list[MailRow] = []
        for seq in sorted(grouped.keys(), key=lambda x: int(x), reverse=True):
            header_bytes = grouped[seq].get("header", b"")
            text_bytes = grouped[seq].get("text", b"")
            if not header_bytes:
                continue
            msg = BytesParser(policy=policy.default).parsebytes(header_bytes)
            subject = decode_mime_header(str(msg.get("Subject", "")).strip())
            snippet = normalize_snippet_text(text_bytes) if text_bytes else ""
            code = extract_verification_code(subject, snippet)
            rows.append(
                MailRow(
                    msg_id=seq,
                    date=decode_mime_header(str(msg.get("Date", "")).strip()),
                    sender=decode_mime_header(str(msg.get("From", "")).strip()),
                    subject=subject,
                    code=code,
                )
            )
        rows = [x for x in rows if (x.msg_id or x.date or x.sender or x.subject)]
        if fetch_count > 0:
            rows = rows[:fetch_count]

        fallback_used = 0
        for row in rows:
            if row.code:
                continue
            if fallback_used >= SMS_CODE_FALLBACK_LIMIT:
                break
            if not likely_code_mail(row.subject, row.sender):
                continue
            try:
                typ, full_data = mail.fetch(row.msg_id.encode("ascii", errors="ignore"), "(BODY.PEEK[])")
                if typ != "OK" or not full_data:
                    continue
                raw_bytes = b""
                for part in full_data:
                    if isinstance(part, tuple) and len(part) >= 2 and isinstance(part[1], (bytes, bytearray)):
                        raw_bytes = bytes(part[1])
                        break
                if not raw_bytes:
                    continue
                full_msg = BytesParser(policy=policy.default).parsebytes(raw_bytes)
                body_text = extract_body_text(full_msg)
                row.code = extract_verification_code(row.subject, body_text)
                fallback_used += 1
            except Exception:
                continue

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
    folders = folder_candidates(folder)[:SMS_MAX_FOLDER_TRIES]
    if not folders:
        folders = ["INBOX"]

    for source, token in token_candidates_fast(
        account=account,
        token_kind=token_kind,
        refresh_endpoint=refresh_endpoint,
        tenant=tenant,
        timeout=timeout,
    ):
        folder_ok_empty = False
        for folder_name in folders:
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
                if is_auth_or_token_error(last_err) or is_network_error(last_err):
                    # No need to retry other folders with the same bad token/network state.
                    break

        # Token worked but common folders have no messages.
        if folder_ok_empty:
            preferred = folders[0]
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
    timeout: int = SMS_VIEW_TIMEOUT,
) -> SmsDetailResult:
    account = parse_account_line(account_line)
    attempts: list[str] = []
    last_err: str | None = None
    folders = folder_candidates(preferred_folder)[:SMS_MAX_FOLDER_TRIES]
    if not folders:
        folders = ["INBOX"]

    for source, token in token_candidates_fast(
        account=account,
        token_kind="auto",
        refresh_endpoint="auto",
        tenant="consumers",
        timeout=timeout,
    ):
        for folder_name in folders:
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
                if is_auth_or_token_error(last_err) or is_network_error(last_err):
                    break

    if not attempts:
        raise RuntimeError("没有可用 token 候选")
    raise RuntimeError("读取邮件内容失败: " + (last_err or "unknown error"))


def mail_date_ts(value: str) -> float | None:
    try:
        dt = parsedate_to_datetime(value or "")
    except Exception:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.timestamp()


def mail_row_within(row: MailRow, seconds: int) -> bool:
    if seconds <= 0:
        return True
    ts = mail_date_ts(row.date)
    if ts is None:
        return False
    age = time.time() - ts
    return -5 <= age <= seconds


def provider_matches(row: MailRow, provider: str) -> bool:
    provider = (provider or "any").strip().lower()
    if provider in ("", "any", "all", "*"):
        return True
    text = f"{row.subject} {row.sender}".lower()
    aliases = {
        "openai": ("openai", "chatgpt"),
        "chatgpt": ("openai", "chatgpt"),
        "microsoft": ("microsoft", "outlook", "live.com"),
    }.get(provider, (provider,))
    return any(alias in text for alias in aliases)


def row_age_seconds(row: MailRow) -> int | None:
    ts = mail_date_ts(row.date)
    if ts is None:
        return None
    return max(0, int(time.time() - ts))


def row_received_iso(row: MailRow) -> str:
    ts = mail_date_ts(row.date)
    if ts is None:
        return ""
    return datetime.fromtimestamp(ts, tz=timezone.utc).isoformat()


def latest_code_response(result: SmsFetchResult, recent_seconds: int = 0, provider: str = "any") -> dict:
    code_row = next(
        (
            row
            for row in result.rows
            if row.code and mail_row_within(row, recent_seconds) and provider_matches(row, provider)
        ),
        None,
    )
    latest_row = next(
        (
            row
            for row in result.rows
            if mail_row_within(row, recent_seconds) and provider_matches(row, provider)
        ),
        None,
    )
    if latest_row is None:
        latest_row = result.rows[0] if result.rows else None
    picked = code_row or latest_row
    return {
        "ok": True,
        "account": result.account,
        "code": code_row.code if code_row else "",
        "found": bool(code_row),
        "recent_seconds": recent_seconds,
        "provider": provider or "any",
        "age_seconds": row_age_seconds(picked) if picked else None,
        "received_at": row_received_iso(picked) if picked else "",
        "msg_id": picked.msg_id if picked else "",
        "date": picked.date if picked else "",
        "sender": picked.sender if picked else "",
        "subject": picked.subject if picked else "",
        "token_source": result.token_source,
        "server": result.server,
        "used_folder": result.used_folder,
        "top": result.top,
        "attempts": result.attempts,
    }


def api_code_response(result: SmsFetchResult, recent_seconds: int, provider: str) -> dict:
    payload = latest_code_response(result, recent_seconds=recent_seconds, provider=provider)
    if not payload.get("found"):
        return {
            "ok": False,
            "reason": "not_found",
            "message": f"{recent_seconds}秒内未找到验证码" if recent_seconds > 0 else "未找到验证码",
            "provider": provider or "any",
            "email": result.account,
            "recent_seconds": recent_seconds,
            "checked": len(result.rows),
            "latest_subject": payload.get("subject") or "",
            "latest_date": payload.get("date") or "",
        }
    return {
        "ok": True,
        "code": payload["code"],
        "provider": provider or "any",
        "email": result.account,
        "received_at": payload.get("received_at") or "",
        "age_seconds": payload.get("age_seconds"),
        "subject": payload.get("subject") or "",
        "sender": payload.get("sender") or "",
        "msg_id": payload.get("msg_id") or "",
        "recent_seconds": recent_seconds,
    }


def latest_code_text(payload: dict) -> str:
    if not payload.get("ok"):
        return "NO|读取失败"
    code = str(payload.get("code") or "").strip()
    if not code:
        recent_seconds = int(payload.get("recent_seconds") or 0)
        if recent_seconds > 0:
            return f"NO|{recent_seconds}秒内未找到验证码"
        return "NO|未找到验证码"
    subject = str(payload.get("subject") or "").lower()
    sender = str(payload.get("sender") or "").lower()
    service = "OpenAI" if ("openai" in subject or "openai" in sender) else "邮箱"
    return f"YES|您的 {service} 验证代码是: {code}"


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
    .mono{{font-family:ui-monospace,SFMono-Regular,Consolas,monospace}}
    table{{width:100%;border-collapse:collapse;margin-top:10px}}
    th,td{{border:1px solid #ddd;padding:8px;font-size:14px;vertical-align:top}}
    th{{background:#f0f2f4}}
    .grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:8px}}
    .small{{font-size:12px;color:#666}}
    .account-table th:first-child,.account-table td:first-child{{width:110px;text-align:center}}
    .account-line{{box-sizing:border-box;min-height:76px;font-size:12px;line-height:1.45;font-family:ui-monospace,SFMono-Regular,Consolas,monospace}}
  </style>
</head>
<body>
  <div class="wrap">{body_html}</div>
</body>
</html>"""


def stored_account_table_html(search: str = "") -> str:
    search = (search or "").strip()
    records = load_account_records()
    row_html = ""

    for record in records:
        account_line = record.get("account_line", "").strip()
        searchable = " ".join(
            [
                account_line,
                record.get("email", ""),
                record.get("client_id", ""),
                record.get("token_tail", ""),
                record.get("updated_at", ""),
            ]
        ).lower()
        if account_line and (not search or search.lower() in searchable):
            row_html += f"""
            <tr>
              <td>
                <form method="post" action="/sms/fetch" target="_blank">
                  <input type="hidden" name="account_line" value="{html.escape(account_line)}">
                  <input type="hidden" name="top" value="20">
                  <button class="btn" type="submit">接码</button>
                </form>
              </td>
              <td>
                <textarea class="account-line" readonly onclick="this.select()">{html.escape(account_line)}</textarea>
              </td>
            </tr>
            """

    if not row_html:
        message = "没有匹配的账号信息" if search else "暂无存储账号信息"
        row_html = f'<tr><td colspan="2" class="small">{message}</td></tr>'

    return f"""
    <div class="card">
      <h3 style="margin-top:0">存储账号信息</h3>
      <form method="post" action="/admin/accounts/add" style="margin-bottom:12px">
        <textarea name="account_lines" placeholder="粘贴账号信息，一行一个：email----password----client_id----token" required></textarea>
        <div class="row" style="margin-top:10px">
          <button class="btn green" type="submit">添加账号</button>
        </div>
      </form>
      <form method="get" action="/admin" class="row" style="margin-bottom:10px">
        <input type="text" name="account_q" placeholder="搜索账号信息" value="{html.escape(search)}">
        <button class="btn" type="submit">搜索</button>
        <a class="btn gray" href="/admin">清空</a>
      </form>
      <table class="account-table">
        <thead>
          <tr>
            <th>接码</th>
            <th>账号信息</th>
          </tr>
        </thead>
        <tbody>{row_html}</tbody>
      </table>
    </div>
    """


def sms_form_html(
    account_line: str = "",
    top: str = "20",
) -> str:
    return f"""
    <div class="card">
      <h1>\u8f7b\u91cf Outlook \u90ae\u7bb1\u67e5\u770b</h1>
      <form method="post" action="/sms/fetch">
        <label>\u8d26\u53f7\u884c\uff08email----password----client_id----token\uff09</label>
        <textarea name="account_line" placeholder="\u7c98\u8d34\u5b8c\u6574\u8d26\u53f7\u884c" required>{html.escape(account_line)}</textarea>
        <div class="grid" style="margin-top:10px;max-width:280px">
          <label>\u8bfb\u53d6\u6570\u91cf
            <input type="number" name="top" min="1" max="200" value="{html.escape(top)}">
          </label>
        </div>
        <div class="row" style="margin-top:10px">
          <button class="btn" type="submit">\u5f00\u59cb\u8bfb\u53d6</button>
          <a class="btn gray" href="/">\u8fd4\u56de\u9996\u9875</a>
        </div>
      </form>
      <p class="note">\u9ed8\u8ba4\u4f18\u5148\u8bfb\u53d6 INBOX\uff0c\u5e76\u81ea\u52a8\u5c1d\u8bd5 token \u5019\u9009\uff08\u7f13\u5b58/\u539f\u59cb/\u5237\u65b0\uff09\u3002</p>
    </div>
    <div class="card">
      <h3 style="margin-top:0">\u5f00\u6e90\u9879\u76ee</h3>
      <p class="small" style="font-size:14px;color:#444;line-height:1.8">
        \u6e90\u7801\u6765\u81ea\uff1a
        <a href="https://github.com/boji1334/outlook-reward-sms-lightweight" target="_blank" rel="noopener noreferrer">
          https://github.com/boji1334/outlook-reward-sms-lightweight
        </a>
      </p>
      <p class="small" style="font-size:14px;color:#444;line-height:1.8">
        \u5982\u679c\u8fd9\u4e2a\u5de5\u5177\u5bf9\u4f60\u6709\u5e2e\u52a9\uff0c\u6b22\u8fce\u5728 GitHub \u70b9\u4e2a Star \u5e76\u5173\u6ce8\uff0c\u652f\u6301\u9879\u76ee\u6301\u7eed\u66f4\u65b0\u3002
      </p>
    </div>
    """


def code_page_html() -> str:
    return """
    <div class="card" style="max-width:760px;margin:32px auto">
      <h1>验证码读取</h1>
      <form id="codeForm">
        <label>账号信息</label>
        <textarea id="accountLine" name="account_line" placeholder="email----password----client_id----refresh_token" required></textarea>
        <div class="grid" style="margin-top:10px">
          <label>平台
            <select id="provider" name="provider">
              <option value="openai" selected>OpenAI</option>
              <option value="any">不限</option>
            </select>
          </label>
          <label>时间窗口
            <input id="withinSeconds" type="number" name="within_seconds" min="1" max="300" value="30">
          </label>
          <label>读取数量
            <input id="top" type="number" name="top" min="1" max="200" value="20">
          </label>
        </div>
        <div class="row" style="margin-top:12px">
          <button id="codeButton" class="btn green" type="submit">获取验证码</button>
          <button id="clearButton" class="btn gray" type="button">清空</button>
        </div>
      </form>
      <div id="codeResult" class="code-result" aria-live="polite">等待输入账号信息</div>
    </div>
    <style>
      .code-result{margin-top:16px;border:1px solid #ddd;border-radius:8px;padding:16px;background:#fbfcff;min-height:74px;font-size:16px;line-height:1.6}
      .code-result.ok{border-color:#b7e4c7;background:#f2fbf6;color:#11623a}
      .code-result.err{border-color:#f2b8b5;background:#fff6f5;color:#b42318}
      .code-big{display:block;font-size:34px;font-weight:800;line-height:1.25;letter-spacing:0;margin-top:4px;color:#111}
      @media (max-width:640px){.code-big{font-size:28px}.card{margin-top:12px!important}}
    </style>
    <script>
      const form = document.querySelector('#codeForm');
      const accountLine = document.querySelector('#accountLine');
      const provider = document.querySelector('#provider');
      const withinSeconds = document.querySelector('#withinSeconds');
      const topInput = document.querySelector('#top');
      const button = document.querySelector('#codeButton');
      const clearButton = document.querySelector('#clearButton');
      const result = document.querySelector('#codeResult');

      function setResult(text, kind) {
        result.className = 'code-result' + (kind ? ' ' + kind : '');
        result.innerHTML = text;
      }

      function escapeHtml(value) {
        return String(value || '').replace(/[&<>"']/g, ch => ({
          '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
        }[ch]));
      }

      form.addEventListener('submit', async (event) => {
        event.preventDefault();
        const raw = accountLine.value.trim();
        if (!raw) {
          setResult('请先粘贴账号信息', 'err');
          return;
        }
        button.disabled = true;
        setResult('正在读取最新邮件...', '');
        try {
          const response = await fetch('/api/v1/code', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({
              account_line: raw,
              provider: provider.value,
              within_seconds: Number(withinSeconds.value || 30),
              top: Number(topInput.value || 20)
            })
          });
          const data = await response.json();
          if (data.ok) {
            const age = Number.isFinite(Number(data.age_seconds)) ? `${data.age_seconds} 秒前` : '';
            setResult(
              `${escapeHtml(data.provider || '验证码')} 验证码<span class="code-big">${escapeHtml(data.code)}</span>` +
              `<span class="small">${escapeHtml(age)} ${escapeHtml(data.subject || '')}</span>`,
              'ok'
            );
          } else {
            setResult(escapeHtml(data.message || '未找到验证码'), 'err');
          }
        } catch (error) {
          setResult('请求失败，请稍后重试', 'err');
        } finally {
          button.disabled = false;
        }
      });

      clearButton.addEventListener('click', () => {
        accountLine.value = '';
        setResult('等待输入账号信息', '');
        accountLine.focus();
      });
    </script>
    """


def sms_result_html(result: SmsFetchResult) -> str:
    rows_html = ""
    for i, row in enumerate(result.rows, start=1):
        code_cell = html.escape(row.code) if row.code else "-"
        rows_html += (
            "<tr>"
            f"<td>{i}</td>"
            f"<td>{html.escape(row.date)}</td>"
            f"<td>{html.escape(row.sender)}</td>"
            f"<td>{html.escape(row.subject)}</td>"
            f"<td><b>{code_cell}</b></td>"
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
        rows_html = '<tr><td colspan="6" class="small">没有读取到邮件</td></tr>'

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
        <thead><tr><th style="width:60px">编号</th><th style="width:250px">日期</th><th>发件人</th><th>主题</th><th style="width:100px">验证码</th><th style="width:130px">操作</th></tr></thead>
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

    def _send_json(self, payload: dict, code: int = 200):
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        if not getattr(self, "_head_only", False):
            self.wfile.write(data)

    def _send_text(self, text: str, code: int = 200):
        data = text.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
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
        ctype = (self.headers.get("Content-Type", "") or "").lower()
        if "application/json" in ctype:
            try:
                data = json.loads(raw or "{}")
                if isinstance(data, dict):
                    return {str(k): str(v).strip() for k, v in data.items() if v is not None}
            except Exception:
                return {}
        parsed = urllib.parse.parse_qs(raw)
        result: dict[str, str] = {}
        for k, v in parsed.items():
            result[k] = (v[0] if v else "").strip()
        return result

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Max-Age", "86400")
        self.end_headers()

    def do_HEAD(self):
        self._head_only = True
        try:
            self.do_GET()
        finally:
            self._head_only = False

    def _send_latest_code_result(self, fields: dict[str, str], wants_text: bool = False, relay_text: bool = False):
        account_line = fields.get("account_line") or fields.get("account") or ""
        top = fields.get("top", "20")
        try:
            top_i = int(top)
        except Exception:
            top_i = 20
        top_i = min(max(top_i, 1), 200)
        recent_raw = fields.get("recent_seconds") or fields.get("within_seconds") or fields.get("within")
        if recent_raw is None and relay_text:
            recent_raw = "30"
        try:
            recent_seconds = int(recent_raw or "0")
        except Exception:
            recent_seconds = 0
        recent_seconds = min(max(recent_seconds, 0), 3600)
        provider = fields.get("provider") or "any"

        try:
            result = perform_sms_fetch(
                account_line=account_line,
                top=top_i,
                folder=fields.get("folder") or "INBOX",
                token_kind=fields.get("token_kind") or "auto",
                refresh_endpoint=fields.get("refresh_endpoint") or "auto",
                tenant=fields.get("tenant") or "consumers",
                timeout=SMS_FETCH_TIMEOUT,
            )
            payload = latest_code_response(result, recent_seconds=recent_seconds, provider=provider)
            if relay_text:
                self._send_text(latest_code_text(payload))
            elif wants_text:
                self._send_text(str(payload.get("code") or ""))
            else:
                self._send_json(payload)
        except Exception as exc:
            if relay_text:
                self._send_text(f"NO|{exc}", code=400)
            elif wants_text:
                self._send_text(str(exc), code=400)
            else:
                self._send_json({"ok": False, "error": str(exc)}, code=400)

    def _send_api_v1_code_result(self, fields: dict[str, str]):
        account_line = fields.get("account_line") or fields.get("account") or ""
        provider = fields.get("provider") or "openai"
        top = fields.get("top", "20")
        within = fields.get("within_seconds") or fields.get("recent_seconds") or fields.get("within") or "30"
        try:
            top_i = int(top)
        except Exception:
            top_i = 20
        try:
            within_i = int(within)
        except Exception:
            within_i = 30
        top_i = min(max(top_i, 1), 200)
        within_i = min(max(within_i, 1), 300)

        try:
            result = perform_sms_fetch(
                account_line=account_line,
                top=top_i,
                folder=fields.get("folder") or "INBOX",
                token_kind=fields.get("token_kind") or "auto",
                refresh_endpoint=fields.get("refresh_endpoint") or "auto",
                tenant=fields.get("tenant") or "consumers",
                timeout=SMS_FETCH_TIMEOUT,
            )
            self._send_json(api_code_response(result, recent_seconds=within_i, provider=provider))
        except ValueError as exc:
            self._send_json({"ok": False, "reason": "bad_account", "message": str(exc)}, code=400)
        except Exception as exc:
            message = str(exc)
            reason = "auth_failed" if is_auth_or_token_error(message) else "mail_timeout" if is_network_error(message) else "upstream_error"
            self._send_json({"ok": False, "reason": reason, "message": message}, code=400)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        qs = urllib.parse.parse_qs(parsed.query)

        if path in ("/api/sms/code", "/api/sms/latest", "/api/sms/code.txt") or path.startswith(
            "/api/sms/code/"
        ):
            fields: dict[str, str] = {}
            for key, values in qs.items():
                fields[key] = (values[0] if values else "").strip()
            if path.startswith("/api/sms/code/"):
                fields.setdefault(
                    "account_line",
                    urllib.parse.unquote(path[len("/api/sms/code/") :]).strip(),
                )
            fmt = (fields.get("format") or "").lower()
            wants_text = path == "/api/sms/code.txt" or fmt in ("text", "txt", "plain") or fields.get("plain") == "1"
            self._send_latest_code_result(fields, wants_text=wants_text)
            return

        if path == "/api/text-relay" or path.startswith("/api/text-relay/"):
            fields: dict[str, str] = {}
            for key, values in qs.items():
                fields[key] = (values[0] if values else "").strip()
            if path.startswith("/api/text-relay/"):
                fields.setdefault(
                    "account_line",
                    urllib.parse.unquote(path[len("/api/text-relay/") :]).strip(),
                )
            self._send_latest_code_result(fields, relay_text=True)
            return

        if path == "/":
            body = """
            <div class="card">
              <h1>功能选择</h1>
              <div class="row">
                <a class="btn" href="/reward">打赏</a>
                <a class="btn green" href="/sms">验证码</a>
                <a class="btn gray" href="/sms/list">邮件列表</a>
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

        if path in ("/sms", "/code"):
            self._send_html(render_layout("验证码读取", code_page_html()))
            return

        if path == "/sms/list":
            self._send_html(render_layout("邮件列表", sms_form_html()))
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
            account_search = (qs.get("account_q", [""])[0] or "").strip()
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
            {stored_account_table_html(account_search)}
            """
            self._send_html(render_layout("管理员面板", body))
            return

        self._send_html(render_layout("404", '<div class="card"><h1>404</h1></div>'), code=404)

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        qs = urllib.parse.parse_qs(parsed.query)

        if path == "/api/v1/code":
            self._send_api_v1_code_result(self._parse_urlencoded())
            return

        if path in ("/api/sms/code", "/api/sms/latest", "/api/sms/code.txt"):
            form = self._parse_urlencoded()
            fmt = (form.get("format") or (qs.get("format", [""])[0] if qs else "")).lower()
            wants_text = path == "/api/sms/code.txt" or fmt in ("text", "txt", "plain") or form.get("plain") == "1"
            self._send_latest_code_result(form, wants_text=wants_text)
            return

        if path == "/api/text-relay":
            form = self._parse_urlencoded()
            self._send_latest_code_result(form, relay_text=True)
            return

        if path == "/api/sms/fetch":
            form = self._parse_urlencoded()
            account_line = form.get("account_line", "")
            top = form.get("top", "20")
            try:
                top_i = int(top)
            except Exception:
                top_i = 20
            top_i = min(max(top_i, 1), 200)

            try:
                result = perform_sms_fetch(
                    account_line=account_line,
                    top=top_i,
                    folder="INBOX",
                    token_kind="auto",
                    refresh_endpoint="auto",
                    tenant="consumers",
                    timeout=SMS_FETCH_TIMEOUT,
                )
                self._send_json(
                    {
                        "ok": True,
                        "account": result.account,
                        "token_source": result.token_source,
                        "server": result.server,
                        "used_folder": result.used_folder,
                        "top": result.top,
                        "attempts": result.attempts,
                        "rows": [
                            {
                                "msg_id": row.msg_id,
                                "date": row.date,
                                "sender": row.sender,
                                "subject": row.subject,
                                "code": row.code,
                            }
                            for row in result.rows
                        ],
                    }
                )
            except Exception as exc:
                self._send_json({"ok": False, "error": str(exc)}, code=400)
            return

        if path == "/api/sms/view":
            form = self._parse_urlencoded()
            account_line = form.get("account_line", "")
            msg_id = form.get("msg_id", "")
            folder = form.get("folder", "INBOX") or "INBOX"
            try:
                detail = perform_sms_view(
                    account_line=account_line,
                    msg_id=msg_id,
                    preferred_folder=folder,
                    timeout=SMS_VIEW_TIMEOUT,
                )
                self._send_json(
                    {
                        "ok": True,
                        "account": detail.account,
                        "token_source": detail.token_source,
                        "server": detail.server,
                        "used_folder": detail.used_folder,
                        "msg_id": detail.msg_id,
                        "subject": detail.subject,
                        "sender": detail.sender,
                        "receiver": detail.receiver,
                        "date": detail.date,
                        "body_text": detail.body_text,
                        "attempts": detail.attempts,
                    }
                )
            except Exception as exc:
                self._send_json({"ok": False, "error": str(exc)}, code=400)
            return

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
                    timeout=SMS_FETCH_TIMEOUT,
                )
                save_account_record(account_line)
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
                    timeout=SMS_VIEW_TIMEOUT,
                )
                parts.append(sms_detail_html(detail, top_i))
            except Exception as exc:
                parts.append(sms_error_html(str(exc)))

            self._send_html(render_layout("邮件内容", "".join(parts)))
            return

        if path == "/admin/accounts/add":
            if not self._is_admin():
                self._send_html(
                    render_layout(
                        "请先登录",
                        '<div class="card"><p>请先登录。</p><a class="btn" href="/admin">去登录</a></div>',
                    ),
                    code=401,
                )
                return
            form = self._parse_urlencoded()
            added, errors = add_account_lines(form.get("account_lines", ""))
            if errors and not added:
                error_html = "<br>".join(html.escape(x) for x in errors)
                body = f"""
                <div class="card">
                  <h1>添加失败</h1>
                  <p class="err">{error_html}</p>
                  <a class="btn gray" href="/admin">返回管理</a>
                </div>
                """
                self._send_html(render_layout("添加失败", body), code=400)
                return
            self.send_response(303)
            self.send_header("Location", "/admin")
            self.end_headers()
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
