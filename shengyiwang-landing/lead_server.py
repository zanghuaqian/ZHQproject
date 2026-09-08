#!/usr/bin/env python3
"""盛意旺落地页留资服务：静态页 + 活码池 + 企微回调。

用法:
  python lead_server.py
  python lead_server.py --port 8088

未配置企微凭证时自动进入 mock 模式（本地可走通表单→弹码→模拟扫码）。
"""
from __future__ import annotations

import argparse
import json
import os
import random
import re
import sqlite3
import string
import threading
import time
import traceback
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from wecom_crypto import WecomCrypto, WecomCryptoError

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
DB_PATH = DATA_DIR / "leads.db"
ENV_PATH = ROOT / ".env"
TZ = timezone(timedelta(hours=8))

PHONE_RE = re.compile(r"^1[3-9]\d{9}$")
LEAD_ID_RE = re.compile(r"^[a-zA-Z0-9]{6,30}$")
PURPOSE_LABELS = {
    "pay": "开通收款 / 商户进件",
    "hw": "采购智能收款硬件",
    "saas": "门店数字化 / 对账系统",
    "agent": "意向加盟 / 代理合作",
    "api": "开放平台 API 对接",
}
PERSONAS = {"merchant", "partner"}

# 本地/联调默认小池，避免一次打满配额；生产在 .env 调大
POOL_MIN = 3
POOL_TARGET = 5
QR_REFRESH_DAYS = 6
LEAD_EXPIRE_DAYS = 3

_db_lock = threading.RLock()
_token_lock = threading.Lock()
_token_cache = {"token": "", "expire_at": 0.0}
_rate = {}  # ip -> [timestamps]
DISPLAY_QR = "/qr/wecom-livecode.png"
WECOM_QR_LINK = "https://work.weixin.qq.com/ca/cawcdec3641618070f"

_captchas: dict[str, dict] = {}
_sms_codes: dict[str, dict] = {}
_verify_lock = threading.Lock()
_industries_cache: list[dict] | None = None


def now() -> datetime:
    return datetime.now(TZ)


def now_iso() -> str:
    return now().strftime("%Y-%m-%d %H:%M:%S")


def load_env(path: Path) -> None:
    if not path.exists():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def env(name: str, default: str = "") -> str:
    return (os.environ.get(name) or default).strip()


def env_int(name: str, default: int) -> int:
    raw = env(name)
    if not raw:
        return default
    try:
        return int(raw)
    except ValueError:
        return default


def is_mock() -> bool:
    if env("LEAD_MOCK") in {"1", "true", "True", "yes"}:
        return True
    if env("LEAD_MOCK") in {"0", "false", "False", "no"}:
        return False
    return not (env("WECOM_CORP_ID") and env("WECOM_SECRET") and env("WECOM_SALES_USERID"))


def phone_key() -> bytes:
    secret = env("LEAD_PHONE_KEY") or env("WECOM_SECRET") or "syw-local-dev-key"
    import hashlib

    return hashlib.sha256(secret.encode("utf-8")).digest()


def encrypt_phone(phone: str) -> str:
    key = phone_key()
    xored = bytes(b ^ key[i % len(key)] for i, b in enumerate(phone.encode("utf-8")))
    import base64

    return base64.urlsafe_b64encode(xored).decode("ascii")


def decrypt_phone(blob: str) -> str:
    if not blob:
        return ""
    import base64

    try:
        raw = base64.urlsafe_b64decode(blob.encode("ascii"))
    except Exception:
        return blob
    key = phone_key()
    return bytes(b ^ key[i % len(key)] for i, b in enumerate(raw)).decode("utf-8", errors="replace")


def mask_phone(phone: str) -> str:
    if len(phone) == 11:
        return phone[:3] + "****" + phone[7:]
    return phone


def short_id(n: int = 10) -> str:
    alphabet = string.ascii_letters + string.digits
    return "".join(random.choice(alphabet) for _ in range(n))


def connect() -> sqlite3.Connection:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS qr_codes (
            code_id TEXT PRIMARY KEY,
            config_id TEXT NOT NULL DEFAULT '',
            qr_code TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'idle',
            created_at TEXT NOT NULL,
            refreshed_at TEXT NOT NULL,
            assigned_at TEXT
        );
        CREATE TABLE IF NOT EXISTS leads (
            id TEXT PRIMARY KEY,
            company TEXT NOT NULL,
            phone_enc TEXT NOT NULL,
            purpose TEXT NOT NULL,
            persona TEXT NOT NULL DEFAULT 'merchant',
            status TEXT NOT NULL,
            submitted_at TEXT NOT NULL,
            expire_at TEXT NOT NULL,
            external_userid TEXT,
            sales_userid TEXT,
            connected_at TEXT,
            consent INTEGER NOT NULL DEFAULT 1,
            mock INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT NOT NULL,
            action TEXT NOT NULL,
            lead_id TEXT,
            detail TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_qr_idle ON qr_codes(status, created_at);
        CREATE INDEX IF NOT EXISTS idx_leads_status ON leads(status);
        """
    )
    cols = {row[1] for row in conn.execute("PRAGMA table_info(leads)")}
    extras = {
        "contact_name": "TEXT NOT NULL DEFAULT ''",
        "province": "TEXT NOT NULL DEFAULT ''",
        "city": "TEXT NOT NULL DEFAULT ''",
        "industry_l1": "TEXT NOT NULL DEFAULT ''",
        "industry_name": "TEXT NOT NULL DEFAULT ''",
    }
    for name, spec in extras.items():
        if name not in cols:
            conn.execute(f"ALTER TABLE leads ADD COLUMN {name} {spec}")
    conn.commit()


def audit(conn: sqlite3.Connection, action: str, lead_id: str = "", detail: str = "") -> None:
    conn.execute(
        "INSERT INTO audit_log(ts, action, lead_id, detail) VALUES (?,?,?,?)",
        (now_iso(), action, lead_id, detail[:2000]),
    )


def wx_http(path: str, body: dict | None = None, token: str | None = None) -> dict:
    q = f"?access_token={urllib.parse.quote(token)}" if token else ""
    url = f"https://qyapi.weixin.qq.com/cgi-bin/{path}{q}"
    data = None if body is None else json.dumps(body, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="GET" if body is None else "POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=12) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        raise RuntimeError(f"企微 HTTP {exc.code}: {exc.read()[:300]!r}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"企微网络错误: {exc}") from exc
    return payload


def get_token(force: bool = False) -> str:
    if is_mock():
        return "mock-token"
    with _token_lock:
        if not force and _token_cache["token"] and _token_cache["expire_at"] > time.time() + 120:
            return _token_cache["token"]
        corp = env("WECOM_CORP_ID")
        secret = env("WECOM_SECRET")
        path = f"gettoken?corpid={urllib.parse.quote(corp)}&corpsecret={urllib.parse.quote(secret)}"
        payload = wx_http(path)
        if payload.get("errcode"):
            raise RuntimeError(f"get_access_token 失败: {payload}")
        _token_cache["token"] = payload["access_token"]
        _token_cache["expire_at"] = time.time() + int(payload.get("expires_in", 7200))
        return _token_cache["token"]


def wx_api(path: str, body: dict) -> dict:
    token = get_token()
    payload = wx_http(path, body, token)
    if payload.get("errcode") in (40014, 42001, 40001):
        token = get_token(force=True)
        payload = wx_http(path, body, token)
    if payload.get("errcode"):
        raise RuntimeError(f"{path} 失败: {payload}")
    return payload


def add_contact_way(code_id: str) -> dict:
    if is_mock():
        return {
            "config_id": f"mock-{code_id}",
            "qr_code": "/qr/wecom-cs.png",
        }
    return wx_api(
        "externalcontact/add_contact_way",
        {
            "type": 1,
            "scene": 2,
            "state": code_id,
            "user": [env("WECOM_SALES_USERID")],
            "skip_verify": True,
            "is_temp": False,
        },
    )


def refresh_qr_url(config_id: str, fallback: str) -> str:
    if is_mock() or not config_id or config_id.startswith("mock-"):
        return fallback
    try:
        payload = wx_api("externalcontact/get_contact_way", {"config_id": config_id})
        contact = payload.get("contact_way") or payload
        return contact.get("qr_code") or fallback
    except Exception as exc:
        print(f"[warn] 刷新二维码 URL 失败: {exc}")
        return fallback


def replenish_pool(conn: sqlite3.Connection, target: int | None = None) -> None:
    want = target if target is not None else env_int("CODE_POOL_TARGET", POOL_TARGET)
    idle = conn.execute("SELECT COUNT(*) FROM qr_codes WHERE status='idle'").fetchone()[0]
    if idle >= env_int("CODE_POOL_MIN", POOL_MIN):
        return
    need = max(0, want - idle)
    for _ in range(need):
        code_id = short_id()
        try:
            resp = add_contact_way(code_id)
        except Exception as exc:
            print(f"[warn] 预生成活码失败: {exc}")
            if is_mock():
                resp = {"config_id": f"mock-{code_id}", "qr_code": "/qr/wecom-cs.png"}
            else:
                break
        ts = now_iso()
        conn.execute(
            "INSERT INTO qr_codes(code_id, config_id, qr_code, status, created_at, refreshed_at) "
            "VALUES (?,?,?,?,?,?)",
            (code_id, resp.get("config_id") or "", resp.get("qr_code") or "/qr/wecom-cs.png", "idle", ts, ts),
        )
        audit(conn, "pool_add", code_id, resp.get("config_id") or "")
    conn.commit()


def pop_idle_qr(conn: sqlite3.Connection) -> dict | None:
    row = conn.execute(
        "SELECT * FROM qr_codes WHERE status='idle' ORDER BY created_at ASC LIMIT 1"
    ).fetchone()
    if not row:
        return None
    item = dict(row)
    refreshed = item["qr_code"]
    try:
        age = now() - datetime.strptime(item["refreshed_at"], "%Y-%m-%d %H:%M:%S").replace(tzinfo=TZ)
        if age.days >= QR_REFRESH_DAYS:
            refreshed = refresh_qr_url(item["config_id"], item["qr_code"])
            conn.execute(
                "UPDATE qr_codes SET qr_code=?, refreshed_at=? WHERE code_id=?",
                (refreshed, now_iso(), item["code_id"]),
            )
    except Exception:
        pass
    conn.execute(
        "UPDATE qr_codes SET status='assigned', assigned_at=? WHERE code_id=?",
        (now_iso(), item["code_id"]),
    )
    item["qr_code"] = refreshed
    item["status"] = "assigned"
    return item


def expire_old_leads(conn: sqlite3.Connection) -> None:
    conn.execute(
        "UPDATE leads SET status='expired' WHERE status='submitted' AND expire_at < ?",
        (now_iso(),),
    )


def purpose_label(purpose: str) -> str:
    return PURPOSE_LABELS.get(purpose, purpose or "未填写")


def tag_ids_for(purpose: str) -> list[str]:
    tags = []
    lead_tag = env("WECOM_TAG_LEAD")
    if lead_tag:
        tags.append(lead_tag)
    purpose_tag = env(f"WECOM_TAG_{purpose.upper()}")
    if purpose_tag:
        tags.append(purpose_tag)
    return tags


def notify_sales(userid: str, lead: dict, phone: str) -> None:
    agent_id = env("WECOM_AGENT_ID")
    if is_mock() or not agent_id:
        return
    desc = (
        f"企业：{lead['company']}\n"
        f"手机：{phone}\n"
        f"意向：{purpose_label(lead['purpose'])}\n"
        f"角色：{'服务商' if lead.get('persona') == 'partner' else '商户'}\n"
        f"时间：{lead['submitted_at']}"
    )
    wx_api(
        "message/send",
        {
            "touser": userid,
            "msgtype": "textcard",
            "agentid": int(agent_id),
            "textcard": {
                "title": "官网新留资",
                "description": desc,
                "url": env("LEAD_NOTIFY_URL") or "https://work.weixin.qq.com",
                "btntxt": "查看",
            },
        },
    )


def after_connected(lead: dict, phone: str, welcome_code: str, sales_userid: str, external_userid: str) -> None:
    errors = []
    if welcome_code and not is_mock():
        try:
            wx_api(
                "externalcontact/send_welcome_msg",
                {
                    "welcome_code": welcome_code,
                    "text": {
                        "content": (
                            f"您好，已收到您在「{lead['company']}」的合作咨询，"
                            f"专属顾问马上为您服务～"
                        )
                    },
                },
            )
        except Exception as exc:
            errors.append(f"welcome: {exc}")
    if not is_mock():
        try:
            wx_api(
                "externalcontact/remark",
                {
                    "userid": sales_userid,
                    "external_userid": external_userid,
                    "remark": f"[官网留资] {lead['company']}",
                    "description": (
                        f"企业名称：{lead['company']}\n"
                        f"联系手机：{phone}\n"
                        f"合作目的：{purpose_label(lead['purpose'])}\n"
                        f"留资时间：{lead['submitted_at']}\n"
                        f"来源渠道：盛意旺官网留资区"
                    ),
                    "remark_mobiles": [phone],
                },
            )
        except Exception as exc:
            errors.append(f"remark: {exc}")
        tags = tag_ids_for(lead["purpose"])
        if tags:
            try:
                wx_api(
                    "externalcontact/mark_tag",
                    {
                        "userid": sales_userid,
                        "external_userid": external_userid,
                        "add_tag": tags,
                    },
                )
            except Exception as exc:
                errors.append(f"tag: {exc}")
        try:
            notify_sales(sales_userid, lead, phone)
        except Exception as exc:
            errors.append(f"notify: {exc}")
    if errors:
        print("[warn] 加好友后置任务部分失败:", "; ".join(errors))


def connect_lead(conn: sqlite3.Connection, state: str, external_userid: str, sales_userid: str, welcome_code: str) -> None:
    row = conn.execute("SELECT * FROM leads WHERE id=?", (state,)).fetchone()
    if not row:
        audit(conn, "callback_unknown_state", state, external_userid)
        conn.commit()
        print(f"[warn] 未知 State 回调: {state}")
        return
    lead = dict(row)
    if lead["status"] == "connected":
        audit(conn, "callback_duplicate", state, external_userid)
        conn.commit()
        return
    conn.execute(
        "UPDATE leads SET status='connected', external_userid=?, sales_userid=?, connected_at=? WHERE id=?",
        (external_userid, sales_userid, now_iso(), state),
    )
    audit(conn, "connected", state, f"{sales_userid}->{external_userid}")
    conn.commit()
    phone = decrypt_phone(lead["phone_enc"])
    threading.Thread(
        target=after_connected,
        args=(lead, phone, welcome_code, sales_userid, external_userid),
        daemon=True,
    ).start()


def handle_wx_event(msg: dict, conn: sqlite3.Connection) -> None:
    event = msg.get("Event") or ""
    change = msg.get("ChangeType") or ""
    if event != "change_external_contact":
        return
    if change == "add_external_contact":
        connect_lead(
            conn,
            msg.get("State") or "",
            msg.get("ExternalUserID") or "",
            msg.get("UserID") or "",
            msg.get("WelcomeCode") or "",
        )
        return
    if change in {"del_follow_user", "del_external_contact"}:
        ext = msg.get("ExternalUserID") or ""
        if ext:
            conn.execute(
                "UPDATE leads SET status='lost' WHERE external_userid=? AND status='connected'",
                (ext,),
            )
            audit(conn, "lost", "", ext)
            conn.commit()


def create_lead(payload: dict, conn: sqlite3.Connection) -> dict:
    company = str(payload.get("company") or "").strip()
    contact_name = str(payload.get("contact_name") or "").strip()
    phone = str(payload.get("phone") or "").strip()
    sms_code = str(payload.get("sms_code") or "").strip()
    purpose = str(payload.get("purpose") or "").strip()
    persona = str(payload.get("persona") or "merchant").strip()
    province = str(payload.get("province") or "").strip()
    city = str(payload.get("city") or "").strip()
    industry_l1 = str(payload.get("industry_l1") or "").strip()
    industry_name = str(payload.get("industry_name") or "").strip()
    consent = bool(payload.get("consent", True))
    if not company:
        raise ValueError("请填写企业名称")
    if len(company) > 80:
        raise ValueError("企业名称过长")
    if not contact_name:
        raise ValueError("请填写联系人姓名")
    if len(contact_name) > 20:
        raise ValueError("联系人姓名过长")
    if not PHONE_RE.match(phone):
        raise ValueError("请输入有效的 11 位手机号")
    consume_sms(phone, sms_code)
    if not valid_region(province, city):
        raise ValueError("请选择所在省市")
    if not valid_industry(industry_l1, industry_name):
        raise ValueError("请选择所属行业")
    if purpose not in PURPOSE_LABELS:
        raise ValueError("请选择合作目的")
    if persona not in PERSONAS:
        persona = "merchant"
    if not consent:
        raise ValueError("请勾选同意后再提交")

    expire_old_leads(conn)
    replenish_pool(conn)
    qr = pop_idle_qr(conn)
    if qr is None:
        code_id = short_id()
        resp = add_contact_way(code_id)
        ts = now_iso()
        conn.execute(
            "INSERT INTO qr_codes(code_id, config_id, qr_code, status, created_at, refreshed_at, assigned_at) "
            "VALUES (?,?,?,?,?,?,?)",
            (
                code_id,
                resp.get("config_id") or "",
                resp.get("qr_code") or DISPLAY_QR,
                "assigned",
                ts,
                ts,
                ts,
            ),
        )
        qr = {"code_id": code_id, "qr_code": resp.get("qr_code") or DISPLAY_QR}

    expire_at = (now() + timedelta(days=LEAD_EXPIRE_DAYS)).strftime("%Y-%m-%d %H:%M:%S")
    conn.execute(
        "INSERT INTO leads(id, company, phone_enc, purpose, persona, status, submitted_at, expire_at, consent, mock, "
        "contact_name, province, city, industry_l1, industry_name) "
        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (
            qr["code_id"],
            company,
            encrypt_phone(phone),
            purpose,
            persona,
            "submitted",
            now_iso(),
            expire_at,
            1,
            1 if is_mock() else 0,
            contact_name,
            province,
            city,
            industry_l1,
            industry_name,
        ),
    )
    audit(
        conn,
        "lead_create",
        qr["code_id"],
        f"{persona}:{purpose}:{mask_phone(phone)}:{province}{city}:{industry_l1}/{industry_name}:{contact_name}",
    )
    conn.commit()
    threading.Thread(target=_replenish_bg, daemon=True).start()
    return {
        "lead_id": qr["code_id"],
        "qr_code": DISPLAY_QR,
        "qr_link": WECOM_QR_LINK,
        "mock": is_mock(),
    }


def _replenish_bg() -> None:
    with _db_lock:
        conn = connect()
        try:
            replenish_pool(conn)
        finally:
            conn.close()


def lead_status(conn: sqlite3.Connection, lead_id: str) -> dict:
    row = conn.execute("SELECT id, status, connected_at FROM leads WHERE id=?", (lead_id,)).fetchone()
    if not row:
        raise KeyError("留资记录不存在")
    return {"lead_id": row["id"], "status": row["status"], "connected_at": row["connected_at"]}


def list_leads(conn: sqlite3.Connection, limit: int = 50) -> list[dict]:
    rows = conn.execute(
        "SELECT id, company, phone_enc, purpose, persona, status, submitted_at, connected_at, mock, "
        "contact_name, province, city, industry_l1, industry_name "
        "FROM leads ORDER BY submitted_at DESC LIMIT ?",
        (limit,),
    ).fetchall()
    out = []
    for row in rows:
        item = dict(row)
        item["phone"] = mask_phone(decrypt_phone(item.pop("phone_enc")))
        item["purpose_label"] = purpose_label(item["purpose"])
        out.append(item)
    return out


def check_rate(ip: str, limit: int = 20, window: int = 60) -> bool:
    ts = time.time()
    bucket = [t for t in _rate.get(ip, []) if ts - t < window]
    if len(bucket) >= limit:
        _rate[ip] = bucket
        return False
    bucket.append(ts)
    _rate[ip] = bucket
    return True


def load_industries() -> list[dict]:
    global _industries_cache
    if _industries_cache is None:
        path = ROOT / "lead-industries.json"
        _industries_cache = json.loads(path.read_text(encoding="utf-8"))
    return _industries_cache


def load_regions() -> list[dict]:
    path = ROOT / "lead-regions.json"
    return json.loads(path.read_text(encoding="utf-8"))


def valid_region(province: str, city: str) -> bool:
    for item in load_regions():
        if item.get("name") == province and city in (item.get("cities") or []):
            return True
    return False


def valid_industry(l1: str, name: str) -> bool:
    for item in load_industries():
        if item.get("name") == l1 and name in (item.get("items") or []):
            return True
    return False


def _purge_expired(store: dict) -> None:
    now_ts = time.time()
    dead = [k for k, v in store.items() if v.get("expire", 0) < now_ts]
    for k in dead:
        store.pop(k, None)


def make_captcha() -> dict:
    from io import BytesIO
    from PIL import Image, ImageDraw, ImageFont, ImageFilter
    import base64

    alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"
    code = "".join(random.choice(alphabet) for _ in range(4))
    img = Image.new("RGB", (128, 44), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    font = None
    for fp in (
        r"C:\Windows\Fonts\msyhbd.ttc",
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\arialbd.ttf",
        r"C:\Windows\Fonts\arial.ttf",
    ):
        if Path(fp).exists():
            try:
                font = ImageFont.truetype(fp, 26)
                break
            except Exception:
                continue
    if font is None:
        font = ImageFont.load_default()
    for i, ch in enumerate(code):
        x = 12 + i * 28 + random.randint(-2, 2)
        y = random.randint(4, 10)
        draw.text((x, y), ch, font=font, fill=(random.randint(20, 90), random.randint(20, 90), random.randint(80, 160)))
    for _ in range(6):
        draw.line(
            [(random.randint(0, 128), random.randint(0, 44)), (random.randint(0, 128), random.randint(0, 44))],
            fill=(180, 190, 210),
            width=1,
        )
    img = img.filter(ImageFilter.SMOOTH)
    buf = BytesIO()
    img.save(buf, format="PNG")
    captcha_id = short_id(12)
    with _verify_lock:
        _purge_expired(_captchas)
        _captchas[captcha_id] = {"code": code, "expire": time.time() + 180}
    return {
        "captcha_id": captcha_id,
        "image": "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("ascii"),
        **({"demo_code": code} if is_mock() else {}),
    }


def verify_captcha(captcha_id: str, code: str) -> bool:
    with _verify_lock:
        item = _captchas.pop((captcha_id or "").strip(), None)
    if not item or item["expire"] < time.time():
        return False
    return str(code or "").strip().upper() == item["code"]


def send_sms(phone: str, captcha_id: str, captcha_code: str) -> dict:
    if not PHONE_RE.match(phone):
        raise ValueError("请输入有效的 11 位手机号")
    if not verify_captcha(captcha_id, captcha_code):
        raise ValueError("图形验证码错误或已过期，请刷新后重试")
    sms = "".join(random.choice("0123456789") for _ in range(6))
    with _verify_lock:
        _purge_expired(_sms_codes)
        _sms_codes[phone] = {"code": sms, "expire": time.time() + 300, "used": False}
    print(f"[sms] {mask_phone(phone)} -> {sms}", flush=True)
    payload = {"ok": True, "ttl": 60, "message": "验证码已发送"}
    if is_mock():
        payload["demo_code"] = sms
    return payload


def consume_sms(phone: str, sms_code: str) -> None:
    with _verify_lock:
        item = _sms_codes.get(phone)
        if not item or item["expire"] < time.time() or item.get("used"):
            raise ValueError("请先获取短信验证码")
        if str(sms_code or "").strip() != item["code"]:
            raise ValueError("短信验证码错误")
        item["used"] = True


class LeadHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, directory: str | None = None, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def log_message(self, fmt: str, *args) -> None:
        print(f"[{now_iso()}] {self.address_string()} {fmt % args}")

    def _cors(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET,POST,OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type,Authorization")

    def _json(self, code: int, payload: dict) -> None:
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self._cors()
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def _text(self, code: int, text: str, content_type: str = "text/plain; charset=utf-8") -> None:
        raw = text.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self._cors()
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def _body(self, limit: int = 8192) -> bytes:
        length = int(self.headers.get("Content-Length") or 0)
        if length > limit:
            raise ValueError("请求体过大")
        return self.rfile.read(length) if length else b""

    def _admin_ok(self) -> bool:
        token = env("LEAD_ADMIN_TOKEN")
        if not token:
            return is_mock()
        given = (self.headers.get("Authorization") or "").replace("Bearer ", "").strip()
        qs = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        if not given:
            given = (qs.get("token") or [""])[0]
        return given == token

    def do_OPTIONS(self) -> None:  # noqa: N802
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_GET(self) -> None:  # noqa: N802
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path.rstrip("/") or "/"
        qs = urllib.parse.parse_qs(parsed.query)
        if path == "/api/health":
            idle = 0
            with _db_lock:
                conn = connect()
                try:
                    idle = conn.execute("SELECT COUNT(*) FROM qr_codes WHERE status='idle'").fetchone()[0]
                finally:
                    conn.close()
            self._json(
                200,
                {
                    "ok": True,
                    "mock": is_mock(),
                    "pool_idle": idle,
                    "time": now_iso(),
                },
            )
            return
        if path == "/api/captcha":
            self._json(200, make_captcha())
            return
        if path.startswith("/api/lead/") and path.endswith("/status"):
            lead_id = path.split("/")[3]
            if not LEAD_ID_RE.match(lead_id):
                self._json(400, {"error": "无效的留资 ID"})
                return
            with _db_lock:
                conn = connect()
                try:
                    self._json(200, lead_status(conn, lead_id))
                except KeyError:
                    self._json(404, {"error": "留资记录不存在"})
                finally:
                    conn.close()
            return
        if path == "/api/leads":
            if not self._admin_ok():
                self._json(401, {"error": "需要管理员令牌"})
                return
            with _db_lock:
                conn = connect()
                try:
                    self._json(200, {"items": list_leads(conn)})
                finally:
                    conn.close()
            return
        if path == "/wx/callback":
            self._handle_wx_verify(qs)
            return
        super().do_GET()

    def do_POST(self) -> None:  # noqa: N802
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path.rstrip("/") or "/"
        qs = urllib.parse.parse_qs(parsed.query)
        if path == "/api/sms":
            ip = self.client_address[0]
            if not check_rate(ip, limit=8, window=60):
                self._json(429, {"error": "发送过于频繁，请稍后再试"})
                return
            try:
                payload = json.loads(self._body().decode("utf-8") or "{}")
                result = send_sms(
                    str(payload.get("phone") or "").strip(),
                    str(payload.get("captcha_id") or ""),
                    str(payload.get("captcha_code") or ""),
                )
                self._json(200, result)
            except ValueError as exc:
                self._json(400, {"error": str(exc)})
            except Exception as exc:
                traceback.print_exc()
                self._json(500, {"error": f"发送失败：{exc}"})
            return
        if path == "/api/lead":
            ip = self.client_address[0]
            if not check_rate(ip):
                self._json(429, {"error": "提交过于频繁，请稍后再试"})
                return
            try:
                payload = json.loads(self._body().decode("utf-8") or "{}")
            except Exception:
                self._json(400, {"error": "请求格式错误"})
                return
            with _db_lock:
                conn = connect()
                try:
                    result = create_lead(payload, conn)
                    self._json(200, result)
                except ValueError as exc:
                    self._json(400, {"error": str(exc)})
                except Exception as exc:
                    traceback.print_exc()
                    self._json(500, {"error": f"提交失败：{exc}"})
                finally:
                    conn.close()
            return
        if path == "/api/debug/simulate-scan":
            if not self._admin_ok():
                self._json(401, {"error": "需要管理员令牌"})
                return
            try:
                payload = json.loads(self._body().decode("utf-8") or "{}")
            except Exception:
                self._json(400, {"error": "请求格式错误"})
                return
            lead_id = str(payload.get("lead_id") or "")
            if not LEAD_ID_RE.match(lead_id):
                self._json(400, {"error": "无效的留资 ID"})
                return
            with _db_lock:
                conn = connect()
                try:
                    connect_lead(conn, lead_id, "mock-external", env("WECOM_SALES_USERID") or "mock-sales", "")
                    self._json(200, lead_status(conn, lead_id))
                except KeyError:
                    self._json(404, {"error": "留资记录不存在"})
                finally:
                    conn.close()
            return
        if path == "/wx/callback":
            self._handle_wx_callback(qs)
            return
        self._json(404, {"error": "not found"})

    def _crypto(self) -> WecomCrypto:
        token = env("WECOM_TOKEN")
        aes_key = env("WECOM_AES_KEY")
        corp = env("WECOM_CORP_ID")
        if not (token and aes_key and corp):
            raise WecomCryptoError("未配置 WECOM_TOKEN / WECOM_AES_KEY / WECOM_CORP_ID")
        return WecomCrypto(token, aes_key, corp)

    def _handle_wx_verify(self, qs: dict) -> None:
        echostr = (qs.get("echostr") or [""])[0]
        signature = (qs.get("msg_signature") or [""])[0]
        timestamp = (qs.get("timestamp") or [""])[0]
        nonce = (qs.get("nonce") or [""])[0]
        try:
            plain = self._crypto().verify_url(signature, timestamp, nonce, echostr)
            self._text(200, plain)
        except Exception as exc:
            print(f"[warn] 回调 URL 验证失败: {exc}")
            self._text(403, "verify failed")

    def _handle_wx_callback(self, qs: dict) -> None:
        try:
            body = self._body(65536).decode("utf-8")
            msg = self._crypto().decrypt_message(
                body,
                (qs.get("msg_signature") or [""])[0],
                (qs.get("timestamp") or [""])[0],
                (qs.get("nonce") or [""])[0],
            )
        except Exception as exc:
            print(f"[warn] 回调解密失败: {exc}")
            self._text(200, "success")
            return
        try:
            with _db_lock:
                conn = connect()
                try:
                    handle_wx_event(msg, conn)
                finally:
                    conn.close()
        except Exception:
            traceback.print_exc()
        self._text(200, "success")


def main() -> None:
    load_env(ENV_PATH)
    parser = argparse.ArgumentParser(description="盛意旺落地页留资服务")
    parser.add_argument("--host", default=env("LEAD_HOST") or "0.0.0.0")
    parser.add_argument("--port", type=int, default=env_int("LEAD_PORT", 8088))
    args = parser.parse_args()

    conn = connect()
    init_db(conn)
    try:
        replenish_pool(conn)
    finally:
        conn.close()

    mode = "mock（本地演示，未接真企微）" if is_mock() else "live（已配置企微凭证）"
    callback = env("WECOM_CALLBACK_URL") or f"http://127.0.0.1:{args.port}/wx/callback"
    print(f"留资服务已启动  http://127.0.0.1:{args.port}/", flush=True)
    print(f"模式: {mode}", flush=True)
    print(f"企微回调: {callback}", flush=True)
    print("接口: POST /api/lead   GET /api/lead/<id>/status   GET|POST /wx/callback", flush=True)
    if is_mock():
        print("本地模拟扫码: POST /api/debug/simulate-scan  {\"lead_id\":\"...\"}", flush=True)
    httpd = ThreadingHTTPServer((args.host, args.port), LeadHandler)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")


if __name__ == "__main__":
    main()
