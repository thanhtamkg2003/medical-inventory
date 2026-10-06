"""Kiểm kê thiết bị y tế bằng QR. Chạy thử: python app.py."""
import hashlib
import io
import ipaddress
import json
import os
import re
import secrets
import sqlite3
import tempfile
import unicodedata
from collections import Counter
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from functools import wraps
from pathlib import Path
from urllib.request import Request, urlopen

import qrcode
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from flask import (
    Flask, after_this_request, flash, jsonify, redirect, render_template_string as flask_render,
    request, send_file, session, url_for,
)
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from werkzeug.security import check_password_hash, generate_password_hash

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(APP_DIR, "medical_inventory.db")
APP_SECRET_FILE = os.path.join(APP_DIR, "app_secret.key")

def load_app_secret():
    # APP_SECRET takes priority for server deployments.
    configured = os.environ.get("APP_SECRET", "").strip()
    if configured:
        return configured
    # Keep the Flask session/CSRF signing key stable across restarts.
    # Without this, every restart invalidates the browser session cookie.
    try:
        if os.path.exists(APP_SECRET_FILE):
            value = Path(APP_SECRET_FILE).read_text(encoding="utf-8").strip()
            if len(value) >= 32:
                return value
        value = secrets.token_urlsafe(48)
        Path(APP_SECRET_FILE).write_text(value, encoding="utf-8")
        try:
            os.chmod(APP_SECRET_FILE, 0o600)
        except OSError:
            pass
        return value
    except OSError:
        # Last-resort fallback. This should only be used if the application
        # directory is not writable; an APP_SECRET environment variable is
        # recommended for that deployment scenario.
        return secrets.token_urlsafe(48)

app = Flask(__name__)
app.config["SECRET_KEY"] = load_app_secret()
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(hours=24)
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.environ.get("SESSION_COOKIE_SECURE", "0").lower() in {"1", "true", "yes"}
INACTIVITY_TIMEOUT = timedelta(hours=8)
ABSOLUTE_SESSION_TIMEOUT = timedelta(hours=24)
IP_GEO_TIMEOUT = float(os.environ.get("IP_GEO_TIMEOUT", "2.5"))
IP_GEO_API = os.environ.get("IP_GEO_API", "https://ipapi.co/{ip}/json/")

EXCEL_HEADERS = ["STT", "MÃ TÀI SẢN", "TÊN VT - TB", "SỐ SERI", "MODEL",
                 "HÃNG SX", "NƯỚC SX", "NĂM SX", "NĂM SD", "NĂM BC TĂNG", "KHOA SD"]
EXTRA_FIELDS = ("model", "maker", "country", "year_made", "year_used", "year_increase", "department")
EXTRA_LABELS = ("Model", "Hãng SX", "Nước SX", "Năm SX", "Năm SD", "Năm BC tăng", "Khoa SD")
PERMISSIONS = {
    "devices_view":"Xem danh mục thiết bị", "devices_manage":"Thêm, sửa, xóa thiết bị",
    "excel_import":"Nhập Excel", "round_manage":"Tạo, xóa đợt kiểm kê",
    "qr_export":"Xem, xuất mã QR", "scan":"Quét QR",
    "results_view":"Xem kết quả", "results_export":"Xuất kết quả Excel",
    "accounts_manage":"Quản lý tài khoản thành viên",
}
ADMIN_PERMISSIONS = list(PERMISSIONS)
MEMBER_PERMISSIONS = ["scan", "results_view", "results_export"]

def extra_values(form):
    return [cell_text(form.get(field)) for field in EXTRA_FIELDS]

def recovery_code():
    return "-".join(secrets.token_hex(4).upper() for _ in range(4))

def report_filename(name, round_id):
    # Tên đợt do người dùng nhập; chuyển thành tên file phù hợp Windows.
    plain=unicodedata.normalize("NFKD", name.replace("đ", "d").replace("Đ", "D"))
    plain="".join(char for char in plain if not unicodedata.combining(char))
    slug=re.sub(r"[^A-Za-z0-9]+", "_", plain).strip("_")[:70]
    return f"Bao_cao_{slug or f'dot_{round_id}'}.xlsx"

BASE = """<!doctype html><html lang="vi"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Kiểm kê thiết bị y tế</title><style>
body{font:15px Arial,sans-serif;margin:0;background:#f4f7fb;color:#172033;min-height:100vh;display:flex;flex-direction:column}
.bar{background:#087e8b;color:white;padding:14px 4%;display:flex;gap:14px;flex-wrap:wrap;align-items:center}
.bar b{margin-right:auto}.bar a{color:white;text-decoration:none}.wrap{width:100%;box-sizing:border-box;max-width:1200px;margin:22px auto;padding:0 15px}
.card{background:white;padding:18px;border-radius:10px;box-shadow:0 1px 5px #ccd;margin:14px 0}
.wrap{flex:1}.site-footer{text-align:center;background:#087e8b;color:#fff;padding:18px 14px;font-size:13px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px}
.device-fields{grid-template-columns:repeat(auto-fit,minmax(180px,1fr))}
.device-fields>div{min-width:0}.device-fields input{box-sizing:border-box;width:100%}
.device-submit{grid-column:1/-1;margin-top:8px}
.devices-table td{padding:4px 6px}.devices-table th{padding:6px}
.device-actions{display:flex;gap:7px;align-items:center;white-space:nowrap}
.device-actions button{padding:4px 7px;margin:0}.device-warning{margin-top:5px}
.device-actions .btn,.round-actions .btn,.round-actions button{padding:5px 8px;margin:0;font-size:13px;white-space:nowrap}
.round-actions{display:flex;gap:6px;align-items:center;flex-wrap:wrap}.round-actions form{margin:0}
.filter-links{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0}.filter-links a{padding:7px 11px;border-radius:6px;background:#e7f2f4;color:#075d68;text-decoration:none}
.filter-links a:hover{background:#c5e4e9}.filter-links a[aria-current="page"],.filter-links a[aria-current="page"]:hover{background:#087e8b;color:white}
[aria-current="page"]{font-weight:bold;color:#075d68}
.stat{font-size:28px;font-weight:bold;color:#087e8b}
.overview-stats{grid-template-columns:repeat(auto-fit,minmax(210px,1fr));align-items:stretch}
.overview-stats .card{margin:0;min-height:90px}.overview-round{margin-bottom:18px}
.progress-track{height:18px;background:#dce8eb;border-radius:20px;overflow:hidden}
.progress-fill{height:100%;background:#087e8b;border-radius:20px;transition:width .2s}
.field-error{color:#9b2234;font-size:13px;font-weight:bold}
.invalid-field{border:2px solid #ae3040}
input,select,button{padding:9px;margin:5px 0;border:1px solid #bbc;border-radius:6px;font:inherit}
button,.btn,.btn:visited{background:#087e8b;color:white;border:0;text-decoration:none;display:inline-block;cursor:pointer;padding:9px;border-radius:6px;margin:4px 0}
.secondary{background:#4e6475}.warn{background:#fff1d9;border-left:5px solid #c47700;padding:12px}
.btn.secondary,.btn.secondary:visited{background:#4e6475}.btn.danger,.btn.danger:visited{background:#ae3040}
.manual-entry{margin-top:18px;border:1px solid #d6dce5;border-radius:8px;padding:12px}.manual-entry summary{cursor:pointer;color:#075d68;font-weight:bold;padding:5px}.manual-entry summary:hover{color:#087e8b;background:#e7f2f4}
.ok{background:#e5f7e8;border-left:5px solid #26823e;padding:12px}
.flash{padding:10px;background:#fff4c6;margin-bottom:10px}
table{border-collapse:collapse;width:100%;font-size:14px}th,td{border:1px solid #d6dce5;padding:8px;text-align:center;vertical-align:middle;overflow-wrap:anywhere;word-break:break-word}
th{background:#edf3f8;text-align:center;vertical-align:middle}.table-name{text-align:left;vertical-align:middle}.device-actions,.round-actions{justify-content:center}.table-wrap{overflow-x:auto}
.audit-filter-actions{display:flex;gap:8px;align-items:center;justify-content:flex-start;flex-wrap:nowrap;white-space:nowrap;align-self:end;padding-bottom:5px}.audit-filter-actions button,.audit-filter-actions .btn{margin:0;height:34px;box-sizing:border-box;padding:0 12px;display:inline-flex;align-items:center;justify-content:center}.audit-table{table-layout:fixed;font-size:13px}.audit-table th,.audit-table td{padding:7px;overflow-wrap:anywhere;word-break:break-word;white-space:normal}.audit-table .col-time{width:8%}.audit-table .col-user{width:9%}.audit-table .col-role{width:8%}.audit-table .col-action{width:9%}.audit-table .col-target{width:12%}.audit-table .col-result{width:8%}.audit-table .col-ip{width:9%}.audit-table .col-location{width:13%}.audit-table .col-details{width:24%}.audit-sessions{table-layout:fixed;font-size:13px}.audit-sessions th,.audit-sessions td{overflow-wrap:anywhere;word-break:break-word;white-space:normal}.audit-sessions .col-session-user{width:10%}.audit-sessions .col-session-role{width:8%}.audit-sessions .col-session-time{width:12%}.audit-sessions .col-session-last{width:12%}.audit-sessions .col-session-ip{width:10%}.audit-sessions .col-session-location{width:13%}.audit-sessions .col-session-device{width:22%}.audit-sessions .col-session-action{width:13%}
label{display:block;font-weight:bold;margin-top:8px}
.muted{color:#687487}.serial{white-space:pre-line;overflow-wrap:anywhere}
.pager{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:16px 0}
.busy-layer{position:fixed;inset:0;background:rgba(9,30,45,.58);z-index:20;display:none;align-items:center;justify-content:center;padding:20px}.busy-layer.show{display:flex}.busy-box{background:#fff;border-radius:12px;padding:24px;max-width:350px;text-align:center;box-shadow:0 4px 20px #123}.spinner{width:34px;height:34px;border:5px solid #d7e4e7;border-top-color:#087e8b;border-radius:50%;margin:0 auto 14px;animation:spin .8s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}
.bar details{position:relative}.bar summary{cursor:pointer;list-style:none}.bar details[open] .admin-menu{position:absolute;right:0;top:28px;background:#087e8b;min-width:145px;padding:8px 12px;border-radius:6px;z-index:3;box-shadow:0 2px 7px #456}.admin-menu a{display:block;padding:8px}
.bar>a{padding:9px 11px;border-radius:6px}.bar>a:hover,.bar summary:hover,.admin-menu a:hover{background:#146879;color:#fff}
.bar>a.active,.bar details.active>summary{background:#fff;color:#075d68;font-weight:bold;box-shadow:inset 0 -3px #f4a340}
.bar summary{padding:9px 11px;border-radius:6px}.bar details.active>summary:hover{color:#075d68;background:#e9f9fa}
button:hover,.btn:hover,.btn:visited:hover{background:#06616c;color:white}.secondary:hover,.btn.secondary:hover{background:#344957}
.danger{background:#ae3040}.danger:hover,.btn.danger:hover{background:#842132}a:hover{text-decoration:underline}
button:focus-visible,.btn:focus-visible,.bar a:focus-visible,.bar summary:focus-visible{outline:3px solid #efab34;outline-offset:2px}
@media(max-width:650px){.bar b{width:100%}table{font-size:12px}.audit-filter-actions{align-items:stretch;align-self:stretch;padding-bottom:0}.audit-filter-actions button,.audit-filter-actions .btn{margin:0;height:38px;flex:1}.audit-table,.audit-sessions{font-size:11px}.audit-table th,.audit-table td,.audit-sessions th,.audit-sessions td{padding:5px}}
.nav-logout{background:transparent;padding:0;color:white;margin:0;border:0;cursor:pointer;font:inherit}.nav-logout:hover{text-decoration:underline}
</style></head><body>
{% if session.user_id %}<nav class="bar"><b>KIỂM KÊ THIẾT BỊ Y TẾ</b>
<a class="{{'active' if request.endpoint=='home' else ''}}" {% if request.endpoint=='home' %}aria-current="page"{% endif %} href="{{url_for('home')}}">Tổng quan</a>
{% if can('devices_view') %}
<a class="{{'active' if request.endpoint in ('devices','edit_device','delete_device','clear_devices','import_excel','import_result','import_issues_excel','excel_template','labels','qr_view','qr_png') else ''}}" href="{{url_for('devices')}}">Thiết bị</a>
{% endif %}
{% if can('round_manage') or can('results_view') or can('results_export') %}<a class="{{'active' if request.endpoint in ('rounds','delete_round','results','report') else ''}}" href="{{url_for('rounds')}}">Đợt kiểm kê</a>{% endif %}
{% if can('scan') %}<a class="{{'active' if request.endpoint in ('active_scan','scan') else ''}}" href="{{url_for('active_scan')}}">Quét QR</a>{% endif %}
<details class="{{'active' if request.endpoint in ('change_password','users','reset_member_password','delete_member','audit_log','audit_export','revoke_user_sessions','backup_db') else ''}}"><summary>{{session.username}} ▾</summary><div class="admin-menu">
<a href="{{url_for('change_password')}}">Đổi mật khẩu</a>
{% if can('accounts_manage') %}<a href="{{url_for('users')}}">Tài khoản</a>{% endif %}
{% if session.role=='superadmin' %}<a href="{{url_for('audit_log')}}">Nhật ký</a><a href="{{url_for('backup_db')}}">Sao lưu dữ liệu</a>{% endif %}
<form method="post" action="{{url_for('logout')}}" style="display:inline;margin:0">
<input type="hidden" name="_csrf" value="{{csrf_token_value}}"><button class="nav-logout" type="submit">Đăng xuất</button></form>
</div></details>
</nav>{% endif %}
<main class="wrap">{% for message in get_flashed_messages() %}<div class="flash">{{message}}</div>{% endfor %}
{{content|safe}}</main><footer class="site-footer">Thiết kế bởi: Nguyễn Thành Tâm - Giáo viên trường THCS Thạnh Quới, xã Thạnh Quới, TP Cần Thơ. SĐT: 0922118113</footer></body></html>"""

def page(template, **values):
    content = inject_csrf_fields(flask_render(template, **values))
    return flask_render(BASE, content=content, csrf_token_value=session.get("csrf_token", ""), **values)

@app.context_processor
def inject_permissions():
    if not session.get("csrf_token"):
        session["csrf_token"]=secrets.token_urlsafe(32)
    return {"can":has_permission,"csrf_token_value":session.get("csrf_token", "")}

@app.before_request
def protect_csrf():
    # Đăng nhập là điểm khởi đầu của phiên. Không kiểm tra CSRF ở endpoint này
    # vì trình duyệt có thể đang giữ cookie phiên/CSRF cũ sau khi Flask được
    # khởi động lại. Với GET /login, tạo một phiên sạch để hiển thị form mới.
    if request.endpoint == "login":
        if request.method == "GET":
            session.clear()
            session["csrf_token"] = secrets.token_urlsafe(32)
        return None

    # Các endpoint còn lại bắt buộc phải có CSRF token.
    if not session.get("csrf_token"):
        session["csrf_token"] = secrets.token_urlsafe(32)

    if request.method in {"POST","PUT","PATCH","DELETE"}:
        expected=session.get("csrf_token", "")
        supplied=request.form.get("_csrf", "") or request.headers.get("X-CSRFToken", "")
        if not expected or not supplied or not secrets.compare_digest(expected,supplied):
            return "Yêu cầu không hợp lệ (CSRF). Hãy tải lại trang và thử lại.",400
    return None

def db():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con

def now():
    return datetime.now(ZoneInfo("Asia/Ho_Chi_Minh")).strftime("%Y-%m-%d %H:%M:%S")

def cell_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()

def serial_key(value):
    return " ".join(cell_text(value).split())

def append_excel_values(sheet, values):
    sheet.append(values)
    for cell in sheet[sheet.max_row]:
        if isinstance(cell.value,str):
            cell.data_type="s"

def dt_text(value):
    return value.strftime("%Y-%m-%d %H:%M:%S") if isinstance(value, datetime) else str(value)

def parse_dt(value):
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
    except (TypeError, ValueError):
        return None

def session_token_hash(token):
    return hashlib.sha256(token.encode("utf-8")).hexdigest()

def request_client_ip():
    # Chỉ tin X-Forwarded-For khi người quản trị chủ động bật TRUST_PROXY.
    trust_proxy=os.environ.get("TRUST_PROXY", "1").lower() in {"1", "true", "yes"}
    if trust_proxy:
        forwarded=request.headers.get("X-Forwarded-For", "")
        if forwarded:
            return forwarded.split(",")[0].strip()
    return (request.remote_addr or "").strip()

def is_public_ip(ip):
    try:
        addr=ipaddress.ip_address(ip)
        return not (addr.is_private or addr.is_loopback or addr.is_link_local or addr.is_reserved or addr.is_unspecified)
    except ValueError:
        return False

def geo_lookup(con, ip):
    """Tra cứu địa lý IP ở mức gần đúng; lỗi dịch vụ bên ngoài không chặn đăng nhập."""
    if not ip:
        return {"location":"Không xác định", "isp":"", "latitude":"", "longitude":""}
    if not is_public_ip(ip):
        return {"location":"Không xác định (IP nội bộ)", "isp":"", "latitude":"", "longitude":""}
    cached=con.execute("SELECT * FROM ip_geolocation WHERE ip_address=?", (ip,)).fetchone()
    if cached:
        return {"location":cached["location"],"isp":cached["isp"],"latitude":cached["latitude"],"longitude":cached["longitude"]}
    try:
        url=IP_GEO_API.format(ip=ip)
        req=Request(url,headers={"User-Agent":"MedicalInventory/2.0"})
        with urlopen(req,timeout=IP_GEO_TIMEOUT) as resp:
            data=json.loads(resp.read().decode("utf-8",errors="replace"))
        city=cell_text(data.get("city"))
        region=cell_text(data.get("region"))
        country=cell_text(data.get("country_name") or data.get("country"))
        location=", ".join(x for x in (city,region,country) if x) or "Không xác định"
        isp=cell_text(data.get("org") or data.get("asn"))
        latitude=cell_text(data.get("latitude"))
        longitude=cell_text(data.get("longitude"))
        con.execute("""INSERT INTO ip_geolocation(ip_address,location,isp,latitude,longitude,fetched_at)
                       VALUES(?,?,?,?,?,?)
                       ON CONFLICT(ip_address) DO UPDATE SET location=excluded.location,isp=excluded.isp,
                       latitude=excluded.latitude,longitude=excluded.longitude,fetched_at=excluded.fetched_at""",
                    (ip,location,isp,latitude,longitude,now()))
        return {"location":location,"isp":isp,"latitude":latitude,"longitude":longitude}
    except Exception:
        return {"location":"Không tra được vị trí IP", "isp":"", "latitude":"", "longitude":""}

def user_agent_summary(ua):
    ua=ua or ""
    browser="Trình duyệt không xác định"
    if "Edg/" in ua: browser="Microsoft Edge"
    elif "Chrome/" in ua: browser="Google Chrome"
    elif "Firefox/" in ua: browser="Mozilla Firefox"
    elif "Safari/" in ua and "Chrome/" not in ua: browser="Safari"
    os_name="Hệ điều hành không xác định"
    if "Windows NT" in ua: os_name="Windows"
    elif "Android" in ua: os_name="Android"
    elif "iPhone" in ua or "iPad" in ua: os_name="iOS"
    elif "Mac OS X" in ua: os_name="macOS"
    elif "Linux" in ua: os_name="Linux"
    return f"{browser} / {os_name}"

def current_session_info():
    return {
        "session_id":session.get("session_id", ""),
        "ip":session.get("client_ip", request_client_ip()),
        "location":session.get("client_location", ""),
        "isp":session.get("client_isp", ""),
        "user_agent":session.get("user_agent_summary", user_agent_summary(request.headers.get("User-Agent", ""))),
    }

def write_audit(con, action, *, success=True, target_type="", target_id=None, target_label="",
                details="", user_id=None, username=None, role=None, session_info=None):
    info=session_info or current_session_info()
    con.execute("""INSERT INTO audit_logs
        (user_id,username,role,action,target_type,target_id,target_label,success,details,
         ip_address,location,isp,user_agent,session_id_hash,created_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (user_id if user_id is not None else session.get("user_id"),
         username if username is not None else session.get("username", ""),
         role if role is not None else session.get("role", ""), action,target_type,target_id,target_label,
         1 if success else 0,details,info.get("ip", ""),info.get("location", ""),info.get("isp", ""),
         info.get("user_agent", ""),session_token_hash(info["session_id"]) if info.get("session_id") else "",now()))

def ensure_auth_session(con, user):
    now_dt=datetime.now()
    sid=session.get("session_id")
    if not sid:
        sid=secrets.token_urlsafe(32)
        session["session_id"]=sid
        session["session_started_at"]=dt_text(now_dt)
        session["last_activity_at"]=dt_text(now_dt)
        client_ip=request_client_ip()
        geo=geo_lookup(con,client_ip)
        session["client_ip"]=client_ip
        session["client_location"]=geo["location"]
        session["client_isp"]=geo["isp"]
        session["user_agent_summary"]=user_agent_summary(request.headers.get("User-Agent", ""))
        con.execute("""INSERT OR IGNORE INTO auth_sessions
            (session_token_hash,user_id,created_at,last_activity,expires_at,ip_address,location,isp,user_agent,status)
            VALUES(?,?,?,?,?,?,?,?,?,?)""",
            (session_token_hash(sid),user["id"],dt_text(now_dt),dt_text(now_dt),
             dt_text(now_dt+ABSOLUTE_SESSION_TIMEOUT),client_ip,geo["location"],geo["isp"],
             session["user_agent_summary"],"active"))
        return True, ""
    row=con.execute("SELECT * FROM auth_sessions WHERE session_token_hash=? AND user_id=?",
                    (session_token_hash(sid),user["id"])).fetchone()
    if not row:
        return False, "Phiên đăng nhập không còn tồn tại."
    started=parse_dt(session.get("session_started_at")) or parse_dt(row["created_at"]) or now_dt
    last=parse_dt(session.get("last_activity_at")) or parse_dt(row["last_activity"]) or started
    if now_dt-last > INACTIVITY_TIMEOUT:
        con.execute("UPDATE auth_sessions SET status='expired',ended_at=? WHERE id=?",(dt_text(now_dt),row["id"]))
        write_audit(con,"Tự động đăng xuất do không hoạt động",details="Phiên vượt quá 8 giờ không hoạt động.")
        return False, "Phiên làm việc đã hết hạn do không hoạt động. Vui lòng đăng nhập lại."
    if now_dt-started > ABSOLUTE_SESSION_TIMEOUT:
        con.execute("UPDATE auth_sessions SET status='expired',ended_at=? WHERE id=?",(dt_text(now_dt),row["id"]))
        write_audit(con,"Tự động đăng xuất do hết hạn phiên",details="Phiên vượt quá thời gian sống tối đa 24 giờ.")
        return False, "Phiên làm việc đã hết hạn. Vui lòng đăng nhập lại."
    con.execute("UPDATE auth_sessions SET last_activity=?,status='active' WHERE id=?",(dt_text(now_dt),row["id"]))
    session["last_activity_at"]=dt_text(now_dt)
    session.modified=True
    return True, ""

def inject_csrf_fields(html):
    token=session.get("csrf_token", "")
    if not token:
        token=secrets.token_urlsafe(32)
        session["csrf_token"]=token
    field=f'<input type="hidden" name="_csrf" value="{token}">'
    def add_field(match):
        attrs=match.group(1) or ""
        method=re.search(r"\bmethod\s*=\s*[\"\'](post|put|patch|delete)[\"\']",attrs,re.I)
        return match.group(0)+field if method else match.group(0)
    return re.sub(r"<form\b([^>]*)>", add_field, html, flags=re.I)

def init_db():
    with db() as con:
        con.executescript("""
        CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY,value TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS users(
          id INTEGER PRIMARY KEY AUTOINCREMENT,username TEXT NOT NULL UNIQUE,
          password_hash TEXT NOT NULL,role TEXT NOT NULL CHECK(role IN ('superadmin','admin','member')),
          auth_version INTEGER NOT NULL DEFAULT 1,created_at TEXT NOT NULL,
          permissions TEXT NOT NULL DEFAULT '[]',disabled INTEGER NOT NULL DEFAULT 0);
        CREATE TABLE IF NOT EXISTS devices(
          id INTEGER PRIMARY KEY AUTOINCREMENT,stt TEXT NOT NULL DEFAULT '',
          asset_code TEXT NOT NULL DEFAULT '',name TEXT NOT NULL,
          serial TEXT NOT NULL DEFAULT '',serial_key TEXT NOT NULL DEFAULT '',
          department TEXT NOT NULL DEFAULT '',maker_model TEXT NOT NULL DEFAULT '',
          status TEXT NOT NULL DEFAULT 'Đang sử dụng',created_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS inventory_rounds(
          id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT NOT NULL,
          created_at TEXT NOT NULL,active INTEGER NOT NULL DEFAULT 1,
          duplicate_mode TEXT NOT NULL DEFAULT 'block');
        CREATE TABLE IF NOT EXISTS scans(
          id INTEGER PRIMARY KEY AUTOINCREMENT,round_id INTEGER NOT NULL,
          device_id INTEGER NOT NULL,scanned_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS scan_once(
          round_id INTEGER NOT NULL,device_id INTEGER NOT NULL,
          PRIMARY KEY(round_id,device_id));
        CREATE TABLE IF NOT EXISTS round_items(
          id INTEGER PRIMARY KEY AUTOINCREMENT,round_id INTEGER NOT NULL,
          device_id INTEGER NOT NULL,stt TEXT NOT NULL DEFAULT '',
          asset_code TEXT NOT NULL DEFAULT '',name TEXT NOT NULL,
          serial TEXT NOT NULL DEFAULT '',serial_key TEXT NOT NULL DEFAULT '',
          UNIQUE(round_id,device_id));
        CREATE TABLE IF NOT EXISTS import_batches(
          id INTEGER PRIMARY KEY AUTOINCREMENT,filename TEXT NOT NULL,
          mode TEXT NOT NULL,created_at TEXT NOT NULL,added INTEGER NOT NULL,
          skipped INTEGER NOT NULL,missing INTEGER NOT NULL,
          duplicate INTEGER NOT NULL);
        CREATE TABLE IF NOT EXISTS import_issues(
          id INTEGER PRIMARY KEY AUTOINCREMENT,batch_id INTEGER NOT NULL,
          source_row INTEGER NOT NULL,reason TEXT NOT NULL,
          stt TEXT NOT NULL,asset_code TEXT NOT NULL,
          name TEXT NOT NULL,serial TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS audit_logs(
          id INTEGER PRIMARY KEY AUTOINCREMENT,user_id INTEGER,username TEXT NOT NULL DEFAULT '',
          role TEXT NOT NULL DEFAULT '',action TEXT NOT NULL,target_type TEXT NOT NULL DEFAULT '',
          target_id INTEGER,target_label TEXT NOT NULL DEFAULT '',success INTEGER NOT NULL DEFAULT 1,
          details TEXT NOT NULL DEFAULT '',ip_address TEXT NOT NULL DEFAULT '',
          location TEXT NOT NULL DEFAULT '',isp TEXT NOT NULL DEFAULT '',
          user_agent TEXT NOT NULL DEFAULT '',session_id_hash TEXT NOT NULL DEFAULT '',
          created_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS auth_sessions(
          id INTEGER PRIMARY KEY AUTOINCREMENT,session_token_hash TEXT NOT NULL UNIQUE,
          user_id INTEGER NOT NULL,created_at TEXT NOT NULL,last_activity TEXT NOT NULL,
          expires_at TEXT NOT NULL,ended_at TEXT NOT NULL DEFAULT '',ip_address TEXT NOT NULL DEFAULT '',
          location TEXT NOT NULL DEFAULT '',isp TEXT NOT NULL DEFAULT '',user_agent TEXT NOT NULL DEFAULT '',
          status TEXT NOT NULL DEFAULT 'active');
        CREATE TABLE IF NOT EXISTS ip_geolocation(
          ip_address TEXT PRIMARY KEY,location TEXT NOT NULL DEFAULT '',isp TEXT NOT NULL DEFAULT '',
          latitude TEXT NOT NULL DEFAULT '',longitude TEXT NOT NULL DEFAULT '',fetched_at TEXT NOT NULL);
        """)
        columns = {r["name"] for r in con.execute("PRAGMA table_info(devices)")}
        if "asset_code" not in columns:
            # Bản rất cũ đặt UNIQUE trên serial. Giữ id để scans không mất liên kết.
            con.executescript("""
            CREATE TABLE devices_new(
              id INTEGER PRIMARY KEY AUTOINCREMENT,stt TEXT NOT NULL DEFAULT '',
              asset_code TEXT NOT NULL DEFAULT '',name TEXT NOT NULL,
              serial TEXT NOT NULL DEFAULT '',serial_key TEXT NOT NULL DEFAULT '',
              department TEXT NOT NULL DEFAULT '',maker_model TEXT NOT NULL DEFAULT '',
              status TEXT NOT NULL DEFAULT 'Đang sử dụng',created_at TEXT NOT NULL);
            INSERT INTO devices_new(id,name,serial,department,maker_model,status,created_at)
              SELECT id,name,serial,department,maker_model,status,created_at FROM devices;
            DROP TABLE devices;
            ALTER TABLE devices_new RENAME TO devices;
            """)
            for row in con.execute("SELECT id,serial FROM devices"):
                con.execute("UPDATE devices SET serial_key=? WHERE id=?",
                            (serial_key(row["serial"]),row["id"]))
        con.execute("CREATE INDEX IF NOT EXISTS ix_device_serial ON devices(serial_key,status)")
        con.execute("CREATE INDEX IF NOT EXISTS ix_round_serial ON round_items(round_id,serial_key)")
        if not con.execute("SELECT 1 FROM settings WHERE key='password'").fetchone():
            con.execute("INSERT INTO settings VALUES('password',?)",
                        (generate_password_hash("admin123"),))
        for table in ("devices", "round_items"):
            columns = {r["name"] for r in con.execute(f"PRAGMA table_info({table})")}
            for field in EXTRA_FIELDS:
                if field not in columns:
                    con.execute(f"ALTER TABLE {table} ADD COLUMN {field} TEXT NOT NULL DEFAULT ''")
            if table == "devices" and "model" not in columns:
                con.execute("UPDATE devices SET model=maker_model WHERE model='' AND maker_model<>''")
        scan_columns={r["name"] for r in con.execute("PRAGMA table_info(scans)")}
        if "scanned_by" not in scan_columns:
            con.execute("ALTER TABLE scans ADD COLUMN scanned_by TEXT NOT NULL DEFAULT ''")
        if "client_event_id" not in scan_columns:
            con.execute("ALTER TABLE scans ADD COLUMN client_event_id TEXT NOT NULL DEFAULT ''")
        if "client_scanned_at" not in scan_columns:
            con.execute("ALTER TABLE scans ADD COLUMN client_scanned_at TEXT NOT NULL DEFAULT ''")
        con.execute("CREATE UNIQUE INDEX IF NOT EXISTS ux_scans_client_event_id ON scans(client_event_id) WHERE client_event_id<>''")
        con.execute("CREATE INDEX IF NOT EXISTS ix_audit_created ON audit_logs(created_at)")
        con.execute("CREATE INDEX IF NOT EXISTS ix_audit_user ON audit_logs(username,created_at)")
        con.execute("CREATE INDEX IF NOT EXISTS ix_sessions_user_status ON auth_sessions(user_id,status)")
        if not con.execute("SELECT 1 FROM settings WHERE key='round_snapshot_v1'").fetchone():
            # Nâng cấp các đợt từ phiên bản trước; khi ấy danh mục chưa có cơ chế thay thế.
            con.execute("""INSERT OR IGNORE INTO round_items
                (round_id,device_id,stt,asset_code,name,serial,serial_key)
                SELECT r.id,d.id,d.stt,d.asset_code,d.name,d.serial,d.serial_key
                FROM inventory_rounds r CROSS JOIN devices d
                WHERE d.status='Đang sử dụng'""")
            con.execute("INSERT INTO settings VALUES('round_snapshot_v1','1')")
        con.execute("""INSERT OR IGNORE INTO scan_once(round_id,device_id)
            SELECT DISTINCT round_id,device_id FROM scans""")
        admin_created=not con.execute("SELECT 1 FROM users WHERE username='admin'").fetchone()
        if admin_created:
            legacy=con.execute("SELECT value FROM settings WHERE key='password'").fetchone()
            con.execute("INSERT INTO users(username,password_hash,role,created_at) VALUES(?,?,?,?)",
                        ("admin",legacy["value"],"admin",now()))
        user_schema=con.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='users'").fetchone()[0]
        if "superadmin" not in user_schema:
            # Bảng cũ chỉ chấp nhận admin/member, cần giữ nguyên id và mật khẩu khi nâng cấp.
            con.execute("ALTER TABLE users RENAME TO users_legacy")
            con.execute("""CREATE TABLE users(
                id INTEGER PRIMARY KEY AUTOINCREMENT,username TEXT NOT NULL UNIQUE,
                password_hash TEXT NOT NULL,role TEXT NOT NULL CHECK(role IN ('superadmin','admin','member')),
                auth_version INTEGER NOT NULL DEFAULT 1,created_at TEXT NOT NULL,
                permissions TEXT NOT NULL DEFAULT '[]',disabled INTEGER NOT NULL DEFAULT 0)""")
            con.execute("""INSERT INTO users(id,username,password_hash,role,auth_version,created_at,permissions)
                SELECT id,username,password_hash,role,auth_version,created_at,
                CASE WHEN role='admin' THEN ? ELSE ? END FROM users_legacy""",
                (json.dumps(ADMIN_PERMISSIONS),json.dumps(MEMBER_PERMISSIONS)))
            con.execute("DROP TABLE users_legacy")
        elif admin_created:
            con.execute("UPDATE users SET permissions=? WHERE username='admin'",
                        (json.dumps(ADMIN_PERMISSIONS),))
        if not con.execute("SELECT 1 FROM settings WHERE key='round_extra_v1'").fetchone():
            con.execute("""UPDATE round_items SET
                model=COALESCE((SELECT d.model FROM devices d WHERE d.id=round_items.device_id),''),
                maker=COALESCE((SELECT d.maker FROM devices d WHERE d.id=round_items.device_id),''),
                country=COALESCE((SELECT d.country FROM devices d WHERE d.id=round_items.device_id),''),
                year_made=COALESCE((SELECT d.year_made FROM devices d WHERE d.id=round_items.device_id),''),
                year_used=COALESCE((SELECT d.year_used FROM devices d WHERE d.id=round_items.device_id),''),
                year_increase=COALESCE((SELECT d.year_increase FROM devices d WHERE d.id=round_items.device_id),''),
                department=COALESCE((SELECT d.department FROM devices d WHERE d.id=round_items.device_id),'')""")
            con.execute("INSERT INTO settings VALUES('round_extra_v1','1')")

def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        user_id=session.get("user_id")
        with db() as con:
            user=con.execute("SELECT * FROM users WHERE id=?",(user_id,)).fetchone() if user_id else None
            if not user or user["disabled"] or session.get("auth_version")!=user["auth_version"]:
                if user and session.get("session_id"):
                    con.execute("UPDATE auth_sessions SET status='revoked',ended_at=? WHERE session_token_hash=?",
                                (now(),session_token_hash(session["session_id"])))
                    write_audit(con,"Phiên bị vô hiệu hóa",details="Tài khoản đã bị khóa, đổi quyền hoặc đổi mật khẩu.",
                                user_id=user["id"],username=user["username"],role=user["role"])
                session.clear()
                if request.accept_mimetypes.best=="application/json" or request.headers.get("X-Requested-With")=="XMLHttpRequest":
                    return jsonify({"success":False,"code":"AUTH_EXPIRED","message":"Phiên đăng nhập đã hết hạn hoặc bị vô hiệu hóa. Vui lòng đăng nhập lại."}),401
                return redirect(url_for("login"))
            valid,message=ensure_auth_session(con,user)
            if not valid:
                session.clear()
                if request.accept_mimetypes.best=="application/json" or request.headers.get("X-Requested-With")=="XMLHttpRequest":
                    return jsonify({"success":False,"code":"SESSION_EXPIRED","message":message}),401
                flash(message)
                return redirect(url_for("login"))
            session["role"]=user["role"]
            session["username"]=user["username"]
            session["permissions"]=json.loads(user["permissions"])
        return fn(*args, **kwargs)
    return wrapper

def has_permission(permission):
    if session.get("role")=="superadmin": return True
    granted=session.get("permissions",[])
    if permission in granted: return True
    if permission=="devices_view" and any(p in granted for p in
           ("devices_manage","excel_import","qr_export")): return True
    return False

def permission_required(permission):
    def decorate(fn):
        @wraps(fn)
        @login_required
        def wrapper(*args, **kwargs):
            if not has_permission(permission):
                with db() as con:
                    write_audit(con,"Từ chối truy cập",success=False,target_type="Chức năng",target_label=permission,
                                details="Tài khoản không có quyền sử dụng chức năng này.")
                return "Tài khoản không có quyền sử dụng chức năng này.",403
            return fn(*args, **kwargs)
        return wrapper
    return decorate

def current_round(con):
    return con.execute("SELECT * FROM inventory_rounds WHERE active=1 ORDER BY id DESC LIMIT 1").fetchone()

def add_round_item(con, round_id, device_id, stt, code, name, serial, key, extras=None):
    extras=extras or [""]*7
    con.execute("""INSERT OR IGNORE INTO round_items
        (round_id,device_id,stt,asset_code,name,serial,serial_key,
         model,maker,country,year_made,year_used,year_increase,department)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (round_id,device_id,stt,code,name,serial,key,*extras))

def next_stt(con):
    numbers=(int(row[0]) for row in con.execute(
        "SELECT stt FROM devices WHERE status='Đang sử dụng'")
        if row[0].isdigit())
    return str(max(numbers,default=0)+1)

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        attempted=request.form.get("username","").strip()
        password=request.form.get("password","")
        ip=request_client_ip()
        with db() as con:
            user=con.execute("SELECT * FROM users WHERE username=?",(attempted,)).fetchone()
            ok=bool(user and not user["disabled"] and check_password_hash(user["password_hash"],password))
            geo=geo_lookup(con,ip)
            ua=user_agent_summary(request.headers.get("User-Agent", ""))
            if ok:
                session.clear()
                session.permanent=True
                now_dt=datetime.now()
                sid=secrets.token_urlsafe(32)
                session.update(user_id=user["id"],username=user["username"],role=user["role"],
                               auth_version=user["auth_version"],permissions=json.loads(user["permissions"]),
                               session_id=sid,session_started_at=dt_text(now_dt),last_activity_at=dt_text(now_dt),
                               client_ip=ip,client_location=geo["location"],client_isp=geo["isp"],
                               user_agent_summary=ua,csrf_token=secrets.token_urlsafe(32))
                con.execute("""INSERT INTO auth_sessions
                    (session_token_hash,user_id,created_at,last_activity,expires_at,ip_address,location,isp,user_agent,status)
                    VALUES(?,?,?,?,?,?,?,?,?,?)""",
                    (session_token_hash(sid),user["id"],dt_text(now_dt),dt_text(now_dt),
                     dt_text(now_dt+ABSOLUTE_SESSION_TIMEOUT),ip,geo["location"],geo["isp"],ua,"active"))
                info={"session_id":sid,"ip":ip,"location":geo["location"],"isp":geo["isp"],"user_agent":ua}
                write_audit(con,"Đăng nhập",success=True,target_type="Tài khoản",target_id=user["id"],
                            target_label=user["username"],details="Đăng nhập thành công.",user_id=user["id"],
                            username=user["username"],role=user["role"],session_info=info)
                return redirect(url_for("home"))
            role=user["role"] if user else ""
            user_id=user["id"] if user else None
            write_audit(con,"Đăng nhập",success=False,target_type="Tài khoản",target_id=user_id,
                        target_label=attempted,details="Đăng nhập thất bại: sai tài khoản/mật khẩu hoặc tài khoản bị khóa.",
                        user_id=user_id,username=attempted,role=role,
                        session_info={"session_id":"","ip":ip,"location":geo["location"],"isp":geo["isp"],"user_agent":ua})
        if user and user["role"]=="superadmin" and not user["disabled"]:
            session["superadmin_failures"]=session.get("superadmin_failures",0)+1
        else:
            session.pop("superadmin_failures",None)
        flash("Tên đăng nhập hoặc mật khẩu chưa đúng.")
    return page("""<div class="card" style="max-width:410px;margin:65px auto">
        <h2>Đăng nhập</h2><form method="post">
        <label>Tên đăng nhập</label><input name="username" autocomplete="username" required>
        <label>Mật khẩu</label><input type="password" name="password" autocomplete="current-password" required>
        <button>Đăng nhập</button></form>
        {% if show_recovery %}<p><a class="btn secondary" href="{{url_for('recover_superadmin')}}">Khôi phục Superadmin</a></p>{% endif %}
        </div>""",show_recovery=session.get("superadmin_failures",0)>=3)

@app.route("/logout",methods=["POST"])
@login_required
def logout():
    with db() as con:
        if session.get("session_id"):
            con.execute("UPDATE auth_sessions SET status='logged_out',ended_at=? WHERE session_token_hash=?",
                        (now(),session_token_hash(session["session_id"])))
            write_audit(con,"Đăng xuất",details="Người dùng chủ động đăng xuất.")
    session.clear()
    return redirect(url_for("login"))

@app.route("/password",methods=["GET","POST"])
@login_required
def change_password():
    if request.method=="POST":
        old=request.form.get("old","")
        new=request.form.get("new","")
        repeat=request.form.get("repeat","")
        with db() as con:
            user=con.execute("SELECT * FROM users WHERE id=?",(session["user_id"],)).fetchone()
            stored=user["password_hash"]
            if not check_password_hash(stored,old):
                write_audit(con,"Đổi mật khẩu",success=False,target_type="Tài khoản",target_id=user["id"],
                            target_label=user["username"],details="Thất bại: mật khẩu hiện tại không đúng.")
                flash("Mật khẩu hiện tại không đúng.")
            elif len(new)<8:
                write_audit(con,"Đổi mật khẩu",success=False,target_type="Tài khoản",target_id=user["id"],
                            target_label=user["username"],details="Thất bại: mật khẩu mới dưới 8 ký tự.")
                flash("Mật khẩu mới cần ít nhất 8 ký tự.")
            elif new!=repeat:
                write_audit(con,"Đổi mật khẩu",success=False,target_type="Tài khoản",target_id=user["id"],
                            target_label=user["username"],details="Thất bại: hai mật khẩu mới không trùng nhau.")
                flash("Hai lần nhập mật khẩu mới chưa giống nhau.")
            else:
                new_version=user["auth_version"]+1
                con.execute("UPDATE users SET password_hash=?,auth_version=? WHERE id=?",
                            (generate_password_hash(new),new_version,user["id"]))
                con.execute("UPDATE auth_sessions SET status='revoked',ended_at=? WHERE user_id=? AND status='active' AND session_token_hash<>?",
                            (now(),user["id"],session_token_hash(session.get("session_id", ""))))
                session["auth_version"]=new_version
                write_audit(con,"Đổi mật khẩu",target_type="Tài khoản",target_id=user["id"],
                            target_label=user["username"],details="Người dùng tự đổi mật khẩu; các phiên khác đã bị vô hiệu hóa.")
                flash("Đã đổi mật khẩu.")
                return redirect(url_for("home"))
    return page("""<h1>Đổi mật khẩu</h1><div class="card" style="max-width:480px">
      <form method="post"><label>Mật khẩu hiện tại</label>
      <input type="password" name="old" required>
      <label>Mật khẩu mới (ít nhất 8 ký tự)</label>
      <input type="password" name="new" minlength="8" required>
      <label>Nhập lại mật khẩu mới</label><input type="password" name="repeat" minlength="8" required>
      <p><button>Đổi mật khẩu</button></p></form></div>""")

@app.route("/recover-superadmin",methods=["GET","POST"])
def recover_superadmin():
    if session.get("superadmin_failures",0)<3:
        return redirect(url_for("login"))
    if request.method=="POST":
        code=request.form.get("code","").strip().upper()
        password=request.form.get("password","")
        repeat=request.form.get("repeat","")
        with db() as con:
            stored=con.execute("SELECT value FROM settings WHERE key='superadmin_recovery_hash'").fetchone()
            if not stored or not check_password_hash(stored["value"],code):
                flash("Mã khôi phục không hợp lệ.")
            elif len(password)<8 or password!=repeat:
                flash("Mật khẩu mới cần ít nhất 8 ký tự và hai lần nhập phải khớp.")
            else:
                new_code=recovery_code()
                con.execute("UPDATE users SET password_hash=?,auth_version=auth_version+1 WHERE role='superadmin'",
                            (generate_password_hash(password),))
                con.execute("UPDATE settings SET value=? WHERE key='superadmin_recovery_hash'",
                            (generate_password_hash(new_code),))
                write_audit(con,"Khôi phục Superadmin",target_type="Tài khoản",
                            target_label="superadmin",details="Đặt lại mật khẩu và tạo mã khôi phục mới.")
                session.clear()
                return page("""<div class="card" style="max-width:600px;margin:50px auto">
                <h2>Đã đặt lại mật khẩu Superadmin</h2><p>Mã khôi phục cũ đã hết hiệu lực.
                Ghi và cất mã mới này; mã chỉ hiển thị một lần:</p>
                <p class="serial" style="font-size:22px;font-weight:bold">{{code}}</p>
                <a class="btn" href="{{url_for('login')}}">Đăng nhập</a></div>""",code=new_code)
    return page("""<div class="card" style="max-width:460px;margin:50px auto">
      <h2>Khôi phục tài khoản Superadmin</h2><form method="post">
      <label>Mã khôi phục đã lưu</label><input name="code" required style="width:90%">
      <label>Mật khẩu mới</label><input type="password" name="password" minlength="8" required>
      <label>Nhập lại mật khẩu</label><input type="password" name="repeat" minlength="8" required>
      <p><button>Đặt lại mật khẩu</button></p></form></div>""")

@app.route("/setup-superadmin",methods=["GET","POST"])
@login_required
def setup_superadmin():
    if session["role"]!="admin": return "Chỉ tài khoản admin hiện tại được khởi tạo Superadmin.",403
    with db() as con:
        if con.execute("SELECT 1 FROM users WHERE role='superadmin'").fetchone():
            return "Superadmin đã được khởi tạo.",403
        if request.method=="POST":
            name=request.form.get("username","").strip()
            password=request.form.get("password","")
            admin_password=request.form.get("admin_password","")
            admin=con.execute("SELECT password_hash FROM users WHERE id=?",(session["user_id"],)).fetchone()
            if not check_password_hash(admin["password_hash"],admin_password):
                flash("Mật khẩu admin xác nhận chưa đúng.")
            elif not valid_username(name) or len(password)<8:
                flash("Tên đăng nhập không dấu (1–50 ký tự); mật khẩu Superadmin ít nhất 8 ký tự.")
            else:
                code=recovery_code()
                try:
                    con.execute("""INSERT INTO users(username,password_hash,role,created_at,permissions)
                        VALUES(?,?,?,?,?)""",(name,generate_password_hash(password),"superadmin",now(),json.dumps(ADMIN_PERMISSIONS)))
                    con.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('superadmin_recovery_hash',?)",
                                (generate_password_hash(code),))
                    write_audit(con,"Tạo tài khoản",target_type="Tài khoản",target_label=name,
                                details="Khởi tạo tài khoản Superadmin.",user_id=session["user_id"],
                                username=session["username"],role=session["role"])
                except sqlite3.IntegrityError:
                    flash("Tên đăng nhập đã tồn tại.")
                else:
                    return page("""<div class="card"><h2>Đã tạo Superadmin: {{name}}</h2>
                    <p>Mã khôi phục chỉ hiển thị một lần. Hãy ghi lại và cất ở nơi an toàn:</p>
                    <p class="serial" style="font-size:22px;font-weight:bold">{{code}}</p>
                    <a class="btn" href="{{url_for('users')}}">Đến quản lý tài khoản</a></div>""",name=name,code=code)
    return page("""<h1>Khởi tạo Superadmin</h1><div class="card" style="max-width:550px">
      <p>Tài khoản admin hiện tại xác nhận mật khẩu để tạo tài khoản quyền cao nhất.</p>
      <form method="post"><label>Tên đăng nhập Superadmin</label><input name="username" required maxlength="50">
      <label>Mật khẩu Superadmin (ít nhất 8 ký tự)</label><input type="password" name="password" minlength="8" required>
      <label>Mật khẩu admin hiện tại</label><input type="password" name="admin_password" required>
      <p><button>Tạo Superadmin</button></p></form></div>""")

def valid_username(username):
    return bool(username) and len(username)<=50 and username.isascii() and all(
        c.isalnum() or c in "._-" for c in username)

@app.route("/users",methods=["GET","POST"])
@permission_required("accounts_manage")
def users():
    if request.method=="POST":
        username=request.form.get("username","").strip()
        password=request.form.get("password","")
        role=request.form.get("role","member") if session["role"]=="superadmin" else "member"
        if role not in ("member","admin"): return "Vai trò không hợp lệ.",400
        permissions=(request.form.getlist("permissions") if session["role"]=="superadmin"
                     else MEMBER_PERMISSIONS)
        if any(p not in PERMISSIONS for p in permissions): return "Quyền không hợp lệ.",400
        if not valid_username(username):
            flash("Tên đăng nhập cần 1–50 ký tự không dấu: chữ, số, dấu chấm, gạch dưới hoặc gạch ngang.")
        elif len(password)<8:
            flash("Mật khẩu cần ít nhất 8 ký tự.")
        else:
            try:
                with db() as con:
                    cur=con.execute("""INSERT INTO users(username,password_hash,role,created_at,permissions)
                        VALUES(?,?,?,?,?)""",(username,generate_password_hash(password),role,now(),json.dumps(permissions)))
                    write_audit(con,"Tạo tài khoản",target_type="Tài khoản",target_id=cur.lastrowid,
                                target_label=username,details=f"Vai trò: {role}; Quyền: {', '.join(permissions)}")
                flash(f"Đã tạo tài khoản {username}.")
            except sqlite3.IntegrityError:
                flash("Tên đăng nhập này đã tồn tại.")
        return redirect(url_for("users"))
    with db() as con:
        accounts=[dict(row) for row in con.execute("SELECT * FROM users ORDER BY id")]
        has_superadmin=any(user["role"]=="superadmin" for user in accounts)
    for user in accounts: user["grants"]=json.loads(user["permissions"])
    return page("""<h1>Quản lý tài khoản</h1>
      {% if not has_superadmin and session.role=='admin' %}<div class="card warn">
      Chưa có Superadmin. <a class="btn" href="{{url_for('setup_superadmin')}}">Khởi tạo Superadmin</a></div>{% endif %}
      <div class="card"><h2>Thêm tài khoản</h2>
      <form method="post"><label>Tên đăng nhập</label><input name="username" required maxlength="50">
      <label>Mật khẩu (ít nhất 8 ký tự)</label><input type="password" name="password" minlength="8" required>
      {% if session.role=='superadmin' %}<label>Vai trò</label><select name="role">
      <option value="member">Thành viên</option><option value="admin">Admin</option></select>
      <p>Quyền ban đầu (chọn các chức năng cần dùng):</p>
      {% for key,label in permission_options.items() %}<label style="font-weight:normal">
      <input type="checkbox" name="permissions" value="{{key}}" {% if key in member_defaults %}checked{% endif %}> {{label}}</label>{% endfor %}{% endif %}
      <p><button>Tạo tài khoản</button></p></form></div><div class="card table-wrap">
      <table><tr><th>Tên đăng nhập</th><th>Vai trò / trạng thái</th><th>Quyền và thao tác</th></tr>
      {% for user in accounts if session.role=='superadmin' or user['role']=='member' %}<tr>
      <td>{{user['username']}}</td><td>{{user['role']}} · {{'Đã khóa' if user['disabled'] else 'Đang hoạt động'}}</td><td>
      {% if user['role']!='superadmin' %}
      {% if session.role=='superadmin' %}<form method="post" action="{{url_for('update_user_permissions',user_id=user['id'])}}">
      {% for key,label in permission_options.items() %}<label style="display:inline-block;margin-right:12px;font-weight:normal">
      <input type="checkbox" name="permissions" value="{{key}}" {% if key in user['grants'] %}checked{% endif %}> {{label}}</label>{% endfor %}
      <p><button>Lưu quyền</button></p></form>{% endif %}
      <form method="post" action="{{url_for('reset_member_password',user_id=user['id'])}}" style="display:inline">
      <input type="password" name="password" minlength="8" placeholder="Mật khẩu mới" required>
      <button class="secondary">Đặt lại mật khẩu</button></form>
      {% if session.role=='superadmin' %}<form method="post" action="{{url_for('toggle_user_lock',user_id=user['id'])}}" style="display:inline">
      <button class="secondary">{{'Mở khóa' if user['disabled'] else 'Khóa'}}</button></form>{% endif %}
      {% if session.role=='superadmin' %}<form method="post" action="{{url_for('revoke_user_sessions',user_id=user['id'])}}" style="display:inline"
      onsubmit='return confirm("Đăng xuất tất cả phiên của " + {{user["username"]|tojson}} + "?");'>
      <button class="secondary">Đăng xuất tất cả phiên</button></form>
      <form method="post" action="{{url_for('delete_member',user_id=user['id'])}}" style="display:inline"
      onsubmit='return confirm("Vô hiệu hóa tài khoản " + {{user["username"]|tojson}} + "? Lịch sử hoạt động vẫn được giữ.");'>
      <button class="danger">Vô hiệu hóa</button></form>{% endif %}
      {% endif %}</td></tr>{% endfor %}</table></div>""",
      accounts=accounts,has_superadmin=has_superadmin,permission_options=PERMISSIONS,member_defaults=MEMBER_PERMISSIONS)

@app.route("/users/<int:user_id>/password",methods=["POST"])
@permission_required("accounts_manage")
def reset_member_password(user_id):
    password=request.form.get("password","")
    if len(password)<8:
        flash("Mật khẩu mới cần ít nhất 8 ký tự.")
    else:
        with db() as con:
            target=con.execute("SELECT username,role FROM users WHERE id=?",(user_id,)).fetchone()
            if not target or target["role"]=="superadmin" or (target["role"]=="admin" and session["role"]!="superadmin"):
                return "Không có quyền đặt lại mật khẩu tài khoản này.",403
            target_username=target["username"]
            con.execute("UPDATE users SET password_hash=?,auth_version=auth_version+1 WHERE id=?",
                        (generate_password_hash(password),user_id))
            con.execute("UPDATE auth_sessions SET status='revoked',ended_at=? WHERE user_id=? AND status='active' AND session_token_hash<>?",
                        (now(),user_id,session_token_hash(session.get("session_id", ""))))
            write_audit(con,"Đặt lại mật khẩu",target_type="Tài khoản",target_id=user_id,
                        target_label=target_username,details="Mật khẩu được thay đổi; các phiên khác đã bị vô hiệu hóa.")
        flash("Đã đặt lại mật khẩu; phiên đăng nhập cũ sẽ hết hiệu lực.")
    return redirect(url_for("users"))

@app.route("/users/<int:user_id>/delete",methods=["POST"])
@permission_required("accounts_manage")
def delete_member(user_id):
    with db() as con:
        target=con.execute("SELECT username,role,disabled FROM users WHERE id=?",(user_id,)).fetchone()
        if not target or target["role"]=="superadmin" or (target["role"]=="admin" and session["role"]!="superadmin"):
            return "Không có quyền vô hiệu hóa tài khoản này.",403
        if user_id==session["user_id"]:
            return "Không thể vô hiệu hóa chính tài khoản đang đăng nhập.",400
        con.execute("UPDATE users SET disabled=1,auth_version=auth_version+1 WHERE id=?",(user_id,))
        con.execute("UPDATE auth_sessions SET status='revoked',ended_at=? WHERE user_id=? AND status='active'",(now(),user_id))
        write_audit(con,"Vô hiệu hóa tài khoản",target_type="Tài khoản",target_id=user_id,
                    target_label=target["username"],details="Tài khoản bị vô hiệu hóa; dữ liệu lịch sử vẫn được giữ.")
    flash(f"Đã vô hiệu hóa tài khoản {target['username']}; lịch sử hoạt động vẫn được giữ.")
    return redirect(url_for("users"))

@app.route("/users/<int:user_id>/permissions",methods=["POST"])
@login_required
def update_user_permissions(user_id):
    if session["role"]!="superadmin": return "Chỉ Superadmin được cấp quyền.",403
    permissions=request.form.getlist("permissions")
    if any(p not in PERMISSIONS for p in permissions): return "Quyền không hợp lệ.",400
    with db() as con:
        new_permissions=list(dict.fromkeys(permissions))
        target=con.execute("SELECT username FROM users WHERE id=? AND role<>'superadmin'",(user_id,)).fetchone()
        cur=con.execute("""UPDATE users SET permissions=?,auth_version=auth_version+1
            WHERE id=? AND role<>'superadmin'""",(json.dumps(new_permissions),user_id))
        con.execute("UPDATE auth_sessions SET status='revoked',ended_at=? WHERE user_id=? AND status='active'",(now(),user_id))
        if cur.rowcount and target:
            write_audit(con,"Thay đổi quyền",target_type="Tài khoản",target_id=user_id,
                        target_label=target["username"],details=f"Quyền mới: {', '.join(new_permissions)}")
    flash("Đã cập nhật quyền; tài khoản cần đăng nhập lại." if cur.rowcount else "Không tìm thấy tài khoản cần cấp quyền.")
    return redirect(url_for("users"))

@app.route("/users/<int:user_id>/lock",methods=["POST"])
@login_required
def toggle_user_lock(user_id):
    if session["role"]!="superadmin": return "Chỉ Superadmin được khóa tài khoản.",403
    with db() as con:
        target=con.execute("SELECT username,role,disabled FROM users WHERE id=?",(user_id,)).fetchone()
        if not target or target["role"]=="superadmin" or (target["role"]=="admin" and session["role"]!="superadmin"):
            return "Không có quyền thay đổi tài khoản này.",403
        if user_id==session["user_id"]:
            return "Không thể khóa chính tài khoản đang đăng nhập.",400
        new_disabled=0 if target["disabled"] else 1
        con.execute("UPDATE users SET disabled=?,auth_version=auth_version+1 WHERE id=?",(new_disabled,user_id))
        if new_disabled:
            con.execute("UPDATE auth_sessions SET status='revoked',ended_at=? WHERE user_id=? AND status='active'",(now(),user_id))
        write_audit(con,"Khóa tài khoản" if new_disabled else "Mở khóa tài khoản",target_type="Tài khoản",
                    target_id=user_id,target_label=target["username"],details="Trạng thái tài khoản đã thay đổi.")
    flash("Đã thay đổi trạng thái tài khoản." if target else "Không tìm thấy tài khoản cần khóa.")
    return redirect(url_for("users"))

@app.route("/")
@login_required
def home():
    with db() as con:
        total=con.execute("SELECT count(*) FROM devices WHERE status='Đang sử dụng'").fetchone()[0]
        active=current_round(con)
        round_total=con.execute("SELECT count(*) FROM round_items WHERE round_id=?",
                                (active["id"],)).fetchone()[0] if active else 0
        counted=con.execute("""SELECT count(*) FROM round_items i WHERE i.round_id=?
            AND EXISTS(SELECT 1 FROM scan_once s WHERE s.round_id=i.round_id
                       AND s.device_id=i.device_id)""",
            (active["id"],)).fetchone()[0] if active else 0
        admin_stats=None; recent_audits=[]
        if session.get("role")=="superadmin":
            now_dt=datetime.now(); day_ago=dt_text(now_dt-timedelta(days=1))
            user_total=con.execute("SELECT count(*) FROM users WHERE role<>'superadmin'").fetchone()[0]
            user_active=con.execute("SELECT count(*) FROM users WHERE role<>'superadmin' AND disabled=0").fetchone()[0]
            session_total=con.execute("SELECT count(*) FROM auth_sessions WHERE status='active'").fetchone()[0]
            login_failures=con.execute("SELECT count(*) FROM audit_logs WHERE action='Đăng nhập' AND success=0 AND created_at>=?",(day_ago,)).fetchone()[0]
            audit_day=con.execute("SELECT count(*) FROM audit_logs WHERE created_at>=?",(day_ago,)).fetchone()[0]
            recent_audits=con.execute("""SELECT created_at,username,action,target_label,success,details
                FROM audit_logs ORDER BY id DESC LIMIT 8""").fetchall()
            admin_stats={"user_total":user_total,"user_active":user_active,"session_total":session_total,
                         "login_failures":login_failures,"audit_day":audit_day}
    remaining=max(0,round_total-counted)
    percent=round(counted*100/round_total,1) if round_total else 0
    return page("""<h1>Tổng quan</h1>
      {% if admin_stats %}<div class="card"><h2>Quản trị Superadmin</h2>
      <div class="grid overview-stats">
      <div class="card"><b>Tài khoản</b><div class="stat">{{admin_stats.user_total}}</div><div class="muted">{{admin_stats.user_active}} đang hoạt động</div></div>
      <div class="card"><b>Phiên đang hoạt động</b><div class="stat">{{admin_stats.session_total}}</div></div>
      <div class="card"><b>Đăng nhập thất bại / 24 giờ</b><div class="stat">{{admin_stats.login_failures}}</div></div>
      <div class="card"><b>Nhật ký / 24 giờ</b><div class="stat">{{admin_stats.audit_day}}</div></div></div>
      {% if recent_audits %}<h3>Hoạt động gần đây</h3><div class="table-wrap"><table><tr><th>Thời gian</th><th>Tài khoản</th><th>Hoạt động</th><th>Đối tượng</th><th>Kết quả</th><th>Chi tiết</th></tr>
      {% for x in recent_audits %}<tr><td>{{x['created_at']}}</td><td>{{x['username']}}</td><td>{{x['action']}}</td><td>{{x['target_label']}}</td><td>{{'Thành công' if x['success'] else 'Thất bại'}}</td><td>{{x['details']}}</td></tr>{% endfor %}
      </table></div><p><a class="btn secondary" href="{{url_for('audit_log')}}">Mở toàn bộ nhật ký</a></p>{% endif %}</div>{% endif %}
      {% if active %}<div class="card overview-round"><h2>Đợt đang thực hiện: {{active['name']}}</h2>
      <p class="muted">Các số liệu điểm danh bên dưới thuộc đợt này.</p></div>
      <div class="grid overview-stats">
      <div class="card">Tổng thiết bị trong đợt<div class="stat">{{round_total}}</div></div>
      <div class="card">Đã điểm danh<div class="stat">{{counted}}</div></div>
      <div class="card">Chưa điểm danh<div class="stat">{{remaining}}</div></div>
      <div class="card">Tỷ lệ hoàn thành<div class="stat">{{percent}}%</div></div></div>
      <div class="card"><label for="overview-progress">Tiến độ kiểm kê: {{counted}}/{{round_total}}</label>
      <div id="overview-progress" class="progress-track" role="progressbar" aria-valuenow="{{percent}}" aria-valuemin="0" aria-valuemax="100" aria-label="Tiến độ kiểm kê">
      <div class="progress-fill" style="width:{{percent}}%"></div></div>
      {% if can('results_view') %}<p><a class="btn secondary" href="{{url_for('results',round_id=active['id'])}}">Xem kết quả</a></p>{% endif %}</div>
      <p class="muted">Danh mục hiện hành: {{total}} thiết bị.{% if total != round_total %} Tổng thiết bị của đợt được tính theo danh sách của đợt kiểm kê.{% endif %}</p>
      {% else %}<div class="grid overview-stats"><div class="card">Thiết bị hiện hành<div class="stat">{{total}}</div></div></div>
      <div class="card"><h2>Chưa có đợt kiểm kê đang thực hiện</h2>
      {% if can('round_manage') %}<p><a class="btn" href="{{url_for('rounds')}}">Xem các đợt kiểm kê</a></p>{% endif %}</div>{% endif %}""",
      total=total,active=active,round_total=round_total,counted=counted,
      remaining=remaining,percent=percent,admin_stats=admin_stats,recent_audits=recent_audits)

def pagination(total, requested, size=100):
    pages=max(1,(total+size-1)//size)
    number=min(max(1,requested or 1),pages)
    return number,pages,size,(number-1)*size

DEVICES_VIEW = """<h1>Danh mục thiết bị</h1>
{% if can('devices_manage') %}<div class="card" id="add-device"><h2>Thêm thiết bị thủ công</h2><form method="post" class="grid device-fields">
<div><label>STT</label><input name="stt" value="{{form_data.get('stt',suggested_stt)}}" inputmode="numeric" pattern="[0-9]+" required class="{{'invalid-field' if 'stt' in form_errors else ''}}">{% if 'stt' in form_errors %}<div class="field-error">{{form_errors['stt']}}</div>{% endif %}</div>
<div><label>Mã tài sản</label><input name="asset_code" value="{{form_data.get('asset_code','')}}" required></div>
<div><label>Tên VT - TB</label><input name="name" value="{{form_data.get('name','')}}" required></div>
<div><label>Số seri</label><input name="serial" value="{{form_data.get('serial','')}}" class="{{'invalid-field' if 'serial' in form_errors else ''}}">{% if 'serial' in form_errors %}<div class="field-error">{{form_errors['serial']}}</div>{% endif %}</div>
{% for field,label in extra_pairs %}<div><label>{{label}}</label><input name="{{field}}" value="{{form_data.get(field,'')}}"></div>{% endfor %}
<div class="device-submit"><button>Thêm thiết bị</button></div></form></div>{% endif %}
{% if can('excel_import') %}<div class="card" id="import"><h2>Thêm từ danh sách Excel</h2>
<p>Đọc sheet <b>TỔNG TÀI SẢN</b>, cột A–K (gồm Model, Hãng, Nước SX, các năm và Khoa SD).
Tiêu đề dòng 7, dữ liệu từ dòng 9. <a class="btn secondary" href="{{url_for('excel_template')}}">Tải file mẫu</a>.</p>
<form method="post" action="{{url_for('import_excel')}}" enctype="multipart/form-data" id="upload-form">
<label>File Excel</label><input type="file" name="file" accept=".xlsx" required>
<label>Cách nhập</label><select name="mode" id="mode">
<option value="add">Cập nhật danh mục: chỉ thêm seri mới</option>
<option value="replace">Thay hoàn toàn danh mục hiện hành</option></select>
<input type="hidden" name="confirm_replace" id="confirm_replace" value="">
<p class="muted">Thay danh mục sẽ đóng đợt đang thực hiện. Báo cáo đợt cũ vẫn xem được.</p>
<button>Đọc và nhập file</button></form>
{% if can('devices_manage') %}<form method="post" action="{{url_for('clear_devices')}}" onsubmit="return confirm('Xóa toàn bộ {{catalog_total}} thiết bị hiện hành và đóng đợt đang thực hiện? Dữ liệu các đợt cũ vẫn được giữ.');">
<button class="danger" {% if not catalog_total %}disabled{% endif %}>Xóa toàn bộ danh mục</button></form>{% endif %}</div>
<script>
document.getElementById('upload-form').addEventListener('submit',function(e){
 if(document.getElementById('mode').value==='replace'){
   if(!window.confirm('Thay toàn bộ {{catalog_total}} thiết bị hiện hành bằng danh sách trong file mới? Đợt đang thực hiện sẽ đóng.')){
     e.preventDefault();return;
   }
   document.getElementById('confirm_replace').value='THAY DANH MUC';
 }
});
</script>
{% elif can('devices_manage') %}<div class="card"><form method="post" action="{{url_for('clear_devices')}}" onsubmit="return confirm('Xóa toàn bộ {{catalog_total}} thiết bị hiện hành?');">
<button class="danger" {% if not catalog_total %}disabled{% endif %}>Xóa toàn bộ danh mục</button></form></div>
{% endif %}
<div class="card"><form method="get"><input type="hidden" name="filter" value="{{device_filter}}"><input name="q" value="{{q}}" placeholder="Tìm mã, tên, seri">
<button>Tìm</button></form>
<nav class="filter-links" aria-label="Lọc thiết bị"><a href="{{url_for('devices')}}" {% if not device_filter %}aria-current="page"{% endif %}>Tất cả ({{catalog_total}})</a>
<a href="{{url_for('devices',filter='missing')}}" {% if device_filter=='missing' %}aria-current="page"{% endif %}>Thiếu seri ({{missing_count}})</a>
<a href="{{url_for('devices',filter='duplicate')}}" {% if device_filter=='duplicate' %}aria-current="page"{% endif %}>Seri trùng ({{duplicate_count}})</a></nav>
{% if can('devices_view') %}<div class="filter-links">
<a class="btn" href="{{url_for('catalog_issues_excel',kind='missing')}}">Xuất danh sách thiếu seri</a>
<a class="btn" href="{{url_for('catalog_issues_excel',kind='duplicate')}}">Xuất danh sách trùng seri</a></div>{% endif %}
<p>Hiển thị {{total}} thiết bị; trang {{number}}/{{pages}}. {% if can('qr_export') %}Chọn thiết bị trên trang này để xuất tem QR. Thiết bị thiếu hoặc trùng seri được bỏ qua khi xuất QR.{% endif %}</p>
{% if can('qr_export') %}<form method="post" action="{{url_for('labels')}}" id="labels-form">
<input type="hidden" name="q" value="{{q}}">
<input type="hidden" name="filter" value="{{device_filter}}">{% endif %}
<div class="table-wrap"><table class="devices-table"><tr>{% if can('qr_export') %}<th>Chọn</th>{% endif %}<th>STT</th><th>Mã tài sản</th>
<th>Tên VT - TB</th><th>Số seri</th><th>Thao tác</th></tr>
{% for d in rows %}<tr>
{% if can('qr_export') %}<td><input type="checkbox" name="selected" value="{{d['id']}}" {% if not d['serial_key'] or d['serial_count']>1 %}disabled{% endif %}></td>{% endif %}
<td>{{d['stt']}}</td><td>{{d['asset_code']}}</td><td class="table-name">{{d['name']}}</td>
<td class="serial">{{d['serial']}}</td><td><div class="device-actions">
{% if can('qr_export') and d['serial_key'] and d['serial_count']==1 %}<a class="btn" href="{{url_for('qr_view',device_id=d['id'])}}">Xem QR</a>{% endif %}
{% if can('devices_manage') %}<a class="btn" href="{{url_for('edit_device',device_id=d['id'])}}">Sửa</a>
<button type="submit" class="danger" form="delete-{{d['id']}}">Xóa</button>{% endif %}</div>
{% if not d['serial_key'] %}<div class="device-warning">Thiếu seri</div>
{% elif d['serial_count']>1 %}<div class="device-warning">Seri trùng với:
  {% for other in duplicate_details.get(d['serial_key'],[]) if other['id']!=d['id'] %}
  {% if can('devices_manage') %}<a href="{{url_for('edit_device',device_id=other['id'])}}">{% endif %}STT {{other['stt'] or '—'}}, MTS {{other['asset_code']}}{% if can('devices_manage') %}</a>{% endif %}{% if not loop.last %}; {% endif %}{% endfor %}</div>{% endif %}</td></tr>{% endfor %}
</table></div>
{% if can('qr_export') %}
<input type="hidden" name="scope" id="label-scope" value="all">
<div class="pager"><button type="submit" data-scope="selected">Xuất mã QR đã chọn</button>
<button type="submit" data-scope="all" class="secondary">Xuất toàn bộ mã QR</button></div>
</form>
<div id="labels-busy" class="busy-layer" role="status" aria-live="assertive"><div class="busy-box">
<div class="spinner"></div><b>Đang tạo file Word chứa mã QR…</b>
<p class="muted">Danh mục lớn có thể cần chờ một lúc. File sẽ tự tải xuống khi hoàn tất.</p>
</div></div>
<script>
document.getElementById('labels-form').addEventListener('submit',async function(event){
 event.preventDefault();
 var form=this;
 var clicked=event.submitter;
 document.getElementById('label-scope').value=clicked && clicked.dataset.scope==='selected' ? 'selected' : 'all';
 var busy=document.getElementById('labels-busy');
 var buttons=form.querySelectorAll('button[data-scope]');
 busy.classList.add('show');
 buttons.forEach(function(button){button.disabled=true;});
 try{
   var response=await fetch(form.action,{method:'POST',body:new FormData(form),credentials:'same-origin',headers:{'X-Requested-With':'XMLHttpRequest'}});
   var type=response.headers.get('content-type')||'';
   if(!response.ok||type.indexOf('application/vnd.openxmlformats-officedocument.wordprocessingml.document')<0){
     busy.classList.remove('show');
     window.location.href=response.url;
     return;
   }
   var blob=await response.blob();
   var fileName='Tem_QR_thiet_bi.docx';
   var disposition=response.headers.get('content-disposition')||'';
   var match=disposition.match(/filename[^;=\\n]*=(?:UTF-8''|\")?([^;\"\\n]+)/i);
   if(match&&match[1])fileName=decodeURIComponent(match[1].trim());
   var link=document.createElement('a');link.href=URL.createObjectURL(blob);link.download=fileName;
   document.body.appendChild(link);link.click();link.remove();URL.revokeObjectURL(link.href);
   busy.classList.remove('show');
 }catch(error){
   busy.classList.remove('show');
   alert('Không thể tạo file Word. Hãy kiểm tra kết nối với chương trình rồi thử lại.');
 }finally{
   buttons.forEach(function(button){button.disabled=false;});
 }
});
</script>
{% endif %}
{% if can('devices_manage') %}{% for d in rows %}<form id="delete-{{d['id']}}" method="post"
 action="{{url_for('delete_device',device_id=d['id'])}}"
 onsubmit='return confirm("Xóa thiết bị " + {{d["asset_code"]|tojson}} + " khỏi danh mục hiện hành?");'></form>{% endfor %}{% endif %}
<div class="pager">
{% if number>1 %}<a class="btn secondary" href="{{url_for('devices',q=q,filter=device_filter,page=1)}}">Đầu</a>
<a class="btn secondary" href="{{url_for('devices',q=q,filter=device_filter,page=number-1)}}">← Trước</a>{% endif %}
<span>Trang {{number}}/{{pages}}</span>
{% if number<pages %}<a class="btn secondary" href="{{url_for('devices',q=q,filter=device_filter,page=number+1)}}">Tiếp →</a>
<a class="btn secondary" href="{{url_for('devices',q=q,filter=device_filter,page=pages)}}">Cuối</a>{% endif %}
</div></div>"""

@app.route("/devices",methods=["GET","POST"])
@permission_required("devices_view")
def devices():
    form_data={}
    form_errors={}
    if request.method=="POST":
        if not has_permission("devices_manage"): return "Không có quyền thêm thiết bị.",403
        form_data=request.form.to_dict()
        stt=cell_text(request.form.get("stt"))
        code=cell_text(request.form.get("asset_code"))
        name=cell_text(request.form.get("name"))
        serial=cell_text(request.form.get("serial"))
        key=serial_key(serial)
        extras=extra_values(request.form)
        if not stt.isdigit() or not code or not name:
            flash("Cần nhập STT dạng số, mã tài sản và tên thiết bị.")
        else:
            with db() as con:
                stt_exists=con.execute("""SELECT stt,asset_code,name FROM devices WHERE status='Đang sử dụng'
                    AND stt=?""",(stt,)).fetchone()
                exists=con.execute("""SELECT stt,asset_code,name FROM devices WHERE status='Đang sử dụng'
                    AND ((serial_key=? AND ?<>'') OR
                         (serial_key='' AND ?='' AND asset_code=?))""",
                    (key,key,key,code)).fetchone()
                if stt_exists:
                    form_errors["stt"]=(f"STT {stt} đã thuộc thiết bị {stt_exists['name']} "
                                        f"(MTS {stt_exists['asset_code']}).")
                if exists:
                    form_errors["serial"]=(f"Seri {serial} đã thuộc STT {exists['stt']}, "
                                           f"thiết bị {exists['name']} (MTS {exists['asset_code']})."
                                           if key else f"MTS {code} thiếu seri đã thuộc STT {exists['stt']}, "
                                                        f"thiết bị {exists['name']}.")
                if not form_errors:
                    cur=con.execute("""INSERT INTO devices(stt,asset_code,name,serial,serial_key,
                        model,maker,country,year_made,year_used,year_increase,department,created_at)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",(stt,code,name,serial,key,*extras,now()))
                    r=current_round(con)
                    if r: add_round_item(con,r["id"],cur.lastrowid,stt,code,name,serial,key,extras)
                    write_audit(con,"Thêm thiết bị",target_type="Thiết bị",target_id=cur.lastrowid,
                                target_label=code,details=f"Tên: {name}; Seri: {serial or 'Thiếu seri'}")
                    flash("Đã thêm thiết bị." + (" Cần bổ sung seri trước khi tạo QR." if not key else ""))
                    return redirect(url_for("devices"))
        if form_errors:
            flash("Thông tin bị trùng. Các ô vừa nhập đã được giữ lại; hãy sửa ô được đánh dấu.")
    q=request.args.get("q","").strip()
    device_filter=request.args.get("filter","")
    if device_filter not in ("", "missing", "duplicate"):
        device_filter=""
    filter_sql={"":"", "missing":" AND d.serial_key=''", "duplicate":""" AND d.serial_key<>''
        AND (SELECT count(*) FROM devices x WHERE x.status='Đang sử dụng'
             AND x.serial_key=d.serial_key)>1"""}[device_filter]
    with db() as con:
        term=f"%{q}%"
        catalog_total=con.execute("SELECT count(*) FROM devices WHERE status='Đang sử dụng'").fetchone()[0]
        total=con.execute(f"""SELECT count(*) FROM devices d WHERE d.status='Đang sử dụng'
            AND (d.stt||d.asset_code||d.name||d.serial) LIKE ? {filter_sql}""",(term,)).fetchone()[0]
        missing_count=con.execute("""SELECT count(*) FROM devices
            WHERE status='Đang sử dụng' AND serial_key=''""").fetchone()[0]
        duplicate_count=con.execute("""SELECT count(*) FROM devices d
            WHERE d.status='Đang sử dụng' AND d.serial_key<>'' AND
            (SELECT count(*) FROM devices x WHERE x.status='Đang sử dụng'
             AND x.serial_key=d.serial_key)>1""").fetchone()[0]
        number,pages,size,offset=pagination(total,request.args.get("page",1,type=int),200)
        suggested_stt=next_stt(con)
        rows=con.execute(f"""SELECT d.*, (SELECT count(*) FROM devices x
            WHERE x.status='Đang sử dụng' AND x.serial_key=d.serial_key) serial_count
            FROM devices d WHERE d.status='Đang sử dụng'
            AND (d.stt||d.asset_code||d.name||d.serial) LIKE ?
            {filter_sql} ORDER BY d.id LIMIT ? OFFSET ?""",(term,size,offset)).fetchall()
        duplicate_keys={d["serial_key"] for d in rows if d["serial_key"] and d["serial_count"]>1}
        duplicate_details={key:con.execute("""SELECT id,stt,asset_code FROM devices
            WHERE status='Đang sử dụng' AND serial_key=? ORDER BY id""",(key,)).fetchall()
            for key in duplicate_keys}
    return page(DEVICES_VIEW,rows=rows,q=q,total=total,catalog_total=catalog_total,
                device_filter=device_filter,number=number,pages=pages,
                missing_count=missing_count,duplicate_count=duplicate_count,
                suggested_stt=suggested_stt,duplicate_details=duplicate_details,
                extra_pairs=zip(EXTRA_FIELDS,EXTRA_LABELS),
                form_data=form_data,form_errors=form_errors)

@app.route("/devices/issues/<kind>.xlsx")
@permission_required("devices_view")
def catalog_issues_excel(kind):
    if kind not in ("missing","duplicate"): return "Loại danh sách không hợp lệ",404
    with db() as con:
        if kind=="missing":
            rows=con.execute("""SELECT * FROM devices WHERE status='Đang sử dụng'
                AND serial_key='' ORDER BY id""").fetchall()
            references={}
        else:
            rows=con.execute("""SELECT d.* FROM devices d WHERE d.status='Đang sử dụng'
                AND d.serial_key<>'' AND EXISTS (
                  SELECT 1 FROM devices x WHERE x.status='Đang sử dụng'
                  AND x.serial_key=d.serial_key AND x.id<>d.id)
                ORDER BY d.serial_key,d.id""").fetchall()
            references={key:con.execute("""SELECT id,stt,asset_code FROM devices
                WHERE status='Đang sử dụng' AND serial_key=? ORDER BY id""",(key,)).fetchall()
                for key in {row['serial_key'] for row in rows}}
    wb=Workbook();ws=wb.active
    ws.title="Thiếu seri" if kind=="missing" else "Trùng seri"
    append_excel_values(ws,[*EXCEL_HEADERS,"Đối chiếu trùng với"])
    for cell in ws[1]: cell.font=Font(bold=True)
    for row in rows:
        peers=references.get(row['serial_key'],[])
        refs="; ".join(f"STT {peer['stt'] or '—'}, MTS {peer['asset_code']}"
                       for peer in peers if peer['id']!=row['id'])
        append_excel_values(ws,[row['stt'],row['asset_code'],row['name'],row['serial'],
            *[row[field] for field in EXTRA_FIELDS],refs])
    for column in "ABCDEFGHIJKL": ws.column_dimensions[column].width=22
    ws.column_dimensions["C"].width=45
    ws.column_dimensions["L"].width=58
    ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
    output=io.BytesIO();wb.save(output);output.seek(0)
    filename="Danh_sach_thieu_seri.xlsx" if kind=="missing" else "Danh_sach_trung_seri.xlsx"
    with db() as con:
        write_audit(con,"Xuất danh sách rà soát",target_type="Danh mục thiết bị",target_label=filename,
                    details=f"Xuất {len(rows)} dòng {kind}.")
    return send_file(output,as_attachment=True,download_name=filename)

@app.route("/devices/clear",methods=["POST"])
@permission_required("devices_manage")
def clear_devices():
    with db() as con:
        count=con.execute("SELECT count(*) FROM devices WHERE status='Đang sử dụng'").fetchone()[0]
        if count:
            con.execute("UPDATE devices SET status='Đã xóa' WHERE status='Đang sử dụng'")
            con.execute("UPDATE inventory_rounds SET active=0 WHERE active=1")
            write_audit(con,"Xóa toàn bộ danh mục",target_type="Danh mục thiết bị",
                        details=f"Đã đánh dấu {count} thiết bị hiện hành là đã xóa và đóng đợt đang làm.")
    flash(f"Đã xóa {count} thiết bị khỏi danh mục hiện hành. Đợt đang làm đã đóng; các báo cáo cũ được giữ.")
    return redirect(url_for("devices"))

@app.route("/devices/<int:device_id>/edit",methods=["GET","POST"])
@permission_required("devices_manage")
def edit_device(device_id):
    with db() as con:
        d=con.execute("SELECT * FROM devices WHERE id=? AND status='Đang sử dụng'",
                      (device_id,)).fetchone()
        if not d: return "Không tìm thấy thiết bị hiện hành",404
        if request.method=="POST":
            stt=cell_text(request.form.get("stt"))
            code=cell_text(request.form.get("asset_code"))
            name=cell_text(request.form.get("name"))
            serial=cell_text(request.form.get("serial"))
            key=serial_key(serial)
            extras=extra_values(request.form)
            stt_exists=con.execute("""SELECT 1 FROM devices
                WHERE status='Đang sử dụng' AND stt=? AND id<>?""",
                (stt,device_id)).fetchone()
            if not stt.isdigit() or not code or not name:
                flash("Cần nhập STT dạng số, mã tài sản và tên thiết bị.")
            elif stt_exists:
                flash(f"STT {stt} đã có ở thiết bị khác. Chưa lưu thay đổi.")
            else:
                serial_changed=key!=d["serial_key"]
                con.execute("""UPDATE devices SET stt=?,asset_code=?,name=?,serial=?,serial_key=?,
                    model=?,maker=?,country=?,year_made=?,year_used=?,year_increase=?,department=?
                    WHERE id=?""",(stt,code,name,serial,key,*extras,device_id))
                r=current_round(con)
                if r:
                    # Cập nhật bản kiểm kê đang làm nếu thiết bị chưa được quét.
                    already=con.execute("""SELECT 1 FROM scans WHERE round_id=? AND device_id=?""",
                                        (r["id"],device_id)).fetchone()
                    if not already:
                        con.execute("""UPDATE round_items SET stt=?,asset_code=?,name=?,serial=?,serial_key=?,
                            model=?,maker=?,country=?,year_made=?,year_used=?,year_increase=?,department=?
                            WHERE round_id=? AND device_id=?""",
                            (stt,code,name,serial,key,*extras,r["id"],device_id))
                old_serial=d["serial"]
                write_audit(con,"Sửa thiết bị",target_type="Thiết bị",target_id=device_id,
                            target_label=code,details=f"Tên: {name}; Seri: {serial or 'Thiếu seri'}; Seri cũ: {old_serial or 'Thiếu seri'}")
                flash("Đã lưu thông tin thiết bị.")
                if serial_changed:
                    flash("Số seri đã thay đổi. Mã QR cũ không còn sử dụng được; hãy in lại mã QR mới cho thiết bị này.")
                    if r and already:
                        flash("Thiết bị đã điểm danh trong đợt hiện tại; báo cáo đợt này giữ seri cũ để lưu lịch sử.")
                return redirect(url_for("devices",q=code))
    return page("""<h1>Sửa thiết bị</h1><div class="card"><form method="post">
      <label>STT</label><input name="stt" value="{{d['stt']}}" inputmode="numeric" pattern="[0-9]+" required>
      <label>Mã tài sản</label><input name="asset_code" value="{{d['asset_code']}}" required>
      <label>Tên VT - TB</label><input name="name" value="{{d['name']}}" required style="width:80%">
      <label>Số seri</label><input name="serial" value="{{d['serial']}}" style="width:80%">
      {% for field,label in extra_pairs %}<label>{{label}}</label>
      <input name="{{field}}" value="{{d[field]}}" style="width:80%">{% endfor %}
      <p class="muted">Nếu seri trùng với thiết bị khác, QR của hai thiết bị sẽ tạm khóa.</p>
      <button>Lưu</button> <a class="btn secondary" href="{{url_for('devices')}}">Hủy</a>
      </form></div>""",d=d,extra_pairs=zip(EXTRA_FIELDS,EXTRA_LABELS))

@app.route("/devices/<int:device_id>/delete",methods=["POST"])
@permission_required("devices_manage")
def delete_device(device_id):
    with db() as con:
        d=con.execute("SELECT * FROM devices WHERE id=? AND status='Đang sử dụng'",
                      (device_id,)).fetchone()
        if not d: return "Không tìm thấy thiết bị hiện hành",404
        r=current_round(con)
        scanned=bool(r and con.execute("""SELECT 1 FROM scans
            WHERE round_id=? AND device_id=?""",(r["id"],device_id)).fetchone())
        con.execute("UPDATE devices SET status='Đã xóa' WHERE id=?",(device_id,))
        if r and not scanned:
            con.execute("DELETE FROM round_items WHERE round_id=? AND device_id=?",
                        (r["id"],device_id))
        write_audit(con,"Xóa thiết bị",target_type="Thiết bị",target_id=device_id,
                    target_label=d["asset_code"],details=f"Tên: {d['name']}; Seri: {d['serial'] or 'Thiếu seri'}")
    flash("Đã xóa thiết bị khỏi danh mục hiện hành."
          + (" Đợt hiện tại giữ kết quả đã quét để bảo toàn lịch sử." if scanned else ""))
    return redirect(url_for("devices"))

def qr_image(value):
    output=io.BytesIO()
    qrcode.make(value,error_correction=qrcode.constants.ERROR_CORRECT_M).save(output,format="PNG")
    output.seek(0)
    return output

def qr_device(device_id):
    with db() as con:
        d=con.execute("SELECT * FROM devices WHERE id=? AND status='Đang sử dụng'",
                      (device_id,)).fetchone()
        if not d: return None,"Không tìm thấy thiết bị hiện hành"
        if not d["serial_key"]: return None,"Thiết bị thiếu seri"
        count=con.execute("""SELECT count(*) FROM devices WHERE status='Đang sử dụng'
            AND serial_key=?""",(d["serial_key"],)).fetchone()[0]
        if count>1: return None,"Seri trùng với thiết bị khác"
    return d,None

@app.route("/qr/<int:device_id>")
@permission_required("qr_export")
def qr_view(device_id):
    d,error=qr_device(device_id)
    if error: return error,409
    return page("""<div class="card" style="max-width:430px;text-align:center;margin:auto">
      <img src="{{url_for('qr_png',device_id=d['id'])}}" alt="Mã QR" style="width:290px;max-width:100%">
      <div class="serial" style="margin:8px;overflow-wrap:anywhere">Seri: {{d['serial']}}</div>
      <div class="serial" style="margin:8px;overflow-wrap:anywhere">MTS: {{d['asset_code']}}</div>
      <div class="serial" style="margin:8px;overflow-wrap:anywhere">Khoa: {{d['department']}}</div>
      <button onclick="window.print()">In mã này</button></div>
      <style>@media print{nav,.wrap>.flash,button,.site-footer{display:none!important}body{background:white}
      .card{box-shadow:none!important}}</style>""",d=d)

@app.route("/qr/<int:device_id>.png")
@permission_required("qr_export")
def qr_png(device_id):
    d,error=qr_device(device_id)
    if error: return error,409
    return send_file(qr_image(d["serial_key"]),mimetype="image/png",
                     download_name=f"QR_{device_id}.png")

@app.route("/labels",methods=["POST"])
@permission_required("qr_export")
def labels():
    # Nút xuất QR trước đây gửi trực tiếp giá trị của nút. Một số trình duyệt
    # có thể không gửi giá trị này khi lớp chờ được bật, vì vậy biểu mẫu dùng
    # một trường ẩn; thiếu giá trị thì an toàn hiểu là xuất toàn bộ.
    scope=request.form.get("scope") or "all"
    q=request.form.get("q","").strip()
    device_filter=request.form.get("filter","")
    if device_filter not in ("missing","duplicate"): device_filter=""
    with db() as con:
        if scope=="selected":
            ids=list(dict.fromkeys(x for x in request.form.getlist("selected") if x.isdigit()))
            if not ids:
                flash("Hãy chọn ít nhất một thiết bị có QR trên trang này.")
                return redirect(url_for("devices",q=q,filter=device_filter))
            placeholders=",".join("?" for _ in ids)
            rows=con.execute(f"""SELECT * FROM devices WHERE status='Đang sử dụng'
                AND id IN ({placeholders}) ORDER BY id""",ids).fetchall()
        elif scope=="all":
            rows=con.execute("SELECT * FROM devices WHERE status='Đang sử dụng' ORDER BY id").fetchall()
        else: return "Lựa chọn không hợp lệ",400
        counts=Counter(r["serial_key"] for r in con.execute(
            "SELECT serial_key FROM devices WHERE status='Đang sử dụng' AND serial_key<>''"))
    valid=[d for d in rows if d["serial_key"] and counts[d["serial_key"]]==1]
    missing=sum(not d["serial_key"] for d in rows)
    duplicate=len(rows)-len(valid)-missing
    if not valid:
        flash(f"Không có QR hợp lệ để xuất. Thiếu seri: {missing}; seri trùng: {duplicate}.")
        return redirect(url_for("devices",q=q,filter=device_filter))
    doc=Document()
    sec=doc.sections[0]
    sec.page_width=Cm(21);sec.page_height=Cm(29.7)
    sec.top_margin=sec.bottom_margin=Cm(0.8)
    sec.left_margin=sec.right_margin=Cm(0.8)
    for begin in range(0,len(valid),20):
        if begin: doc.add_page_break()
        group=valid[begin:begin+20]
        table=doc.add_table(rows=5,cols=4)
        table.style="Table Grid"
        table.autofit=False
        for row in table.rows:
            row.height=Cm(5.25)
            row.height_rule=WD_ROW_HEIGHT_RULE.EXACTLY
            for cell in row.cells:
                cell.width=Cm(4.85)
                cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for i,d in enumerate(group):
            cell=table.cell(i//4,i%4)
            p=cell.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before=p.paragraph_format.space_after=Pt(0)
            qr_width=3.4 if len(d["serial_key"])>65 else 3.7
            p.add_run().add_picture(qr_image(d["serial_key"]),width=Cm(qr_width))
            caption=cell.add_paragraph()
            caption.alignment=WD_ALIGN_PARAGRAPH.CENTER
            caption.paragraph_format.space_before=caption.paragraph_format.space_after=Pt(0)
            run=caption.add_run("Seri: "+d["serial_key"])
            run.font.size=Pt(5 if len(d["serial_key"])>150 else 7 if len(d["serial_key"])>65 else 9)
            run.add_break()
            code_run=caption.add_run("MTS: "+d["asset_code"])
            code_run.font.size=Pt(7 if len(d["asset_code"])>35 else 9)
            code_run.add_break()
            department_run=caption.add_run("Khoa: "+d["department"])
            department_run.font.size=Pt(7 if len(d["department"])>35 else 9)
    output=io.BytesIO();doc.save(output);output.seek(0)
    flash(f"Đã tạo {len(valid)} tem; bỏ qua {missing} thiếu seri và {duplicate} seri trùng.")
    with db() as con:
        write_audit(con,"Xuất mã QR",target_type="QR",target_label="Tem_QR_thiet_bi.docx",details=f"Đã tạo {len(valid)} tem QR.")
    return send_file(output,as_attachment=True,
                     download_name="Tem_QR_thiet_bi.docx",
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

@app.route("/excel-template")
@permission_required("excel_import")
def excel_template():
    wb=Workbook();ws=wb.active;ws.title="TỔNG TÀI SẢN"
    for _ in range(6): ws.append([])
    ws.append(EXCEL_HEADERS)
    for cell in ws[7]: cell.font=Font(bold=True)
    ws.append(list(range(1,12)))
    ws.append([1,"TSTB00001","Máy đo SPO2","SERIAL-001","BT-710",
               "Bistos","Korea","2021","2021","2021","HHTM"])
    for col,width in zip("ABCDEFGHIJK",[10,20,42,35,22,22,20,16,16,20,22]):
        ws.column_dimensions[col].width=width
    output=io.BytesIO();wb.save(output);output.seek(0)
    with db() as con:
        write_audit(con,"Xuất file mẫu Excel",target_type="File mẫu",target_label="Mau_nhap_thiet_bi.xlsx",details="Tải mẫu nhập danh mục.")
    return send_file(output,as_attachment=True,download_name="Mau_nhap_thiet_bi.xlsx")

def read_inventory(file):
    wb=load_workbook(file,read_only=True,data_only=True)
    try:
        if "TỔNG TÀI SẢN" not in wb.sheetnames:
            raise ValueError("Không thấy sheet TỔNG TÀI SẢN")
        ws=wb["TỔNG TÀI SẢN"]
        header=[" ".join(cell_text(x).upper().split()) for x in next(
            ws.iter_rows(min_row=7,max_row=7,max_col=11,values_only=True))]
        if header!=EXCEL_HEADERS:
            raise ValueError("Dòng 7 cần đủ tiêu đề A–K đúng như file mẫu, kết thúc bằng KHOA SD")
        items=[]
        for source_row,values in enumerate(ws.iter_rows(min_row=9,max_col=11,values_only=True),9):
            stt,code,name,serial,*extras=map(cell_text,values)
            if not code and not name: continue
            if not code or not name:
                raise ValueError(f"Dòng {source_row} thiếu mã tài sản hoặc tên thiết bị; chưa nhập dữ liệu")
            items.append((source_row,stt,code,name,serial,serial_key(serial),extras))
        if not items: raise ValueError("Không có thiết bị hợp lệ từ dòng 9")
        return items
    finally:
        wb.close()

@app.route("/import",methods=["GET","POST"])
@permission_required("excel_import")
def import_excel():
    if request.method=="POST":
        file=request.files.get("file")
        mode=request.form.get("mode","add")
        if mode not in ("add","replace"): return "Chế độ nhập không hợp lệ",400
        if mode=="replace" and request.form.get("confirm_replace")!="THAY DANH MUC":
            flash("Chưa xác nhận thay danh mục.");return redirect(url_for("devices")+"#import")
        if not file or not file.filename.lower().endswith(".xlsx"):
            flash("Hãy chọn file .xlsx.");return redirect(url_for("devices")+"#import")
        try:
            items=read_inventory(file)
            with db() as con:
                old_count=con.execute("SELECT count(*) FROM devices WHERE status='Đang sử dụng'").fetchone()[0]
                initial=(mode=="add" and old_count==0)
                active_devices={}
                for existing in con.execute("""SELECT id,serial_key,stt,asset_code FROM devices
                    WHERE status='Đang sử dụng' AND serial_key<>'' ORDER BY id"""):
                    active_devices.setdefault(existing["serial_key"],[]).append(existing)
                blank_codes={row[0] for row in con.execute(
                    "SELECT asset_code FROM devices WHERE status='Đang sử dụng' AND serial_key=''")}
                counts=Counter(item[5] for item in items if item[5])
                file_lines={}
                for source_row,stt,code,name,serial,key,extras in items:
                    if key: file_lines.setdefault(key,[]).append(source_row)
                seen={}
                issues=[]
                added=skipped=missing=duplicate=0
                if mode=="replace":
                    con.execute("UPDATE devices SET status='Đã thay thế' WHERE status='Đang sử dụng'")
                    con.execute("UPDATE inventory_rounds SET active=0 WHERE active=1")
                r=current_round(con) if mode=="add" else None
                for source_row,stt,code,name,serial,key,extras in items:
                    reasons=[]
                    if not key:
                        missing+=1;reasons.append("Thiếu seri")
                    elif counts[key]>1:
                        others=[str(line) for line in file_lines[key] if line!=source_row]
                        reasons.append("Trùng seri với dòng Excel " + ", ".join(others))
                    should_skip=False
                    if mode=="add" and not initial:
                        if key and (key in active_devices or key in seen):
                            should_skip=True
                            if key in active_devices:
                                refs=", ".join(f"STT {d['stt'] or '—'}, MTS {d['asset_code']}"
                                               for d in active_devices[key])
                                reasons.append("Đã có seri trong danh mục: "+refs)
                            elif counts[key]==1:
                                reasons.append(f"Trùng seri với dòng Excel {seen[key]}")
                        elif not key and code in blank_codes:
                            should_skip=True;reasons.append("Thiếu seri; mã tài sản đã có")
                    if key and (counts[key]>1 or key in active_devices and mode=="add" and not initial):
                        duplicate+=1
                    if should_skip:
                        skipped+=1
                    else:
                        cur=con.execute("""INSERT INTO devices
                            (stt,asset_code,name,serial,serial_key,model,maker,country,
                             year_made,year_used,year_increase,department,created_at)
                            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",(stt,code,name,serial,key,*extras,now()))
                        added+=1
                        if r: add_round_item(con,r["id"],cur.lastrowid,stt,code,name,serial,key,extras)
                        if key: seen[key]=source_row
                        else: blank_codes.add(code)
                    for reason in dict.fromkeys(reasons):
                        issues.append((source_row,reason,stt,code,name,serial))
                cur=con.execute("""INSERT INTO import_batches
                    (filename,mode,created_at,added,skipped,missing,duplicate)
                    VALUES(?,?,?,?,?,?,?)""",
                    (file.filename,mode,now(),added,skipped,missing,duplicate))
                con.executemany("""INSERT INTO import_issues
                    (batch_id,source_row,reason,stt,asset_code,name,serial)
                    VALUES(?,?,?,?,?,?,?)""",
                    [(cur.lastrowid,*entry) for entry in issues])
                batch_id=cur.lastrowid
            with db() as con:
                write_audit(con,"Nhập Excel",target_type="Danh mục thiết bị",target_label=file.filename,
                            details=f"Chế độ: {mode}; Thêm {added}; Bỏ qua {skipped}; Thiếu seri {missing}; Trùng seri {duplicate}.")
            return redirect(url_for("import_result",batch_id=batch_id))
        except Exception as exc:
            with db() as con:
                write_audit(con,"Nhập Excel",success=False,target_type="Danh mục thiết bị",target_label=file.filename,details=str(exc))
            flash(f"Không thể nhập file: {exc}")
            return redirect(url_for("devices")+"#import")
    return redirect(url_for("devices")+"#import")

@app.route("/import/<int:batch_id>")
@permission_required("excel_import")
def import_result(batch_id):
    with db() as con:
        batch=con.execute("SELECT * FROM import_batches WHERE id=?",(batch_id,)).fetchone()
        issues=con.execute("SELECT * FROM import_issues WHERE batch_id=? ORDER BY source_row,id",
                           (batch_id,)).fetchall()
    if not batch: return "Không tìm thấy lần nhập",404
    return page("""<h1>Kết quả nhập Excel</h1><div class="card">
    <p>File: {{batch['filename']}} · {{'Thay danh mục' if batch['mode']=='replace' else 'Cập nhật danh mục'}}</p>
    <div class="grid"><div>Đã thêm <div class="stat">{{batch['added']}}</div></div>
    <div>Bỏ qua <div class="stat">{{batch['skipped']}}</div></div>
    <div>Thiếu seri <div class="stat">{{batch['missing']}}</div></div>
    <div>Seri trùng <div class="stat">{{batch['duplicate']}}</div></div></div>
    <p><a class="btn" href="{{url_for('import_issues_excel',batch_id=batch['id'],kind='missing')}}">Xuất danh sách thiếu seri</a>
    <a class="btn" href="{{url_for('import_issues_excel',batch_id=batch['id'],kind='duplicate')}}">Xuất danh sách trùng seri</a>
    <a class="btn secondary" href="{{url_for('import_issues_excel',batch_id=batch['id'])}}">Xuất toàn bộ lỗi</a>
    <a class="btn secondary" href="{{url_for('devices')}}">Xem danh mục</a></p></div>
    <div class="card"><h2>Chi tiết cần rà soát ({{issues|length}} mục)</h2>
    <div class="table-wrap"><table><tr><th>Dòng Excel</th><th>Lý do</th><th>STT</th>
    <th>Mã tài sản</th><th>Tên thiết bị</th><th>Seri</th></tr>
    {% for x in issues[:200] %}<tr><td>{{x['source_row']}}</td><td>{{x['reason']}}</td>
    <td>{{x['stt']}}</td><td>{{x['asset_code']}}</td><td class="table-name">{{x['name']}}</td>
    <td class="serial">{{x['serial']}}</td></tr>{% endfor %}</table></div>
    {% if issues|length>200 %}<p>Còn {{issues|length-200}} mục trong file Excel tải về.</p>{% endif %}
    </div>""",batch=batch,issues=issues)

@app.route("/import/<int:batch_id>/issues.xlsx")
@permission_required("excel_import")
def import_issues_excel(batch_id):
    kind=request.args.get("kind","all")
    if kind not in ("all","missing","duplicate"): return "Loại danh sách không hợp lệ",404
    with db() as con:
        if not con.execute("SELECT 1 FROM import_batches WHERE id=?",(batch_id,)).fetchone():
            return "Không tìm thấy lần nhập",404
        issues=con.execute("SELECT * FROM import_issues WHERE batch_id=? ORDER BY source_row,id",
                           (batch_id,)).fetchall()
    if kind=="missing": issues=[x for x in issues if "Thiếu seri" in x["reason"]]
    elif kind=="duplicate": issues=[x for x in issues if "Trùng seri" in x["reason"] or "Đã có seri" in x["reason"]]
    wb=Workbook();ws=wb.active;ws.title="Cần rà soát"
    ws.append(["Dòng Excel","Lý do","STT","Mã tài sản","Tên thiết bị","Số seri"])
    for cell in ws[1]:cell.font=Font(bold=True)
    for x in issues:
        append_excel_values(ws,[x["source_row"],x["reason"],x["stt"],
                                x["asset_code"],x["name"],x["serial"]])
    for col,width in zip("ABCDEF",[14,36,12,22,44,50]):ws.column_dimensions[col].width=width
    output=io.BytesIO();wb.save(output);output.seek(0)
    prefix={"all":"Can_ra_soat","missing":"Thieu_seri","duplicate":"Trung_seri"}[kind]
    with db() as con:
        write_audit(con,"Xuất danh sách lỗi nhập",target_type="Lần nhập Excel",target_id=batch_id,target_label=prefix,details=f"Xuất {len(issues)} dòng lỗi.")
    return send_file(output,as_attachment=True,download_name=f"{prefix}_lan_nhap_{batch_id}.xlsx")

@app.route("/sessions/<int:user_id>/revoke",methods=["POST"])
@login_required
def revoke_user_sessions(user_id):
    if session["role"]!="superadmin": return "Chỉ Superadmin được quản lý phiên đăng nhập.",403
    with db() as con:
        target=con.execute("SELECT username,role FROM users WHERE id=?",(user_id,)).fetchone()
        if not target or target["role"]=="superadmin":
            return "Không tìm thấy tài khoản phù hợp.",404
        con.execute("UPDATE auth_sessions SET status='revoked',ended_at=? WHERE user_id=? AND status='active'",(now(),user_id))
        con.execute("UPDATE users SET auth_version=auth_version+1 WHERE id=?",(user_id,))
        write_audit(con,"Đăng xuất tất cả phiên",target_type="Tài khoản",target_id=user_id,
                    target_label=target["username"],details="Superadmin đã yêu cầu đăng xuất toàn bộ phiên của tài khoản.")
    flash(f"Đã vô hiệu hóa tất cả phiên đăng nhập của {target['username']}.")
    return redirect(url_for("users"))

@app.route("/backup-db")
@login_required
def backup_db():
    if session["role"]!="superadmin":
        return "Chỉ Superadmin được sao lưu dữ liệu.",403
    stamp=datetime.now().strftime("%Y%m%d_%H%M%S")
    fd, backup_path=tempfile.mkstemp(prefix=f"medical_inventory_backup_{stamp}_",suffix=".db")
    os.close(fd)
    try:
        source=sqlite3.connect(DB_PATH)
        target=sqlite3.connect(backup_path)
        try:
            source.backup(target)
        finally:
            target.close(); source.close()
        with db() as con:
            write_audit(con,"Sao lưu dữ liệu",target_type="Cơ sở dữ liệu",target_label="medical_inventory.db",details="Tạo bản sao SQLite đầy đủ bằng cơ chế backup của SQLite.")
        @after_this_request
        def cleanup(response):
            try: os.remove(backup_path)
            except OSError: pass
            return response
        return send_file(backup_path,as_attachment=True,download_name=f"medical_inventory_backup_{stamp}.db",mimetype="application/octet-stream")
    except Exception as exc:
        try: os.remove(backup_path)
        except OSError: pass
        with db() as con:
            write_audit(con,"Sao lưu dữ liệu",success=False,target_type="Cơ sở dữ liệu",target_label="medical_inventory.db",details=f"Sao lưu thất bại: {type(exc).__name__}")
        flash("Không thể tạo bản sao dữ liệu. Hãy thử lại và kiểm tra quyền ghi của thư mục.")
        return redirect(url_for("home"))

@app.route("/audit-log")
@login_required
def audit_log():
    if session["role"]!="superadmin": return "Chỉ Superadmin được xem nhật ký.",403
    q=request.args.get("q","").strip()
    user_filter=request.args.get("user","").strip()
    kind=request.args.get("kind","all")
    success=request.args.get("success","")
    date_from=request.args.get("from","").strip()
    date_to=request.args.get("to","").strip()
    clauses=[];params=[]
    if kind=="login": clauses.append("action='Đăng nhập'")
    elif kind=="activity": clauses.append("action<>'Đăng nhập'")
    if user_filter: clauses.append("username=?");params.append(user_filter)
    if success in ("0","1"): clauses.append("success=?");params.append(int(success))
    if q:
        clauses.append("(username||action||target_label||details||ip_address||location||isp) LIKE ?")
        params.append(f"%{q}%")
    if date_from: clauses.append("created_at>=?");params.append(date_from+" 00:00:00")
    if date_to: clauses.append("created_at<=?");params.append(date_to+" 23:59:59")
    where=" AND ".join(clauses) or "1=1"
    with db() as con:
        cutoff=dt_text(datetime.now()-INACTIVITY_TIMEOUT)
        absolute_cutoff=dt_text(datetime.now()-ABSOLUTE_SESSION_TIMEOUT)
        con.execute("UPDATE auth_sessions SET status='expired',ended_at=? WHERE status='active' AND (last_activity<? OR created_at<?)",
                    (now(),cutoff,absolute_cutoff))
        users=con.execute("SELECT DISTINCT username FROM audit_logs WHERE username<>'' ORDER BY username").fetchall()
        total=con.execute(f"SELECT count(*) FROM audit_logs WHERE {where}",params).fetchone()[0]
        rows=con.execute(f"SELECT * FROM audit_logs WHERE {where} ORDER BY id DESC LIMIT 200 OFFSET ?",params+[max(0,(request.args.get('page',1,type=int)-1))*200]).fetchall()
        active_sessions=con.execute("SELECT s.*,u.username,u.role FROM auth_sessions s JOIN users u ON u.id=s.user_id WHERE s.status='active' ORDER BY s.last_activity DESC").fetchall()
    page_num,max_pages,_,_=pagination(total,request.args.get("page",1,type=int),200)
    return page(r'''<h1>Nhật ký hệ thống</h1>
    <div class="filter-links"><a href="{{url_for('audit_log',kind='all')}}" {% if kind=='all' %}aria-current="page"{% endif %}>Tất cả</a>
      <a href="{{url_for('audit_log',kind='login')}}" {% if kind=='login' %}aria-current="page"{% endif %}>Lịch sử đăng nhập</a>
      <a href="{{url_for('audit_log',kind='activity')}}" {% if kind=='activity' %}aria-current="page"{% endif %}>Nhật ký hoạt động</a>
    </div>
    <div class="card"><form method="get" class="grid">
      <input type="hidden" name="kind" value="{{kind}}"><div><label>Tài khoản</label><select name="user"><option value="">Tất cả</option>{% for u in users %}<option value="{{u['username']}}" {% if user_filter==u['username'] %}selected{% endif %}>{{u['username']}}</option>{% endfor %}</select></div>
      <div><label>Kết quả</label><select name="success"><option value="">Tất cả</option><option value="1" {% if success=='1' %}selected{% endif %}>Thành công</option><option value="0" {% if success=='0' %}selected{% endif %}>Thất bại</option></select></div>
      <div><label>Từ ngày</label><input type="date" name="from" value="{{date_from}}"></div><div><label>Đến ngày</label><input type="date" name="to" value="{{date_to}}"></div>
      <div><label>Tìm kiếm</label><input name="q" value="{{q}}" placeholder="Tài khoản, thao tác, IP, nội dung..." style="width:100%;box-sizing:border-box"></div><div class="audit-filter-actions"><button type="submit">Tìm</button><a class="btn secondary" href="{{url_for('audit_export',q=q,user=user_filter,kind=kind,success=success,from=date_from,to=date_to)}}">Xuất Excel</a></div>
    </form></div>
    {% if kind in ('all','login') %}<div class="card"><h2>Phiên đăng nhập đang hoạt động ({{active_sessions|length}})</h2>
      {% if active_sessions %}<div class="table-wrap"><table class="audit-sessions"><tr><th class="col-session-user">Tài khoản</th><th class="col-session-role">Vai trò</th><th class="col-session-time">Thời điểm đăng nhập</th><th class="col-session-last">Hoạt động gần nhất</th><th class="col-session-ip">IP</th><th class="col-session-location">Vị trí ước tính</th><th class="col-session-device">Thiết bị</th><th class="col-session-action">Thao tác</th></tr>
      {% for s in active_sessions %}<tr><td>{{s['username']}}</td><td>{{s['role']}}</td><td>{{s['created_at']}}</td><td>{{s['last_activity']}}</td><td>{{s['ip_address']}}</td><td>{{s['location']}}</td><td>{{s['user_agent']}}</td><td>{% if s['role']!='superadmin' %}<form method="post" action="{{url_for('revoke_user_sessions',user_id=s['user_id'])}}" style="margin:0" onsubmit='return confirm("Đăng xuất tất cả phiên của {{s['username']}}?");'><button class="danger">Đăng xuất tất cả</button></form>{% else %}—{% endif %}</td></tr>{% endfor %}</table></div>{% else %}<p class="muted">Không có phiên đang hoạt động.</p>{% endif %}</div>{% endif %}
    <div class="card"><p>Hiển thị {{rows|length}} / {{total}} bản ghi · Trang {{page_num}}/{{max_pages}}</p>
    <table class="audit-table"><tr><th class="col-time">Thời gian</th><th class="col-user">Tài khoản</th><th class="col-role">Vai trò</th><th class="col-action">Hoạt động</th><th class="col-target">Đối tượng</th><th class="col-result">Kết quả</th><th class="col-ip">IP</th><th class="col-location">Vị trí ước tính</th><th class="col-details">Chi tiết</th></tr>
    {% for x in rows %}<tr><td>{{x['created_at']}}</td><td>{{x['username']}}</td><td>{{x['role']}}</td><td>{{x['action']}}</td><td>{{x['target_label']}}</td><td>{{'Thành công' if x['success'] else 'Thất bại'}}</td><td>{{x['ip_address']}}</td><td>{{x['location']}}</td><td>{{x['details']}}</td></tr>{% endfor %}</table>
    <div class="pager">{% if page_num>1 %}<a class="btn secondary" href="{{url_for('audit_log',q=q,user=user_filter,kind=kind,success=success,from=date_from,to=date_to,page=page_num-1)}}">← Trước</a>{% endif %}<span>Trang {{page_num}}/{{max_pages}}</span>{% if page_num<max_pages %}<a class="btn secondary" href="{{url_for('audit_log',q=q,user=user_filter,kind=kind,success=success,from=date_from,to=date_to,page=page_num+1)}}">Sau →</a>{% endif %}</div></div>''',
           users=users,active_sessions=active_sessions,rows=rows,total=total,page_num=page_num,max_pages=max_pages,
           q=q,user_filter=user_filter,kind=kind,success=success,date_from=date_from,date_to=date_to)

@app.route("/audit-log.xlsx")
@login_required
def audit_export():
    if session["role"]!="superadmin": return "Chỉ Superadmin được xuất nhật ký.",403
    q=request.args.get("q","").strip();user_filter=request.args.get("user","").strip();kind=request.args.get("kind","all");success=request.args.get("success","")
    date_from=request.args.get("from","").strip();date_to=request.args.get("to","").strip()
    clauses=[];params=[]
    if kind=="login": clauses.append("action='Đăng nhập'")
    elif kind=="activity": clauses.append("action<>'Đăng nhập'")
    if user_filter: clauses.append("username=?");params.append(user_filter)
    if success in ("0","1"): clauses.append("success=?");params.append(int(success))
    if q: clauses.append("(username||action||target_label||details||ip_address||location||isp) LIKE ?");params.append(f"%{q}%")
    if date_from: clauses.append("created_at>=?");params.append(date_from+" 00:00:00")
    if date_to: clauses.append("created_at<=?");params.append(date_to+" 23:59:59")
    where=" AND ".join(clauses) or "1=1"
    with db() as con:
        rows=con.execute(f"SELECT * FROM audit_logs WHERE {where} ORDER BY id DESC",params).fetchall()
        write_audit(con,"Xuất nhật ký",target_type="Nhật ký",target_label="audit_logs",details=f"Xuất {len(rows)} bản ghi nhật ký.")
    wb=Workbook();ws=wb.active;ws.title="Nhật ký"
    ws.append(["Thời gian","Tài khoản","Vai trò","Hoạt động","Đối tượng","Kết quả","IP","Vị trí ước tính","ISP","Thiết bị","Chi tiết"])
    for c in ws[1]: c.font=Font(bold=True)
    for x in rows:
        append_excel_values(ws,[x["created_at"],x["username"],x["role"],x["action"],x["target_label"],"Thành công" if x["success"] else "Thất bại",x["ip_address"],x["location"],x["isp"],x["user_agent"],x["details"]])
    for col,width in zip("ABCDEFGHIJK",[22,18,14,26,24,14,18,34,30,32,70]): ws.column_dimensions[col].width=width
    ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
    out=io.BytesIO();wb.save(out);out.seek(0)
    return send_file(out,as_attachment=True,download_name="Nhat_ky_he_thong.xlsx")

@app.route("/rounds",methods=["GET","POST"])
@login_required
def rounds():
    if request.method=="POST":
        if not has_permission("round_manage"): return "Không có quyền tạo đợt kiểm kê.",403
        name=request.form.get("name","").strip()
        if not name: flash("Hãy nhập tên đợt kiểm kê.")
        else:
            with db() as con:
                con.execute("UPDATE inventory_rounds SET active=0")
                cur=con.execute("INSERT INTO inventory_rounds(name,created_at) VALUES(?,?)",
                                (name,now()))
                con.execute("""INSERT INTO round_items
                    (round_id,device_id,stt,asset_code,name,serial,serial_key,
                     model,maker,country,year_made,year_used,year_increase,department)
                    SELECT ?,id,stt,asset_code,name,serial,serial_key,
                           model,maker,country,year_made,year_used,year_increase,department
                    FROM devices WHERE status='Đang sử dụng' ORDER BY id""",(cur.lastrowid,))
                write_audit(con,"Tạo đợt kiểm kê",target_type="Đợt kiểm kê",target_id=cur.lastrowid,
                            target_label=name,details="Tạo đợt mới và chụp danh sách thiết bị hiện hành.")
            flash("Đã tạo đợt kiểm kê và lưu danh sách thiết bị của đợt.")
        return redirect(url_for("rounds"))
    if not any(has_permission(p) for p in ("round_manage","results_view","results_export")):
        return "Không có quyền xem đợt kiểm kê.",403
    with db() as con:
        rows=con.execute("""SELECT r.*,
            (SELECT count(*) FROM round_items i WHERE i.round_id=r.id) total,
            (SELECT count(DISTINCT s.device_id) FROM scans s WHERE s.round_id=r.id) counted
            FROM inventory_rounds r ORDER BY r.id DESC""").fetchall()
    return page("""<h1>Đợt kiểm kê</h1>{% if can('round_manage') %}<div class="card"><form method="post">
    <label>Tên đợt mới</label><input name="name" required>
    <button>Tạo đợt mới</button></form>
    <p class="muted">Mỗi seri chỉ được ghi nhận một lần trong một đợt. Đợt mới trở thành đợt đang thực hiện.</p>
    </div>{% endif %}<div class="card table-wrap"><table><tr><th>Đợt</th><th>Ngày tạo</th>
    <th>Đã quét / tổng</th><th>Trạng thái</th><th>Thao tác</th></tr>
    {% for r in rows %}<tr><td>{{r['name']}}</td><td>{{r['created_at']}}</td>
    <td>{{r['counted']}} / {{r['total']}}</td>
    <td>{{'Đang thực hiện' if r['active'] else 'Đã lưu'}}</td><td><div class="round-actions">
    {% if can('results_view') %}<a class="btn" href="{{url_for('results',round_id=r['id'])}}">Xem kết quả</a>{% endif %}
    {% if can('results_export') %}<a class="btn" href="{{url_for('report',round_id=r['id'])}}">Xuất Excel</a>{% endif %}
    {% if r['active'] and can('scan') %}<a class="btn" href="{{url_for('scan',round_id=r['id'])}}">Quét QR</a>{% endif %}
    {% if can('round_manage') %}<form method="post" action="{{url_for('delete_round',round_id=r['id'])}}"
      onsubmit='return confirm("Xóa đợt " + {{r["name"]|tojson}} + " cùng {{r["counted"]}} kết quả quét? Báo cáo đợt này sẽ không còn.");'
      ><button class="danger">Xóa đợt</button></form>{% endif %}</div>
    </td></tr>{% endfor %}</table></div>""",rows=rows)

@app.route("/rounds/<int:round_id>/delete",methods=["POST"])
@permission_required("round_manage")
def delete_round(round_id):
    with db() as con:
        r=con.execute("SELECT name FROM inventory_rounds WHERE id=?",(round_id,)).fetchone()
        if not r: return "Không tìm thấy đợt kiểm kê",404
        con.execute("DELETE FROM scans WHERE round_id=?",(round_id,))
        con.execute("DELETE FROM scan_once WHERE round_id=?",(round_id,))
        con.execute("DELETE FROM round_items WHERE round_id=?",(round_id,))
        con.execute("DELETE FROM inventory_rounds WHERE id=?",(round_id,))
        write_audit(con,"Xóa đợt kiểm kê",target_type="Đợt kiểm kê",target_id=round_id,
                    target_label=r["name"],details="Xóa đợt và toàn bộ kết quả quét của đợt.")
    flash(f"Đã xóa đợt kiểm kê {r['name']} cùng các kết quả quét của đợt.")
    return redirect(url_for("rounds"))

def record_scan(con, round_id, serial, username="", client_event_id="", client_scanned_at=""):
    key=serial_key(serial)
    if client_event_id:
        existing=con.execute("SELECT id,device_id,round_id FROM scans WHERE client_event_id=?",(client_event_id,)).fetchone()
        if existing:
            return {"success":True,"message":"Đã đồng bộ thành công trước đó.","already_synced":True}
    matches=con.execute("""SELECT * FROM round_items WHERE round_id=?
        AND serial_key=? AND serial_key<>''""",(round_id,key)).fetchall()
    if not matches:
        return {"success":False,"message":"Không tìm thấy số seri này trong danh mục của đợt."}
    if len(matches)>1:
        return {"success":False,"message":"Seri trùng ở nhiều thiết bị. Hãy sửa dữ liệu trước khi điểm danh."}
    d=matches[0]
    inserted=con.execute("INSERT OR IGNORE INTO scan_once(round_id,device_id) VALUES(?,?)",
                         (round_id,d["device_id"]))
    if inserted.rowcount==0:
        return {"success":False,"message":"Thiết bị này đã được điểm danh trước đó."}
    try:
        con.execute("""INSERT INTO scans(round_id,device_id,scanned_at,scanned_by,client_event_id,client_scanned_at)
                       VALUES(?,?,?,?,?,?)""",
                    (round_id,d["device_id"],now(),username,client_event_id or "",client_scanned_at or ""))
    except sqlite3.IntegrityError:
        # Nếu request bị gửi lại sau khi giao dịch đã thành công, trả về thành công idempotent.
        existing=con.execute("SELECT 1 FROM scans WHERE client_event_id=?",(client_event_id,)).fetchone() if client_event_id else None
        if existing:
            return {"success":True,"message":"Đã đồng bộ thành công trước đó.","already_synced":True}
        raise
    return {"success":True,"message":f"ĐÃ GHI NHẬN CÓ: {d['asset_code']} — {d['name']} — Seri: {d['serial']}",
            "device_id":d["device_id"],"asset_code":d["asset_code"],"name":d["name"]}

SCAN_VIEW = """<h1>Quét QR</h1><div class="card"><h2>Đợt đang điểm danh: {{r['name']}}</h2>
{% if not r['active'] %}<div class="warn">Đợt này đã lưu, không thể quét thêm.</div>{% else %}
<p>Dùng camera để quét mã QR. Mỗi thiết bị được ghi nhận một lần trong đợt.</p>
<p class="muted">Quét thành công sẽ tự phát tiếng tít. Khi mất kết nối máy chủ, mã sẽ được lưu tạm an toàn trên trình duyệt và tự đồng bộ khi kết nối trở lại.</p>
<div class="scan-connect card" style="margin:10px 0;padding:12px">
  <b id="network-status">Đang kiểm tra kết nối…</b> · <span id="queue-status">Hàng đợi offline: 0</span>
  <button type="button" id="sync-now" class="secondary" style="margin-left:8px">Đồng bộ ngay</button>
  <div id="queue-detail" class="muted" style="margin-top:5px"></div>
</div>
<div><label for="camera-choice">Chọn camera</label>
<select id="camera-choice" aria-label="Chọn camera"><option value="">Chưa chọn camera</option></select>
<button type="button" id="start-camera">Bắt đầu quét</button>
<button type="button" id="stop-camera" class="secondary" disabled>Dừng quét</button>
<label for="qr-file" class="btn secondary" style="display:inline-block;cursor:pointer">Chọn ảnh chứa mã QR</label>
<input type="file" id="qr-file" accept="image/*" style="position:absolute;width:1px;height:1px;opacity:0" aria-label="Chọn ảnh chứa mã QR"></div>
<p id="camera-status" class="muted" role="status" aria-live="polite">Đang nhận diện camera…</p>
<div id="reader" style="max-width:480px"></div>
<div id="scan-result" role="status" aria-live="polite" {% if result %}class="{{'ok' if result['success'] else 'warn'}}"{% endif %}>
{{result['message'] if result else ''}}</div>
<details class="manual-entry" {% if entered %}open{% endif %}><summary>Không quét được mã QR? Nhập số seri thủ công</summary>
<p class="muted">Chỉ sử dụng khi camera hoặc mã QR không đọc được. Mỗi seri chỉ được ghi nhận một lần trong đợt.</p>
<form method="post" id="scan-form"><label for="serial">Số seri</label>
<input name="serial" id="serial" value="{{entered}}" autocomplete="off" required style="width:min(90%,440px)">
<button>Ghi nhận bằng seri</button></form></details>
<script src="https://unpkg.com/html5-qrcode" defer></script>
<script>
(function(){
 const form=document.getElementById('scan-form'),field=document.getElementById('serial');
 const result=document.getElementById('scan-result');
 const cameraChoice=document.getElementById('camera-choice'),cameraStatus=document.getElementById('camera-status');
 const startButton=document.getElementById('start-camera'),stopButton=document.getElementById('stop-camera');
 const fileInput=document.getElementById('qr-file');
 let reader=null,scanning=false,transitioning=false;
 let audioContext=null,busy=false,lastValue='',lastAt=0;
 async function prepareSound(){
   try{
     audioContext=audioContext||new (window.AudioContext||window.webkitAudioContext)();
     await audioContext.resume();
   }catch(e){/* Trình duyệt có thể tắt âm; điểm danh vẫn hoạt động. */}
 }
 document.addEventListener('pointerdown',prepareSound,{capture:true,once:true});
 document.addEventListener('keydown',prepareSound,{capture:true,once:true});
 function beep(){
   if(!audioContext||audioContext.state!=='running')return;
   const tone=audioContext.createOscillator(),gain=audioContext.createGain(),t=audioContext.currentTime;
   tone.type='square';tone.frequency.setValueAtTime(940,t);
   gain.gain.setValueAtTime(0.001,t);
   gain.gain.exponentialRampToValueAtTime(0.30,t+0.02);
   gain.gain.setValueAtTime(0.30,t+0.19);
   gain.gain.exponentialRampToValueAtTime(0.001,t+0.25);
   tone.connect(gain);gain.connect(audioContext.destination);tone.start(t);tone.stop(t+0.26);
 }
 async function openQueue(){
   if(!window.indexedDB)return null;
   return new Promise(function(resolve,reject){
     const req=indexedDB.open('MedicalInventoryOffline',1);
     req.onupgradeneeded=function(){
       const db=req.result;
       if(!db.objectStoreNames.contains('scanQueue')){
         const store=db.createObjectStore('scanQueue',{keyPath:'id'});
         store.createIndex('round_serial',['roundId','serial'],{unique:true});
         store.createIndex('status','status',{unique:false});
       }
     };
     req.onsuccess=function(){resolve(req.result);};
     req.onerror=function(){reject(req.error);};
   });
 }
 async function queueAll(){
   const db=await openQueue();if(!db)return [];
   return new Promise(function(resolve,reject){
     const tx=db.transaction('scanQueue','readonly'),req=tx.objectStore('scanQueue').getAll();
     req.onsuccess=function(){resolve(req.result||[])};req.onerror=function(){reject(req.error)};
   });
 }
 async function putQueue(item){
   const db=await openQueue();if(!db)throw new Error('Trình duyệt không hỗ trợ lưu offline.');
   return new Promise(function(resolve,reject){
     const tx=db.transaction('scanQueue','readwrite');
     tx.objectStore('scanQueue').put(item);
     tx.oncomplete=function(){resolve()};tx.onerror=function(){reject(tx.error)};
   });
 }
 async function deleteQueue(id){
   const db=await openQueue();if(!db)return;
   return new Promise(function(resolve,reject){
     const tx=db.transaction('scanQueue','readwrite');tx.objectStore('scanQueue').delete(id);
     tx.oncomplete=function(){resolve()};tx.onerror=function(){reject(tx.error)};
   });
 }
 function clientSerialKey(value){
   return String(value||'').trim().split(/\s+/).join(' ');
 }
 const offlineSerials=new Set({{offline_serials|tojson}});
 const scannedSerials=new Set({{scanned_serials|tojson}});
 async function cleanupQueue(){
   const items=await queueAll();
   const seen=new Map();
   for(const item of items){
     const serialKey=clientSerialKey(item.serial);
     const key=String(item.roundId)+'\u0000'+serialKey;
     // Nếu thiết bị đã được ghi nhận thành công trên máy chủ, bản ghi offline cũ
     // không còn cần thiết và phải được loại khỏi hàng đợi.
     if(item.roundId==={{r['id']}} && scannedSerials.has(serialKey)){
       await deleteQueue(item.id);
       continue;
     }
     // Với chính đợt đang mở, mọi seri không thuộc snapshot đều bị loại khỏi hàng đợi.
     if(item.roundId==={{r['id']}} && !offlineSerials.has(serialKey)){
       await deleteQueue(item.id);
       continue;
     }
     const previous=seen.get(key);
     if(!previous){
       seen.set(key,item);
       continue;
     }
     // Giữ bản ghi tốt hơn: pending trước failed; nếu cùng trạng thái thì giữ bản ghi cũ hơn.
     const itemRank=item.status==='failed'?0:1;
     const prevRank=previous.status==='failed'?0:1;
     const keepItem=(itemRank>prevRank || (itemRank===prevRank && (item.createdAt||0)<(previous.createdAt||0)))?item:previous;
     const removeItem=keepItem.id===item.id?previous:item;
     await deleteQueue(removeItem.id);
     seen.set(key,keepItem);
   }
 }
 async function refreshQueue(){
   try{
     await cleanupQueue();
     const items=await queueAll();
     const pending=items.filter(function(x){return x.status!=='failed'}).length;
     const failed=items.filter(function(x){return x.status==='failed'}).length;
     document.getElementById('queue-status').textContent='Hàng đợi offline: '+pending+(failed?' · Lỗi cần xử lý: '+failed:'');
     document.getElementById('queue-detail').textContent=failed?'Một số mã đã đồng bộ thất bại do dữ liệu trên máy chủ thay đổi hoặc đã được ghi nhận ở nơi khác.':' ';
   }catch(e){document.getElementById('queue-status').textContent='Không đọc được hàng đợi offline';}
 }
 async function queueOffline(value,reason){
   try{
     const key=clientSerialKey(value);
     if(!offlineSerials.has(key)){
       result.className='warn';
       result.textContent='Mất kết nối máy chủ. Mã QR này không thuộc danh mục của đợt nên không được lưu vào hàng đợi.';
       return;
     }
     if(scannedSerials.has(key)){
       result.className='warn';
       result.textContent='Mất kết nối máy chủ. Thiết bị này đã được ghi nhận trước đó trong đợt, nên không đưa lại vào hàng đợi offline.';
       await refreshQueue();
       return;
     }
     const all=await queueAll();
     const existing=all.find(function(x){return x.roundId==={{r['id']}} && clientSerialKey(x.serial)===key;});
     if(existing){
       result.className='warn';
       result.textContent=existing.status==='failed'
         ? 'Mất kết nối máy chủ. Mã này đã có trong hàng đợi (đang ở trạng thái lỗi), không tạo thêm bản ghi mới.'
         : 'Mất kết nối máy chủ. Mã này đã được lưu trong hàng đợi offline trước đó.';
       await refreshQueue();
       return;
     }
     const item={id:crypto.randomUUID?crypto.randomUUID():('evt-'+Date.now()+'-'+Math.random().toString(16).slice(2)),
       roundId:{{r['id']}},serial:value,clientScannedAt:new Date().toISOString(),status:'pending',error:reason||'',createdAt:Date.now()};
     await putQueue(item);
     result.className='warn';result.textContent='Mất kết nối máy chủ. QR hợp lệ đã được lưu tạm và sẽ tự đồng bộ khi có mạng.';
     await refreshQueue();
   }catch(e){result.className='warn';result.textContent='Mất kết nối và không thể lưu tạm. Vui lòng quét lại khi kết nối ổn định.';}
 }
 function setNetworkStatus(online,detail){
   const el=document.getElementById('network-status');
   if(!el)return;
   el.textContent=online?'Đã kết nối máy chủ':'Mất kết nối máy chủ';
   el.style.color=online?'#26823e':'#ae3040';
   el.title=detail||'';
 }
 async function checkServerConnection(){
   try{
     const response=await fetch(location.href,{method:'GET',cache:'no-store',headers:{'Accept':'text/html'}});
     if(!response.ok)throw new Error('HTTP '+response.status);
     if(response.redirected && response.url.indexOf('/login')!==-1){
       setNetworkStatus(false,'Phiên đăng nhập đã hết hạn.');
       return false;
     }
     setNetworkStatus(true,'Máy chủ phản hồi bình thường.');
     const queued=await queueAll();
     if(queued.some(function(x){return x.status!=='failed';})) syncQueue();
     return true;
   }catch(e){
     setNetworkStatus(false,'Không thể kết nối tới máy chủ.');
     return false;
   }
 }

 async function syncQueue(){
   if(busy)return;
   const items=(await queueAll()).filter(function(x){return x.status!=='failed';});
   if(!items.length){await refreshQueue();return;}
   busy=true;
   document.getElementById('sync-now').disabled=true;
   document.getElementById('queue-detail').textContent='Đang đồng bộ '+items.length+' mã…';
   try{
     for(const item of items){
       const body=new FormData();
       body.append('serial',item.serial);body.append('client_event_id',item.id);
       body.append('client_scanned_at',item.clientScannedAt||'');body.append('sync_mode','offline');
       const csrf=form.querySelector('input[name="_csrf"]').value;body.append('_csrf',csrf);
       try{
         const endpoint='/scan/{{r['id']}}';
     const response=await fetch(endpoint,{method:'POST',body:body,headers:{'Accept':'application/json'}});
         if(response.status===401){setNetworkStatus(true,'Máy chủ hoạt động nhưng phiên đăng nhập đã hết hạn.');document.getElementById('queue-detail').textContent='Phiên đăng nhập đã hết hạn. Đăng nhập lại để tiếp tục đồng bộ.';break;}
         const data=await response.json();
         if(response.ok && data.success){
           scannedSerials.add(clientSerialKey(item.serial));
           await deleteQueue(item.id);
         }
         else if(response.status>=500){break;}
         else{item.status='failed';item.error=data.message||'Máy chủ từ chối dữ liệu.';await putQueue(item);}
       }catch(e){setNetworkStatus(false,'Mất kết nối trong khi đồng bộ hàng đợi.');break;}
     }
   }finally{busy=false;document.getElementById('sync-now').disabled=false;await refreshQueue();}
 }
 document.getElementById('sync-now').addEventListener('click',async function(){
   if(busy)return;
   const button=this;
   button.disabled=true;
   document.getElementById('queue-detail').textContent='Đang chuẩn bị đồng bộ…';
   try{await syncQueue();}
   finally{button.disabled=false;}
 });
 window.addEventListener('online',function(){
   setNetworkStatus(true,'Trình duyệt vừa báo có kết nối trở lại.');
   syncQueue();
 });

 async function submit(value){
   if(busy)return;
   value=value.trim();if(!value)return;
   const moment=Date.now();if(value===lastValue&&moment-lastAt<1800)return;
   lastValue=value;lastAt=moment;busy=true;field.value=value;
   const clientEventId=crypto.randomUUID?crypto.randomUUID():('evt-'+Date.now()+'-'+Math.random().toString(16).slice(2));
   try{
     const body=new FormData(form);body.set('serial',value);body.set('client_event_id',clientEventId);
     body.set('client_scanned_at',new Date().toISOString());body.set('sync_mode','online');
     const endpoint=form.action||location.href;
     let response;
     try{
       response=await fetch(endpoint,{method:'POST',body:body,headers:{'Accept':'application/json'}});
     }catch(e){
       setNetworkStatus(false,'Không thể kết nối tới máy chủ. QR được chuyển sang hàng đợi offline.');
       await queueOffline(value,e&&e.message?e.message:'Không thể kết nối máy chủ.');
       return;
     }
     if(response.status===401){
       result.className='warn';result.textContent='Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.';
       location.href='{{url_for('login')}}';return;
     }
     const contentType=(response.headers.get('content-type')||'').toLowerCase();
     if(!contentType.includes('application/json')){
       const text=await response.text();
       result.className='warn';
       result.textContent='Máy chủ trả về HTTP '+response.status+'. '+(text.slice(0,180)||'Phản hồi không hợp lệ.');
       return;
     }
     const data=await response.json();
     if(!response.ok){
       result.className='warn';
       result.textContent=data.message||('Máy chủ trả về HTTP '+response.status+'.');
       return;
     }
     result.className=data.success?'ok':'warn';result.textContent=data.message;
     if(data.success){
       scannedSerials.add(clientSerialKey(value));
       setNetworkStatus(true,'Vừa ghi nhận thành công từ máy chủ.');beep();field.value='';
     }else{
       // Máy chủ xác nhận thiết bị đã được điểm danh: ghi nhớ phía trình duyệt để
       // nếu mất kết nối ngay sau đó, cùng QR sẽ không bị đưa vào hàng đợi.
       if(String(data.message||'').toLowerCase().includes('đã được điểm danh')){
         scannedSerials.add(clientSerialKey(value));
       }
       field.select();
     }
   }catch(e){
     result.className='warn';
     result.textContent='Không thể xử lý phản hồi từ máy chủ: '+(e&&e.message?e.message:String(e));
   }finally{busy=false;}
 }
 function cameraMessage(message){cameraStatus.textContent=message;}
 async function stopCamera(){
   if(!reader||!scanning)return;
   transitioning=true;stopButton.disabled=true;
   try{await reader.stop();scanning=false;startButton.disabled=false;cameraMessage('Đã dừng quét.');}
   catch(e){cameraMessage('Không thể dừng camera. Hãy tải lại trang nếu cần.');}
   finally{transitioning=false;stopButton.disabled=!scanning;}
 }
 async function loadCameras(){
   startButton.disabled=true;
   if(!navigator.mediaDevices||!navigator.mediaDevices.getUserMedia){
     cameraMessage('Trình duyệt không hỗ trợ camera tại địa chỉ này. Bạn có thể nhập seri thủ công.');
     startButton.disabled=false;return [];
   }
   try{
     // Xin quyền để trình duyệt cung cấp tên camera; đóng luồng thử ngay sau đó.
     const stream=await navigator.mediaDevices.getUserMedia({video:true});
     stream.getTracks().forEach(function(track){track.stop();});
     const cameras=await Html5Qrcode.getCameras();
     const previous=cameraChoice.value;
     cameraChoice.replaceChildren();
     if(!cameras.length){cameraMessage('Không tìm thấy camera. Bạn có thể nhập seri thủ công.');return []}
     cameras.forEach(function(camera,index){
       const option=document.createElement('option');option.value=camera.id;
       option.textContent=camera.label||('Camera '+(index+1));
       cameraChoice.appendChild(option);
     });
     const preferred=cameras.find(function(camera){return camera.id===previous;})
       ||cameras.find(function(camera){return /back|rear|environment|sau/i.test(camera.label);})||cameras[0];
     cameraChoice.value=preferred.id;
     cameraMessage('Đã nhận diện camera. Chọn camera rồi bấm “Bắt đầu quét”.');
     return cameras;
   }catch(e){
     cameraMessage('Chưa được cấp quyền camera. Hãy cho phép truy cập rồi tải lại trang, hoặc nhập seri thủ công.');
     return [];
   }finally{
     if(!transitioning)startButton.disabled=false;
   }
 }
 async function startCamera(){
   if(!reader||scanning||transitioning)return;
   transitioning=true;startButton.disabled=true;cameraMessage('Đang khởi động camera…');
   try{
     const cameras=cameraChoice.options.length&&cameraChoice.value
       ? Array.from(cameraChoice.options).map(function(option){return {id:option.value,label:option.textContent};})
       : await loadCameras();
     if(!cameras.length)return;
     await reader.start(cameraChoice.value,{fps:10,qrbox:220},function(value){submit(value);});
     scanning=true;stopButton.disabled=false;cameraMessage('Đang quét mã QR bằng camera.');
   }catch(e){cameraMessage('Không thể mở camera. Hãy kiểm tra quyền truy cập camera hoặc nhập seri thủ công.');}
   finally{transitioning=false;startButton.disabled=scanning||!reader;}
 }
 startButton.addEventListener('click',startCamera);
 stopButton.addEventListener('click',stopCamera);
 cameraChoice.addEventListener('change',async function(){
   if(scanning&&!transitioning){await stopCamera();await startCamera();}
 });
 fileInput.addEventListener('change',async function(){
   const file=this.files&&this.files[0];if(!file||!reader)return;
   if(scanning)await stopCamera();
   try{
     cameraMessage('Đang đọc mã QR trong ảnh…');
     const value=await reader.scanFile(file,true);
     cameraMessage('Đã đọc mã QR từ ảnh.');await submit(value);
   }catch(e){cameraMessage('Không thể đọc mã QR từ ảnh này. Hãy chọn ảnh rõ hơn hoặc nhập seri.');}
   finally{fileInput.value='';}
 });
 window.addEventListener('load',function(){
   checkServerConnection();
   setInterval(checkServerConnection,10000);
   refreshQueue();
   prepareSound();
   if(window.Html5Qrcode){
     reader=new Html5Qrcode('reader');
     loadCameras();
   }else{
     startButton.disabled=true;cameraMessage('Chưa tải được tính năng đọc QR; vẫn có thể nhập seri thủ công.');
   }
 });
})();
</script>{% endif %}</div>"""

@app.route("/scan")
@permission_required("scan")
def active_scan():
    with db() as con:r=current_round(con)
    if not r:
        flash("Hãy tạo đợt kiểm kê trước khi quét.")
        return redirect(url_for("rounds"))
    return redirect(url_for("scan",round_id=r["id"]))

@app.route("/scan/<int:round_id>",methods=["GET","POST"])
@permission_required("scan")
def scan(round_id):
    with db() as con:
        r=con.execute("SELECT * FROM inventory_rounds WHERE id=?",(round_id,)).fetchone()
        if not r:return "Không tìm thấy đợt kiểm kê",404
        result=None
        if request.method=="POST":
            if not r["active"]:
                result={"success":False,"message":"Đợt này đã kết thúc, không thể điểm danh thêm."}
                write_audit(con,"Quét QR",success=False,target_type="Đợt kiểm kê",target_id=round_id,target_label=r["name"],
                            details=result["message"])
            else:
                serial=request.form.get("serial","")
                client_event_id=request.form.get("client_event_id","").strip()
                client_scanned_at=request.form.get("client_scanned_at","").strip()
                result=record_scan(con,round_id,serial,session["username"],client_event_id,client_scanned_at)
                write_audit(con,"Quét QR",success=result.get("success",False),target_type="Thiết bị",
                            target_id=result.get("device_id"),target_label=result.get("asset_code") or serial_key(serial),
                            details=result.get("message","")+ (" | Nguồn: offline" if request.form.get("sync_mode")=="offline" else " | Nguồn: online"))
            if request.accept_mimetypes.best=="application/json":
                return jsonify(result)
    offline_serials=[row["serial_key"] for row in con.execute(
        "SELECT serial_key FROM round_items WHERE round_id=? AND serial_key<>''", (round_id,)
    ).fetchall()]
    scanned_serials=[row["serial_key"] for row in con.execute(
        """SELECT DISTINCT i.serial_key FROM round_items i
             INNER JOIN scans s ON s.round_id=i.round_id AND s.device_id=i.device_id
             WHERE i.round_id=? AND i.serial_key<>''""", (round_id,)
    ).fetchall()]
    return page(SCAN_VIEW,r=r,result=result,entered=request.form.get("serial","") if result and not result["success"] else "",
                offline_serials=offline_serials,scanned_serials=scanned_serials)

def round_rows(con, round_id, q="", status_filter=""):
    clause={"":"", "yes":" AND EXISTS (SELECT 1 FROM scans f WHERE f.round_id=i.round_id AND f.device_id=i.device_id)",
            "no":" AND NOT EXISTS (SELECT 1 FROM scans f WHERE f.round_id=i.round_id AND f.device_id=i.device_id)"}[status_filter]
    return con.execute(f"""SELECT i.*, (SELECT min(s.scanned_at) FROM scans s
             WHERE s.round_id=i.round_id AND s.device_id=i.device_id) scanned_at,
             (SELECT s.scanned_by FROM scans s WHERE s.round_id=i.round_id
              AND s.device_id=i.device_id ORDER BY s.id LIMIT 1) scanned_by,
             (SELECT count(*) FROM scans s WHERE s.round_id=i.round_id
             AND s.device_id=i.device_id) scan_count
         FROM round_items i WHERE i.round_id=?
         AND (i.stt||i.asset_code||i.name||i.serial) LIKE ?
         {clause} ORDER BY i.id""",(round_id,f"%{q}%")).fetchall()

@app.route("/results/<int:round_id>")
@permission_required("results_view")
def results(round_id):
    q=request.args.get("q","").strip()
    status_filter=request.args.get("filter","")
    if status_filter not in ("", "yes", "no"): status_filter=""
    with db() as con:
        r=con.execute("SELECT * FROM inventory_rounds WHERE id=?",(round_id,)).fetchone()
        if not r:return "Không tìm thấy đợt kiểm kê",404
        total=con.execute("SELECT count(*) FROM round_items WHERE round_id=?",(round_id,)).fetchone()[0]
        counted=con.execute("""SELECT count(DISTINCT s.device_id) FROM scans s
            JOIN round_items i ON i.round_id=s.round_id AND i.device_id=s.device_id
            WHERE s.round_id=?""",(round_id,)).fetchone()[0]
        all_rows=round_rows(con,round_id,q,status_filter)
    number,pages,size,offset=pagination(len(all_rows),request.args.get("page",1,type=int),200)
    rows=all_rows[offset:offset+size]
    return page("""<h1>Kết quả: {{r['name']}}</h1><div class="grid">
      <div class="card">Tổng thiết bị<div class="stat">{{total}}</div></div>
      <div class="card">Đã điểm danh<div class="stat">{{counted}}</div></div>
      <div class="card">Chưa điểm danh<div class="stat">{{total-counted}}</div></div></div>
      <div class="card">{% if can('results_export') %}<a class="btn" href="{{url_for('report',round_id=r['id'])}}">Xuất Excel toàn bộ</a>{% endif %}
      <nav class="filter-links" aria-label="Lọc kết quả">
      <a href="{{url_for('results',round_id=r['id'],q=q)}}" {% if not status_filter %}aria-current="page"{% endif %}>Tất cả ({{total}})</a>
      <a href="{{url_for('results',round_id=r['id'],q=q,filter='yes')}}" {% if status_filter=='yes' %}aria-current="page"{% endif %}>Có ({{counted}})</a>
      <a href="{{url_for('results',round_id=r['id'],q=q,filter='no')}}" {% if status_filter=='no' %}aria-current="page"{% endif %}>Chưa điểm danh ({{total-counted}})</a></nav>
      <form method="get"><input type="hidden" name="filter" value="{{status_filter}}"><input name="q" value="{{q}}" placeholder="Tìm mã, tên, seri"><button>Tìm</button></form>
      <p>Hiển thị {{all_count}} kết quả tìm kiếm · Trang {{number}}/{{pages}}</p>
      <div class="table-wrap"><table><tr><th>STT</th><th>Mã tài sản</th><th>Tên VT - TB</th>
      <th>Số seri</th><th>Kết quả</th><th>Thời gian quét</th><th>Tài khoản quét</th></tr>
      {% for d in rows %}<tr><td>{{d['stt']}}</td><td>{{d['asset_code']}}</td>
      <td class="table-name">{{d['name']}}</td><td class="serial">{{d['serial']}}</td>
      <td>{{'Có' if d['scan_count'] else 'Chưa điểm danh'}}</td>
      <td>{{d['scanned_at'] or ''}}</td><td>{{d['scanned_by'] or ''}}</td></tr>{% endfor %}
      </table></div><div class="pager">
      {% if number>1 %}<a class="btn secondary" href="{{url_for('results',round_id=r['id'],q=q,filter=status_filter,page=1)}}">Đầu</a>
      <a class="btn secondary" href="{{url_for('results',round_id=r['id'],q=q,filter=status_filter,page=number-1)}}">← Trước</a>{% endif %}
      <span>Trang {{number}}/{{pages}}</span>
      {% if number<pages %}<a class="btn secondary" href="{{url_for('results',round_id=r['id'],q=q,filter=status_filter,page=number+1)}}">Tiếp →</a>
      <a class="btn secondary" href="{{url_for('results',round_id=r['id'],q=q,filter=status_filter,page=pages)}}">Cuối</a>{% endif %}
      </div></div>""",r=r,total=total,counted=counted,rows=rows,q=q,
      number=number,pages=pages,all_count=len(all_rows),status_filter=status_filter)

@app.route("/report/<int:round_id>")
@permission_required("results_export")
def report(round_id):
    with db() as con:
        r=con.execute("SELECT * FROM inventory_rounds WHERE id=?",(round_id,)).fetchone()
        if not r:return "Không tìm thấy đợt kiểm kê",404
        rows=round_rows(con,round_id)
    wb=Workbook();ws=wb.active;ws.title="Kết quả kiểm kê"
    ws.append(EXCEL_HEADERS+["Kết quả","Thời gian quét","Tài khoản quét"])
    for cell in ws[1]:cell.font=Font(bold=True)
    for d in rows:
        append_excel_values(ws,[d["stt"],d["asset_code"],d["name"],d["serial"],
                                *[d[field] for field in EXTRA_FIELDS],
                                "Có" if d["scan_count"] else "Chưa điểm danh",
                                d["scanned_at"] or "",d["scanned_by"] or ""])
    for col,width in zip("ABCDEFGHIJKLMN",[10,22,42,40,24,22,20,16,16,20,22,20,23,22]):
        ws.column_dimensions[col].width=width
    ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
    output=io.BytesIO();wb.save(output);output.seek(0)
    with db() as con:
        write_audit(con,"Xuất kết quả Excel",target_type="Báo cáo",target_id=round_id,target_label=r["name"],details=f"Xuất báo cáo đợt với {len(rows)} dòng.")
    return send_file(output,as_attachment=True,
                     download_name=report_filename(r["name"],round_id))

# Khởi tạo CSDL ngay khi module được Gunicorn nạp.
# Trên Render, Gunicorn không chạy khối __main__.
init_db()

if __name__=="__main__":
    # Bản chạy thử trên máy tính. Dùng HTTPS và cấu hình bảo mật trước khi đưa lên Internet.
    app.run(host="127.0.0.1",port=5056,debug=False,use_reloader=False,threaded=False)

