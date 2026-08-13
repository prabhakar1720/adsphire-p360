from __future__ import annotations

import functools
import hmac
import io
import json
import mimetypes
import os
import posixpath
import re
import secrets
import tempfile
import urllib.error
import urllib.parse
import urllib.request
import uuid
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

from flask import (
    Flask,
    jsonify,
    redirect,
    render_template_string,
    request,
    send_file,
    send_from_directory,
    session,
    url_for,
)

BASE_DIR = Path(__file__).resolve().parent
DASHBOARD_FILE = BASE_DIR / "dashboard.html"
TARGETS_FILE = BASE_DIR / "targets.html"
EXCEL_FILE = BASE_DIR / "P360_Config_Template.xlsx"
README_FILE = BASE_DIR / "README.md"

# Render persistent disk should be mounted at /data. Local runs use ./data.
configured_data_dir = os.getenv("DATA_DIR", "").strip()
if configured_data_dir:
    DATA_DIR = Path(configured_data_dir)
elif Path("/data").exists() and os.access("/data", os.W_OK):
    DATA_DIR = Path("/data")
else:
    DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)

TOKEN_FILE = DATA_DIR / "prt_tokens.json"
DATA_ACCOUNTS_FILE = DATA_DIR / "data_accounts.json"
TARGET_CAMPAIGNS_FILE = DATA_DIR / "target_campaigns.json"
SAVED_APPS_FILE = DATA_DIR / "saved_apps.json"
SAVED_EXCEL_META_FILE = DATA_DIR / "saved_excel_config.json"
SAVED_EXCEL_WORKBOOK_FILE = DATA_DIR / "saved_excel_config.xlsx"
MAX_UPLOAD_BYTES = 10 * 1024 * 1024
MAX_UNZIPPED_BYTES = 50 * 1024 * 1024
MAX_ZIP_ENTRIES = 250
AF_HOST = "https://hq1.appsflyer.com"
RAW_REPORTS = {
    "installs_report",
    "in_app_events_report",
    "postbacks",
    "in-app-events-postbacks",
    "blocked_installs_report",
    "detection",
}
AGG_REPORTS = {"partners_by_date_report"}
ALLOWED_TIMEZONES = {"UTC", "Asia/Kolkata", "Asia/Manila", "Asia/Jakarta"}

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY") or secrets.token_hex(32)
app.permanent_session_lifetime = timedelta(hours=12)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=bool(os.getenv("RENDER")),
    MAX_CONTENT_LENGTH=MAX_UPLOAD_BYTES,
)

TEAM_PASSWORD = os.getenv("TEAM_PASSWORD", "").strip()
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "").strip()


def load_json(path: Path, default):
    try:
        with path.open("r", encoding="utf-8") as handle:
            return json.load(handle)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return default


def save_json(path: Path, payload) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=path.name + ".", dir=str(path.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
            handle.flush()
            os.fsync(handle.fileno())
        os.chmod(temp_name, 0o600)
        os.replace(temp_name, path)
    finally:
        if os.path.exists(temp_name):
            os.unlink(temp_name)


def target_accounts() -> list[dict]:
    """Return saved data accounts plus legacy PRT mappings without exposing tokens."""
    saved = load_json(DATA_ACCOUNTS_FILE, [])
    if not isinstance(saved, list):
        saved = []
    legacy = load_json(TOKEN_FILE, {})
    known = {(str(row.get("account_type", "")), str(row.get("identifier", ""))) for row in saved}
    for prt, token in legacy.items():
        if ("agency", str(prt)) in known:
            continue
        saved.append({
            "id": f"legacy:{prt}",
            "label": f"{prt} (legacy agency token)",
            "account_type": "agency",
            "identifier": str(prt),
            "pids": [],
            "token": str(token),
            "enabled": True,
            "legacy": True,
        })
    return saved


def public_target_account(row: dict) -> dict:
    return {
        "id": str(row.get("id", "")),
        "label": str(row.get("label", "")),
        "account_type": str(row.get("account_type", "agency")),
        "identifier": str(row.get("identifier", "")),
        "pids": row.get("pids", []) if isinstance(row.get("pids", []), list) else [],
        "enabled": bool(row.get("enabled", True)),
        "legacy": bool(row.get("legacy", False)),
        "token_configured": bool(row.get("token")),
        "token_masked": "••••••••" + str(row.get("token", ""))[-6:] if row.get("token") else "",
    }


def normalize_data_account(form) -> tuple[dict | None, str | None]:
    label = str(form.get("account_label", "")).strip()
    account_type = str(form.get("account_type", "agency")).strip().lower()
    identifier = str(form.get("account_identifier", "")).strip()
    token = str(form.get("account_token", "")).strip()
    pids = sorted({part.strip() for part in str(form.get("account_pids", "")).split(",") if part.strip()})
    if account_type not in {"agency", "adnetwork"}:
        return None, "Account type must be Agency or Ad Network."
    if not label or not identifier or len(token) < 20:
        return None, "Enter an account name, identifier and valid AppsFlyer V2 token."
    existing = load_json(DATA_ACCOUNTS_FILE, [])
    match = next((row for row in existing if row.get("account_type") == account_type and row.get("identifier") == identifier), None)
    return {
        "id": str(match.get("id") if match else uuid.uuid4()),
        "label": label,
        "account_type": account_type,
        "identifier": identifier,
        "pids": pids,
        "token": token,
        "enabled": bool(match.get("enabled", True)) if match else True,
    }, None


def normalize_target_campaign(form, accounts: list[dict]) -> tuple[dict | None, str | None]:
    name = str(form.get("campaign_name", "")).strip()
    app_id = str(form.get("campaign_app_id", "")).strip()
    pid = str(form.get("campaign_pid", "")).strip()
    credential_id = str(form.get("campaign_credential_id", "")).strip()
    prt = str(form.get("campaign_prt", "")).strip()
    billable_type = str(form.get("billable_type", "event")).strip().lower()
    billable_event = str(form.get("billable_event", "")).strip()
    count_method = str(form.get("count_method", "unique_users")).strip().lower()
    timezone_name = str(form.get("campaign_timezone", "Asia/Kolkata")).strip() or "Asia/Kolkata"
    try:
        daily_target = int(str(form.get("daily_target", "0")).strip())
    except ValueError:
        daily_target = 0
    if not name or not app_id:
        return None, "Campaign name and App ID are required. Leave PID blank to track all PIDs for the app."
    if credential_id not in {str(row.get("id")) for row in accounts}:
        return None, "Select a configured AppsFlyer data account."
    if billable_type not in {"install", "event"}:
        return None, "Billable type must be Install or In-app event."
    if billable_type == "event" and not billable_event:
        return None, "Billable Event Name is required for an event campaign."
    if count_method not in {"unique_users", "event_counter"}:
        return None, "Invalid billable counting method."
    if daily_target < 1:
        return None, "Daily target must be at least 1."
    try:
        import zoneinfo
        zoneinfo.ZoneInfo(timezone_name)
    except Exception:
        return None, "Invalid IANA timezone."
    saved = load_json(TARGET_CAMPAIGNS_FILE, [])
    match = next((row for row in saved if str(row.get("campaign_name", "")).casefold() == name.casefold() and row.get("app_id") == app_id and row.get("pid") == pid), None)
    return {
        "id": str(match.get("id") if match else uuid.uuid4()),
        "enabled": bool(match.get("enabled", True)) if match else True,
        "campaign_name": name,
        "offer_id": str(form.get("offer_id", "")).strip(),
        "app_id": app_id,
        "pid": pid,
        "prt": prt,
        "credential_id": credential_id,
        "campaign_filter": str(form.get("campaign_filter", "")).strip(),
        "billable_type": billable_type,
        "billable_event": billable_event,
        "count_method": count_method,
        "daily_target": daily_target,
        "timezone": timezone_name,
    }, None




def save_bytes(path: Path, payload: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=path.name + ".", dir=str(path.parent))
    try:
        with os.fdopen(fd, "wb") as handle:
            handle.write(payload)
            handle.flush()
            os.fsync(handle.fileno())
        os.chmod(temp_name, 0o600)
        os.replace(temp_name, path)
    finally:
        if os.path.exists(temp_name):
            os.unlink(temp_name)


def safe_upload_name(value: str | None) -> str:
    raw = urllib.parse.unquote(value or "").replace("\\", "/")
    name = raw.rsplit("/", 1)[-1].strip()
    if not name.lower().endswith(".xlsx"):
        name = "P360_Config.xlsx"
    name = re.sub(r"[^A-Za-z0-9._ -]+", "_", name)[:120]
    return name or "P360_Config.xlsx"

def csrf_token() -> str:
    token = session.get("csrf_token")
    if not token:
        token = secrets.token_urlsafe(24)
        session["csrf_token"] = token
    return token


def valid_csrf(value: str | None) -> bool:
    expected = session.get("csrf_token", "")
    return bool(value and expected and hmac.compare_digest(value, expected))


def login_required(view):
    @functools.wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("team_authenticated"):
            if request.path.startswith("/api/") or request.path.startswith("/af/") or request.path == "/config/import":
                return jsonify({"error": "Authentication required"}), 401
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)

    return wrapped


def admin_required(view):
    @functools.wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("team_authenticated"):
            return redirect(url_for("login", next="/admin"))
        if not session.get("admin_authenticated"):
            return redirect(url_for("admin_login"))
        return view(*args, **kwargs)

    return wrapped


LOGIN_HTML = """
<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><title>Adsphire P360 Login</title>
<style>body{margin:0;background:#fff8f2;color:#2a1d12;font-family:Inter,system-ui;display:grid;place-items:center;min-height:100vh}.card{width:min(410px,calc(100% - 32px));background:white;border:1px solid #f0d2b8;border-radius:18px;padding:30px;box-shadow:0 18px 45px rgba(245,130,32,.12)}h1{margin:0 0 8px;font-size:25px}h1 span{color:#f58220}p{color:#765a45;line-height:1.5}label{display:block;font-size:12px;font-weight:700;margin:20px 0 7px;text-transform:uppercase;letter-spacing:.08em}input{width:100%;box-sizing:border-box;padding:12px;border:1px solid #e6bea0;border-radius:10px;font-size:15px}button{width:100%;margin-top:14px;padding:12px;border:0;border-radius:10px;background:#f58220;color:#24170d;font-weight:800;cursor:pointer}.error{background:#fff0ee;color:#a33f34;border-left:3px solid #e65353;padding:10px;border-radius:8px;margin-top:12px}</style></head>
<body><form class="card" method="post"><h1>Adsphire <span>·</span> P360</h1><p>Private team dashboard</p><input type="hidden" name="csrf" value="{{ csrf }}"><input type="hidden" name="next" value="{{ next_url }}"><label>Team password</label><input name="password" type="password" required autofocus><button>Sign in</button>{% if error %}<div class="error">{{ error }}</div>{% endif %}</form></body></html>
"""

ADMIN_LOGIN_HTML = """
<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><title>P360 Admin Login</title>
<style>body{margin:0;background:#fff8f2;color:#2a1d12;font-family:Inter,system-ui;display:grid;place-items:center;min-height:100vh}.card{width:min(410px,calc(100% - 32px));background:white;border:1px solid #f0d2b8;border-radius:18px;padding:30px;box-shadow:0 18px 45px rgba(245,130,32,.12)}h1{margin:0 0 8px;font-size:23px}p{color:#765a45}label{display:block;font-size:12px;font-weight:700;margin:20px 0 7px;text-transform:uppercase;letter-spacing:.08em}input{width:100%;box-sizing:border-box;padding:12px;border:1px solid #e6bea0;border-radius:10px;font-size:15px}button{width:100%;margin-top:14px;padding:12px;border:0;border-radius:10px;background:#f58220;color:#24170d;font-weight:800}.error{background:#fff0ee;color:#a33f34;border-left:3px solid #e65353;padding:10px;border-radius:8px;margin-top:12px}</style></head>
<body><form class="card" method="post"><h1>Admin access</h1><p>Manage PRT-to-token mappings.</p><input type="hidden" name="csrf" value="{{ csrf }}"><label>Admin password</label><input name="password" type="password" required autofocus><button>Open admin</button>{% if error %}<div class="error">{{ error }}</div>{% endif %}</form></body></html>
"""

ADMIN_HTML = """
<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><title>Adsphire P360 Admin</title>
<style>:root{--orange:#f58220;--bg:#fff8f2;--line:#f0d2b8;--ink:#2a1d12;--dim:#765a45;--red:#e65353;--green:#10a978}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:Inter,system-ui}.wrap{max-width:1180px;margin:auto;padding:28px 20px}header{display:flex;justify-content:space-between;gap:15px;align-items:center;flex-wrap:wrap}h1{font-size:25px;margin:0}h1 span{color:var(--orange)}a{color:var(--orange);font-weight:700;text-decoration:none}.panel{background:#fff;border:1px solid var(--line);border-radius:15px;padding:20px;margin-top:18px;box-shadow:0 10px 25px rgba(245,130,32,.08)}.grid{display:grid;grid-template-columns:1fr 2fr auto;gap:10px;align-items:end}.wide-grid{display:grid;grid-template-columns:repeat(4,minmax(150px,1fr));gap:10px;align-items:end}.span2{grid-column:span 2}@media(max-width:900px){.wide-grid{grid-template-columns:repeat(2,minmax(150px,1fr))}}@media(max-width:700px){.grid,.wide-grid{grid-template-columns:1fr}.span2{grid-column:auto}}label{display:block;font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:var(--dim);font-weight:800;margin-bottom:6px}input,select{width:100%;padding:11px;border:1px solid var(--line);border-radius:9px;font:inherit;background:#fff}button{padding:11px 16px;border:0;border-radius:9px;font-weight:800;cursor:pointer;background:var(--orange);color:#24170d}.danger{background:#fff;color:var(--red);border:1px solid #efb7b1}.row{display:grid;grid-template-columns:1.2fr 1fr 1fr auto;gap:12px;align-items:center;padding:12px 0;border-bottom:1px solid #f3e3d6}.campaign-row{grid-template-columns:1.3fr 1fr .8fr 1.1fr auto}.row:last-child{border-bottom:0}.token{font-family:monospace;color:var(--green)}.msg{padding:10px;border-radius:8px;background:#fff4e9;color:var(--dim);margin-top:12px}.hint{font-size:13px;color:var(--dim);line-height:1.55}.badge{display:inline-block;padding:4px 8px;border-radius:999px;background:#fff1e5;color:#8b4b18;font-size:11px;font-weight:800;text-transform:uppercase}.badge.paused{background:#fff0ee;color:var(--red)}.actions-inline{display:flex;gap:7px}.actions-inline button{padding:8px 10px}small{color:var(--dim)}</style></head>
<body><div class="wrap"><header><div><h1>Adsphire <span>·</span> P360 Admin</h1><div class="hint">Tokens stay on the server and are never included in the Excel file or sent to the browser.</div></div><div><a href="/">P360</a> &nbsp; · &nbsp; <a href="/targets">Target Dashboard</a> &nbsp; · &nbsp; <a href="/logout">Logout</a></div></header>
<div class="panel"><h2>Add or update a PRT token</h2><form method="post" class="grid"><input type="hidden" name="csrf" value="{{ csrf }}"><div><label>PRT ID</label><input name="prt" placeholder="adsphirein749" required></div><div><label>AppsFlyer V2 API token</label><input name="token" type="password" placeholder="Paste token" required></div><button name="action" value="save">Save token</button></form>{% if message %}<div class="msg">{{ message }}</div>{% endif %}</div>
<div class="panel"><h2>Configured PRTs</h2>{% if mappings %}{% for prt, masked in mappings %}<form method="post" class="row"><input type="hidden" name="csrf" value="{{ csrf }}"><input type="hidden" name="prt" value="{{ prt }}"><strong>{{ prt }}</strong><span class="token">{{ masked }}</span><button class="danger" name="action" value="delete" onclick="return confirm('Delete this token mapping?')">Delete</button></form>{% endfor %}{% else %}<p class="hint">No PRT tokens configured yet.</p>{% endif %}</div>
<div class="panel"><h2>AppsFlyer Data Accounts</h2><p class="hint">Add one reusable credential for each Agency or Ad Network account. Saving the same account type and identifier updates it.</p><form method="post" class="wide-grid"><input type="hidden" name="csrf" value="{{ csrf }}"><div><label>Account Name</label><input name="account_label" placeholder="Adsphire Ad Network" required></div><div><label>Account Type</label><select name="account_type"><option value="agency">Agency</option><option value="adnetwork">Ad Network</option></select></div><div><label>PRT / Partner ID</label><input name="account_identifier" placeholder="adsphirein749" required></div><div><label>Associated PIDs</label><input name="account_pids" placeholder="kreditgator_int, crichit67_int"></div><div class="span2"><label>AppsFlyer V2 API Token</label><input name="account_token" type="password" placeholder="Paste token" required></div><div><button name="action" value="save_data_account">Save Data Account</button></div></form>{% if account_rows %}{% for row in account_rows %}<form method="post" class="row"><input type="hidden" name="csrf" value="{{ csrf }}"><input type="hidden" name="account_id" value="{{ row.id }}"><strong>{{ row.label }}<br><small>{{ row.identifier }}</small></strong><span>{{ row.pids|join(', ') or 'PIDs selected per campaign' }}</span><span><span class="badge">{{ row.account_type }}</span> <span class="token">{{ row.token_masked }}</span></span><button class="danger" name="action" value="delete_data_account" onclick="return confirm('Delete this data account?')">Delete</button></form>{% endfor %}{% endif %}</div>
<div class="panel"><h2>Campaign Daily Targets</h2><p class="hint">The billing KPI drives achievement. Clicks and impressions use aggregate data; installs and events use raw → postback → aggregate fallback.</p><form method="post" class="wide-grid"><input type="hidden" name="csrf" value="{{ csrf }}"><div><label>Campaign Name</label><input name="campaign_name" required></div><div><label>Offer ID</label><input name="offer_id" placeholder="Optional"></div><div><label>App ID</label><input name="campaign_app_id" required></div><div><label>Media Source / PID</label><input name="campaign_pid" required></div><div><label>PRT</label><input name="campaign_prt" placeholder="Optional for ad network"></div><div><label>Data Account</label><select name="campaign_credential_id" required><option value="">Select account</option>{% for row in account_options %}<option value="{{ row.id }}">{{ row.label }} · {{ row.account_type }}</option>{% endfor %}</select></div><div><label>AF Campaign Filter</label><input name="campaign_filter" placeholder="Optional exact match"></div><div><label>Billing KPI</label><select name="billable_type"><option value="event">In-app event</option><option value="install">Install</option></select></div><div><label>Billable Event Name</label><input name="billable_event" placeholder="Exact case-sensitive name"></div><div><label>Event Counting</label><select name="count_method"><option value="unique_users">Unique Users</option><option value="event_counter">Event Counter</option></select></div><div><label>Daily Target</label><input name="daily_target" type="number" min="1" required></div><div><label>Reporting Timezone</label><input name="campaign_timezone" value="Asia/Kolkata" required></div><div><button name="action" value="save_target_campaign">Save Campaign Target</button></div></form>{% if campaign_rows %}{% for row in campaign_rows %}<form method="post" class="row campaign-row"><input type="hidden" name="csrf" value="{{ csrf }}"><input type="hidden" name="campaign_id" value="{{ row.id }}"><strong>{{ row.campaign_name }}<br><small>{{ row.app_id }}</small></strong><span>{{ row.pid }}{% if row.prt %} · {{ row.prt }}{% endif %}</span><span>Target: <b>{{ row.daily_target }}</b></span><span>{{ row.billable_event or 'Installs' }}<br><span class="badge{% if not row.get('enabled', True) %} paused{% endif %}">{{ 'Enabled' if row.get('enabled', True) else 'Paused' }}</span></span><div class="actions-inline"><button name="action" value="toggle_target_campaign">{{ 'Pause' if row.get('enabled', True) else 'Enable' }}</button><button class="danger" name="action" value="delete_target_campaign" onclick="return confirm('Delete this campaign target?')">Delete</button></div></form>{% endfor %}{% endif %}</div>
<div class="panel"><h2>Storage status</h2><p class="hint">Data directory: <code>{{ data_dir }}</code><br>For Render, attach a persistent disk at <code>/data</code>. Without a disk, admin changes can be lost after a restart or redeploy.</p></div></div></body></html>
"""
ADMIN_HTML = ADMIN_HTML.replace(
    '<label>Media Source / PID</label><input name="campaign_pid" required>',
    '<label>Media Source / PID (optional)</label><input name="campaign_pid" placeholder="Leave blank for all PIDs / app total">',
).replace(
    "{{ row.pid }}{% if row.prt %}",
    "{{ row.pid or 'All PIDs' }}{% if row.prt %}",
)


@app.after_request
def security_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "no-referrer"
    response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
    response.headers["Cache-Control"] = "no-store, max-age=0"
    response.headers["Content-Security-Policy"] = "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'"
    return response


@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("team_authenticated"):
        return redirect(request.args.get("next") or "/")
    error = ""
    if request.method == "POST":
        if not valid_csrf(request.form.get("csrf")):
            error = "Session expired. Please retry."
        elif not TEAM_PASSWORD:
            error = "TEAM_PASSWORD is not configured on the server."
        elif hmac.compare_digest(request.form.get("password", ""), TEAM_PASSWORD):
            session.clear()
            session.permanent = True
            session["team_authenticated"] = True
            csrf_token()
            next_url = request.form.get("next") or "/"
            if not next_url.startswith("/") or next_url.startswith("//"):
                next_url = "/"
            return redirect(next_url)
        else:
            error = "Incorrect password."
    return render_template_string(LOGIN_HTML, error=error, csrf=csrf_token(), next_url=request.args.get("next", "/"))


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/admin/login", methods=["GET", "POST"])
@login_required
def admin_login():
    error = ""
    if request.method == "POST":
        if not valid_csrf(request.form.get("csrf")):
            error = "Session expired. Please retry."
        elif not ADMIN_PASSWORD:
            error = "ADMIN_PASSWORD is not configured on the server."
        elif hmac.compare_digest(request.form.get("password", ""), ADMIN_PASSWORD):
            session["admin_authenticated"] = True
            return redirect(url_for("admin"))
        else:
            error = "Incorrect admin password."
    return render_template_string(ADMIN_LOGIN_HTML, error=error, csrf=csrf_token())


@app.route("/admin", methods=["GET", "POST"])
@admin_required
def admin():
    message = ""
    mappings = load_json(TOKEN_FILE, {})
    if request.method == "POST":
        if not valid_csrf(request.form.get("csrf")):
            message = "Session expired. Reload and retry."
        else:
            action = request.form.get("action", "save")
            prt = request.form.get("prt", "").strip()
            if action == "save_data_account":
                row, error = normalize_data_account(request.form)
                if error:
                    message = error
                else:
                    rows = load_json(DATA_ACCOUNTS_FILE, [])
                    rows = [item for item in rows if item.get("id") != row["id"]]
                    rows.append(row)
                    save_json(DATA_ACCOUNTS_FILE, rows)
                    message = f"Saved {row['label']} data account."
            elif action == "delete_data_account":
                account_id = request.form.get("account_id", "").strip()
                rows = [item for item in load_json(DATA_ACCOUNTS_FILE, []) if item.get("id") != account_id]
                save_json(DATA_ACCOUNTS_FILE, rows)
                message = "Deleted data account."
            elif action == "save_target_campaign":
                row, error = normalize_target_campaign(request.form, target_accounts())
                if error:
                    message = error
                else:
                    rows = load_json(TARGET_CAMPAIGNS_FILE, [])
                    rows = [item for item in rows if item.get("id") != row["id"]]
                    rows.append(row)
                    save_json(TARGET_CAMPAIGNS_FILE, rows)
                    message = f"Saved daily target for {row['campaign_name']}."
            elif action == "delete_target_campaign":
                campaign_id = request.form.get("campaign_id", "").strip()
                rows = [item for item in load_json(TARGET_CAMPAIGNS_FILE, []) if item.get("id") != campaign_id]
                save_json(TARGET_CAMPAIGNS_FILE, rows)
                message = "Deleted campaign target."
            elif action == "toggle_target_campaign":
                campaign_id = request.form.get("campaign_id", "").strip()
                rows = load_json(TARGET_CAMPAIGNS_FILE, [])
                changed = next((item for item in rows if item.get("id") == campaign_id), None)
                if changed:
                    changed["enabled"] = not bool(changed.get("enabled", True))
                    save_json(TARGET_CAMPAIGNS_FILE, rows)
                    message = f"{'Enabled' if changed['enabled'] else 'Paused'} {changed.get('campaign_name', 'campaign')}."
                else:
                    message = "Campaign target was not found."
            elif action == "delete":
                mappings.pop(prt, None)
                save_json(TOKEN_FILE, mappings)
                message = f"Deleted token mapping for {prt}."
            else:
                token = request.form.get("token", "").strip()
                if not prt or len(token) < 20:
                    message = "Enter a valid PRT ID and AppsFlyer V2 token."
                else:
                    mappings[prt] = token
                    save_json(TOKEN_FILE, mappings)
                    message = f"Saved token mapping for {prt}."
    mappings = load_json(TOKEN_FILE, {})
    masked = [(prt, "••••••••" + token[-6:]) for prt, token in sorted(mappings.items())]
    stored_accounts = load_json(DATA_ACCOUNTS_FILE, [])
    account_rows = [public_target_account(row) for row in stored_accounts]
    account_options = [public_target_account(row) for row in target_accounts() if row.get("enabled", True)]
    campaign_rows = load_json(TARGET_CAMPAIGNS_FILE, [])
    return render_template_string(
        ADMIN_HTML,
        mappings=masked,
        account_rows=account_rows,
        account_options=account_options,
        campaign_rows=campaign_rows,
        message=message,
        csrf=csrf_token(),
        data_dir=str(DATA_DIR),
    )


@app.route("/")
@login_required
def dashboard():
    return send_file(DASHBOARD_FILE, mimetype="text/html")


@app.route("/targets")
@login_required
def targets_dashboard():
    return send_file(TARGETS_FILE, mimetype="text/html")


@app.route("/api/target-config")
@login_required
def target_config():
    accounts = []
    for row in target_accounts():
        if not row.get("enabled", True):
            continue
        public = public_target_account(row)
        public.pop("token_masked", None)
        accounts.append(public)
    campaigns = [row for row in load_json(TARGET_CAMPAIGNS_FILE, []) if row.get("enabled", True)]
    return jsonify({"accounts": accounts, "campaigns": campaigns})


@app.route("/P360_Config_Template.xlsx")
@login_required
def excel_template():
    return send_file(EXCEL_FILE, as_attachment=True, download_name="P360_Config_Template.xlsx")


@app.route("/README.md")
@login_required
def readme():
    return send_file(README_FILE, mimetype="text/markdown")


@app.route("/health")
def health():
    return jsonify({"ok": True, "service": "Adsphire P360 Hosted Team", "saved_excel": SAVED_EXCEL_META_FILE.exists()})


@app.route("/api/prts")
@login_required
def api_prts():
    mappings = load_json(TOKEN_FILE, {})
    return jsonify({"prts": sorted(mappings.keys())})


def validate_saved_app(payload: dict) -> tuple[dict | None, str | None]:
    app_name = str(payload.get("app_name", "")).strip()
    app_id = str(payload.get("app_id", "")).strip()
    pid = str(payload.get("pid", "")).strip()
    prt = str(payload.get("prt", "")).strip()
    timezone = str(payload.get("timezone", "UTC")).strip() or "UTC"
    notes = str(payload.get("notes", "")).strip()
    if not app_id or any(ch in app_id for ch in "/\\"):
        return None, "A valid App ID is required."
    if not pid:
        return None, "Media Source / PID is required."
    if not prt:
        return None, "PRT ID is required because the server selects the token by PRT."
    try:
        import zoneinfo
        zoneinfo.ZoneInfo(timezone)
    except Exception:
        return None, "Invalid IANA timezone."
    return {
        "id": str(payload.get("id") or uuid.uuid4()),
        "enabled": True,
        "app_name": app_name or app_id,
        "app_id": app_id,
        "pid": pid,
        "prt": prt,
        "timezone": timezone,
        "notes": notes,
    }, None


@app.route("/api/saved-apps", methods=["GET", "POST"])
@login_required
def api_saved_apps():
    rows = load_json(SAVED_APPS_FILE, [])
    if request.method == "GET":
        return jsonify({"apps": rows, "csrf": csrf_token()})
    payload = request.get_json(silent=True) or {}
    if not valid_csrf(request.headers.get("X-CSRF-Token")):
        return jsonify({"error": "Session expired. Reload and retry."}), 403
    row, error = validate_saved_app(payload)
    if error:
        return jsonify({"error": error}), 400
    rows = [item for item in rows if item.get("id") != row["id"]]
    rows.append(row)
    save_json(SAVED_APPS_FILE, rows)
    return jsonify({"ok": True, "app": row})


@app.route("/api/saved-apps/<row_id>", methods=["DELETE"])
@login_required
def delete_saved_app(row_id: str):
    if not valid_csrf(request.headers.get("X-CSRF-Token")):
        return jsonify({"error": "Session expired. Reload and retry."}), 403
    rows = load_json(SAVED_APPS_FILE, [])
    new_rows = [item for item in rows if item.get("id") != row_id]
    save_json(SAVED_APPS_FILE, new_rows)
    return jsonify({"ok": True})


# ----- XLSX parser (stdlib only) -----
def column_index(cell_ref: str) -> int:
    letters = re.match(r"[A-Za-z]+", cell_ref or "")
    if not letters:
        return 0
    result = 0
    for ch in letters.group(0).upper():
        result = result * 26 + (ord(ch) - 64)
    return result - 1


def safe_zip_member(name: str) -> bool:
    normalized = posixpath.normpath(name.replace("\\", "/"))
    return not (normalized.startswith("../") or normalized == ".." or normalized.startswith("/") or ":" in normalized.split("/")[0])


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    name = "xl/sharedStrings.xml"
    if name not in zf.namelist():
        return []
    root = ET.fromstring(zf.read(name))
    values: list[str] = []
    for si in root:
        parts = []
        for node in si.iter():
            if node.tag.endswith("}t") or node.tag == "t":
                parts.append(node.text or "")
        values.append("".join(parts))
    return values


def read_sheet_rows(zf: zipfile.ZipFile, sheet_path: str, shared: list[str]) -> list[list[object]]:
    if sheet_path not in zf.namelist():
        return []
    root = ET.fromstring(zf.read(sheet_path))
    rows_out: list[list[object]] = []
    for row in root.iter():
        if not (row.tag.endswith("}row") or row.tag == "row"):
            continue
        values: dict[int, object] = {}
        max_col = -1
        for cell in row:
            if not (cell.tag.endswith("}c") or cell.tag == "c"):
                continue
            idx = column_index(cell.attrib.get("r", "A1"))
            max_col = max(max_col, idx)
            cell_type = cell.attrib.get("t", "")
            value_node = None
            inline_parts: list[str] = []
            for node in cell.iter():
                if node.tag.endswith("}v") or node.tag == "v":
                    value_node = node
                elif node.tag.endswith("}t") or node.tag == "t":
                    inline_parts.append(node.text or "")
            raw = value_node.text if value_node is not None else None
            if cell_type == "s":
                try:
                    value: object = shared[int(raw or "0")]
                except (ValueError, IndexError):
                    value = ""
            elif cell_type == "inlineStr":
                value = "".join(inline_parts)
            elif cell_type == "b":
                value = raw == "1"
            elif cell_type in {"str", "e"}:
                value = raw or ""
            else:
                if raw is None:
                    value = "".join(inline_parts)
                else:
                    try:
                        number = float(raw)
                        value = int(number) if number.is_integer() else number
                    except ValueError:
                        value = raw
            values[idx] = value
        if max_col >= 0:
            rows_out.append([values.get(i, "") for i in range(max_col + 1)])
    while rows_out and not any(str(v).strip() for v in rows_out[-1]):
        rows_out.pop()
    return rows_out


def parse_xlsx(data: bytes) -> dict[str, list[list[object]]]:
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as exc:
        raise ValueError("The file is not a valid .xlsx workbook.") from exc
    infos = zf.infolist()
    if len(infos) > MAX_ZIP_ENTRIES or sum(info.file_size for info in infos) > MAX_UNZIPPED_BYTES:
        raise ValueError("The workbook is unusually large.")
    if any(not safe_zip_member(info.filename) for info in infos):
        raise ValueError("The workbook contains an unsafe path.")
    if not {"xl/workbook.xml", "xl/_rels/workbook.xml.rels"}.issubset(set(zf.namelist())):
        raise ValueError("The workbook structure is incomplete.")
    shared = read_shared_strings(zf)
    wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
    rel_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    relationships = {rel.attrib.get("Id", ""): rel.attrib.get("Target", "") for rel in rel_root}
    sheets: dict[str, list[list[object]]] = {}
    rel_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    for node in wb_root.iter():
        if not (node.tag.endswith("}sheet") or node.tag == "sheet"):
            continue
        sheet_name = node.attrib.get("name", "Sheet")
        rel_id = node.attrib.get(rel_ns) or node.attrib.get("id", "")
        target = relationships.get(rel_id, "")
        if not target:
            continue
        path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        if not safe_zip_member(path):
            raise ValueError("The workbook contains an unsafe sheet path.")
        sheets[sheet_name] = read_sheet_rows(zf, path, shared)
    if not sheets:
        raise ValueError("No readable sheets were found.")
    return sheets


@app.route("/api/saved-excel-config", methods=["GET", "DELETE"])
@login_required
def saved_excel_config():
    if request.method == "GET":
        payload = load_json(SAVED_EXCEL_META_FILE, {})
        if not payload or not payload.get("sheets"):
            return jsonify({"saved": False, "csrf": csrf_token()})
        return jsonify({
            "saved": True,
            "filename": payload.get("filename", "P360_Config.xlsx"),
            "saved_at": payload.get("saved_at", ""),
            "sheets": payload.get("sheets", {}),
            "csrf": csrf_token(),
        })

    if not valid_csrf(request.headers.get("X-CSRF-Token")):
        return jsonify({"error": "Session expired. Reload and retry."}), 403
    for path in (SAVED_EXCEL_META_FILE, SAVED_EXCEL_WORKBOOK_FILE):
        try:
            path.unlink()
        except FileNotFoundError:
            pass
        except OSError as exc:
            return jsonify({"error": f"Could not remove saved Excel configuration: {exc}"}), 500
    return jsonify({"ok": True})


@app.route("/api/saved-excel-config/download")
@login_required
def download_saved_excel_config():
    payload = load_json(SAVED_EXCEL_META_FILE, {})
    if not SAVED_EXCEL_WORKBOOK_FILE.exists() or not payload:
        return jsonify({"error": "No saved Excel configuration is available."}), 404
    return send_file(
        SAVED_EXCEL_WORKBOOK_FILE,
        as_attachment=True,
        download_name=payload.get("filename", "P360_Config.xlsx"),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/config/import", methods=["POST"])
@login_required
def import_config():
    if not valid_csrf(request.headers.get("X-CSRF-Token")):
        return jsonify({"error": "Session expired. Reload and retry."}), 403
    body = request.get_data(cache=False)
    if not body:
        return jsonify({"error": "Empty upload."}), 400
    try:
        sheets = parse_xlsx(body)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    filename = safe_upload_name(request.headers.get("X-P360-Filename"))
    saved_at = datetime.now(timezone.utc).isoformat()
    payload = {"filename": filename, "saved_at": saved_at, "sheets": sheets}
    try:
        save_bytes(SAVED_EXCEL_WORKBOOK_FILE, body)
        save_json(SAVED_EXCEL_META_FILE, payload)
    except OSError as exc:
        return jsonify({"error": f"The workbook was valid but could not be saved on the server: {exc}"}), 500
    return jsonify({"saved": True, "filename": filename, "saved_at": saved_at, "sheets": sheets})


@app.route("/af/api/<kind>/export/app/<app_id>/<report>/v5")
@login_required
def appsflyer_proxy(kind: str, app_id: str, report: str):
    if kind not in {"raw-data", "agg-data"}:
        return jsonify({"error": "AppsFlyer endpoint not allowed"}), 403
    if kind == "raw-data" and report not in RAW_REPORTS:
        return jsonify({"error": "Raw report not allowed"}), 403
    if kind == "agg-data" and report not in AGG_REPORTS:
        return jsonify({"error": "Aggregate report not allowed"}), 403
    if not app_id or any(ch in app_id for ch in "/\\"):
        return jsonify({"error": "Invalid App ID"}), 400

    credential_id = request.headers.get("X-AF-Credential-ID", "").strip()
    prt = request.headers.get("X-P360-PRT", "").strip()
    token = ""
    credential_label = prt
    if credential_id:
        account = next((row for row in target_accounts() if str(row.get("id")) == credential_id), None)
        if not account:
            return jsonify({"error": "The selected AppsFlyer data account no longer exists."}), 422
        token = str(account.get("token", ""))
        credential_label = str(account.get("label") or account.get("identifier") or credential_id)
    else:
        mappings = load_json(TOKEN_FILE, {})
        token = mappings.get(prt, "")
        if not prt:
            return jsonify({"error": "PRT ID or AppsFlyer credential ID is required for server-side token selection."}), 422
    if not token:
        return jsonify({"error": f"No AppsFlyer token is configured for {credential_label}. Ask the admin to add it."}), 422

    route = f"/api/{kind}/export/app/{urllib.parse.quote(app_id, safe='')}/{report}/v5"
    target = AF_HOST + route
    query = request.query_string.decode("utf-8", errors="ignore")
    if query:
        target += "?" + query
    outbound = urllib.request.Request(target, method="GET")
    outbound.add_header("Authorization", "Bearer " + token)
    outbound.add_header("Accept", "text/csv")
    outbound.add_header("User-Agent", "Adsphire-P360-Hosted/1.0")
    try:
        with urllib.request.urlopen(outbound, timeout=300) as response:
            data = response.read()
            return app.response_class(data, status=response.status, content_type=response.headers.get("Content-Type", "text/csv; charset=utf-8"))
    except urllib.error.HTTPError as exc:
        data = exc.read()
        return app.response_class(data, status=exc.code, content_type=exc.headers.get("Content-Type", "text/plain; charset=utf-8"))
    except Exception as exc:
        return jsonify({"error": f"AppsFlyer proxy error: {exc}"}), 502


if __name__ == "__main__":
    port = int(os.getenv("PORT", "8787"))
    app.run(host="0.0.0.0", port=port, debug=False)
