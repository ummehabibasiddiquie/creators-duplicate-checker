"""Local creator duplicate checker.

A creator matches on one cleaned Instagram username, taken from the handle
or from the profile link. Validity and client decision are filters on that
saved creator. They do not decide the match.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import hmac
import io
import os
import re
import secrets
import sqlite3
import threading
import traceback
import webbrowser
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import openpyxl
from flask import Flask, jsonify, render_template, request, send_file, session
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from werkzeug.utils import secure_filename

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
DB_PATH = DATA / "master.db"

NON_PROFILE = {
    "p",
    "reel",
    "reels",
    "stories",
    "explore",
    "accounts",
    "about",
    "legal",
    "directory",
    "tv",
    "share",
    "direct",
}

HEADER_SKIP = {
    "instagram",
    "instagramhandle",
    "handle",
    "username",
    "profilename",
    "profileurl",
    "profilelink",
    "country",
    "url",
    "link",
    "validity",
    "client",
    "notes",
    "name",
}

COUNTRY_MAP = {
    "usa": "United States",
    "us": "United States",
    "u.s.": "United States",
    "u.s.a.": "United States",
    "united states": "United States",
    "united states of america": "United States",
    "uk": "United Kingdom",
    "u.k.": "United Kingdom",
    "united kingdom": "United Kingdom",
    "great britain": "United Kingdom",
    "england": "United Kingdom",
    "aus": "Australia",
    "australia": "Australia",
    "canada": "Canada",
    "italy": "Italy",
    "spain": "Spain",
    "chile": "Chile",
    "armenia": "Armenia",
    "guatemala": "Guatemala",
    "ireland": "Ireland",
    "global": "Global",
    "canada/usa": "Canada / United States",
    "canada / usa": "Canada / United States",
    "usa/canada": "Canada / United States",
}

FILE_COLUMNS = [
    ("Instagram Handle", "username"),
    ("Instagram Profile URL", "url"),
    ("Country", "country"),
    ("Validity", "validity"),
    ("Client decision", "client"),
    ("Notes", "notes"),
    ("Date", "found_on"),
    ("Agent name", "agent_name"),
]
IGNORED_FILE_HEADERS = {
    "result",
    "name in link",
    "link check",
    "first seen",
    "last updated",
    "source file",
}
def improper_file(detail):
    return ValueError(
        "This file is not in the proper format. "
        + detail
        + " Download the template and keep these columns: "
        "Instagram Handle, Instagram Profile URL, Country."
    )

VALIDITY_MAP = {
    "valid": "valid",
    "invalid": "invalid",
    "unchecked": "unchecked",
    "not checked": "unchecked",
    "not checked yet": "unchecked",
}

CLIENT_MAP = {
    "accepted": "accepted",
    "accepted by client": "accepted",
    "client accepted": "accepted",
    "rejected": "rejected",
    "not accepted": "rejected",
    "not accepted by client": "rejected",
    "rejected by client": "rejected",
    "client rejected": "rejected",
    "pending": "unset",
    "not decided": "unset",
    "unset": "unset",
}

VALIDITY_LABEL = {
    "unchecked": "Not checked yet",
    "valid": "Valid",
    "invalid": "Invalid",
}
CLIENT_LABEL = {
    "unset": "Not decided",
    "accepted": "Accepted",
    "rejected": "Rejected",
}
LINK_LABEL = {
    "ok": "Handle and link match",
    "mismatch": "Handle and link differ",
    "non_instagram": "Not an Instagram link",
    "missing_url": "No profile link",
    "post_link": "Link is a post, not a profile",
    "confirmed": "Checked by the team",
    "missing_username": "Link has no username",
}
OUTCOME_LABEL = {
    "new": "New — saved to the master list",
    "unchecked": "Already listed — not checked yet",
    "pending_client": "Valid — waiting on the client",
    "invalid": "Invalid — skip",
    "rejected": "Rejected by the client — skip",
    "accepted": "Already accepted — skip",
    "ambiguous": "Matches more than one saved creator",
    "not_instagram": "Not an Instagram profile link",
    "duplicate_in_batch": "Repeated in this list",
    "empty": "Blank row",
}

EXPORT_HEADERS = [
    "Instagram Handle",
    "Instagram Profile URL",
    "Name in link",
    "Country",
    "Validity",
    "Client decision",
    "Notes",
    "Link check",
    "Date",
    "Agent name",
    "Last updated",
    "Source file",
]

app = Flask(__name__)
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)

AGENT_API = {
    ("GET", "/api/creators"),
    ("GET", "/api/me"),
    ("POST", "/api/logout"),
    ("POST", "/api/check"),
    ("POST", "/api/export-check"),
    ("POST", "/api/agent/match"),
    ("POST", "/api/agent/update"),
}


@app.after_request
def disable_cache(response):
    response.headers["Cache-Control"] = "no-store"
    return response


def accounts_ready():
    return bool(os.environ.get("ADMIN_PASSWORD_HASH") and os.environ.get("AGENT_PASSWORD_HASH"))


def hash_password(password):
    salt = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 120000).hex()
    return f"{salt}${digest}"


def password_matches(password, stored):
    try:
        salt, digest = stored.split("$", 1)
    except ValueError:
        return False
    check = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 120000).hex()
    return hmac.compare_digest(check, digest)


def save_env_values(updates):
    path = BASE / "config.env"
    lines = []
    seen = set()
    if path.exists():
        for line in path.read_text(encoding="utf-8").splitlines():
            stripped = line.strip()
            if stripped and not stripped.startswith("#") and "=" in stripped:
                key = stripped.split("=", 1)[0].strip()
                if key in updates:
                    lines.append(f"{key}={updates[key]}")
                    seen.add(key)
                    continue
            lines.append(line)
    for key, value in updates.items():
        if key not in seen:
            lines.append(f"{key}={value}")
    path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")
    for key, value in updates.items():
        os.environ[key] = value


def current_role():
    role = session.get("role")
    return role if role in {"admin", "agent"} else ""


@app.before_request
def require_login():
    if not request.path.startswith("/api/"):
        return None
    if request.path == "/api/me":
        return None
    if request.path in {"/api/login", "/api/setup", "/api/check-one", "/api/check-list", "/api/template.xlsx"}:
        return None
    if not accounts_ready():
        return jsonify(error="Choose the admin and agent passwords first."), 403
    role = current_role()
    if not role:
        return jsonify(error="Sign in first."), 401
    if role == "agent" and (request.method, request.path) not in AGENT_API:
        return jsonify(error="An agent cannot do that."), 403
    return None


@app.get("/api/me")
def me():
    if not accounts_ready():
        return jsonify(setup=True, role="")
    return jsonify(setup=False, role=current_role())


@app.post("/api/setup")
def setup_accounts():
    if accounts_ready():
        return jsonify(error="The accounts are already set up."), 403
    payload = request.get_json(silent=True) or {}
    admin_password = str(payload.get("admin_password") or "")
    agent_password = str(payload.get("agent_password") or "")
    if len(admin_password) < 6 or len(agent_password) < 6:
        return jsonify(error="Use at least 6 characters for each password."), 400
    if admin_password == agent_password:
        return jsonify(error="The admin password and the agent password must be different."), 400
    secret = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
    save_env_values({
        "SECRET_KEY": secret,
        "ADMIN_PASSWORD_HASH": hash_password(admin_password),
        "AGENT_PASSWORD_HASH": hash_password(agent_password),
    })
    app.secret_key = secret
    session.clear()
    session["role"] = "admin"
    return jsonify(role="admin")


@app.post("/api/login")
def login():
    if not accounts_ready():
        return jsonify(error="Choose the admin and agent passwords first."), 403
    payload = request.get_json(silent=True) or {}
    username = cell_str(payload.get("username")).lower()
    password = str(payload.get("password") or "")
    if username not in {"admin", "agent"}:
        return jsonify(error="Sign in as admin or agent."), 400
    stored = os.environ.get("ADMIN_PASSWORD_HASH" if username == "admin" else "AGENT_PASSWORD_HASH")
    if not password_matches(password, stored or ""):
        return jsonify(error="That account or password is wrong."), 401
    session.clear()
    session["role"] = username
    return jsonify(role=username)


@app.post("/api/logout")
def logout():
    session.clear()
    return jsonify(ok=True)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024
write_lock = threading.Lock()


def now_stamp() -> str:
    return datetime.now().replace(microsecond=0).isoformat(sep=" ")


def today_stamp() -> str:
    return datetime.now().date().isoformat()


def clean_token(value) -> str:
    text = "" if value is None else str(value)
    text = text.replace("*", "").strip().lower().lstrip("@")
    text = re.sub(r"\s+", "", text)
    text = text.strip(".,;:/")
    text = re.sub(r"[^a-z0-9._]", "", text)
    if text in NON_PROFILE or text in HEADER_SKIP:
        return ""
    return text[:50]


def country_group(raw) -> str:
    if raw is None or not str(raw).strip():
        return "No country"
    cleaned = re.sub(r"\s+", " ", str(raw).replace("*", " ")).strip(" -/")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    if not cleaned:
        return "No country"
    return COUNTRY_MAP.get(cleaned.lower(), cleaned)


def parse_validity(value, strict=True):
    if value is None:
        return None
    key = re.sub(r"\s+", " ", str(value).strip().lower())
    if not key:
        return None
    if key in VALIDITY_MAP:
        return VALIDITY_MAP[key]
    return None if strict else None


def parse_client(value, strict=True):
    if value is None:
        return None
    key = re.sub(r"\s+", " ", str(value).strip().lower())
    if not key:
        return None
    return CLIENT_MAP.get(key)


def cell_str(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def inspect_url(value):
    text = cell_str(value)
    if not text:
        return None, "empty"
    lower = text.lower()
    looks_like_url = (
        "instagram.com" in lower
        or lower.startswith("http://")
        or lower.startswith("https://")
        or lower.startswith("www.")
    )
    if not looks_like_url:
        token = clean_token(text)
        return (token or None), ("ok" if token else "missing_username")

    url = text if "://" in text else "https://" + text
    try:
        parsed = urlparse(url)
    except ValueError:
        return None, "non_instagram"
    host = (parsed.netloc or "").lower().split("@")[-1].split(":")[0]
    if host.startswith("www."):
        host = host[4:]
    if host not in {"instagram.com", "m.instagram.com", "l.instagram.com"} and not host.endswith(
        ".instagram.com"
    ):
        return None, "non_instagram"
    parts = [part for part in parsed.path.split("/") if part]
    if not parts:
        return None, "missing_username"
    first = clean_token(parts[0])
    if not first or first in NON_PROFILE:
        return None, "post_link"
    return first, "ok"


def identity_from_text(value):
    text = cell_str(value)
    if not text:
        return None
    lower = text.lower()
    if "instagram.com" in lower or lower.startswith("http://") or lower.startswith("https://") or lower.startswith("www."):
        username, kind = inspect_url(text)
        return username if kind == "ok" else None
    token = clean_token(text)
    return token or None


def classify_link(primary, url, url_username, url_kind) -> str:
    if url_kind == "empty" or not cell_str(url):
        return "missing_url"
    if url_kind == "non_instagram":
        return "non_instagram"
    if url_kind == "post_link":
        return "post_link"
    if not url_username:
        return "missing_username"
    if primary and url_username and primary != url_username:
        return "mismatch"
    return "ok"


def canonical_profile_url(username: str) -> str:
    return f"https://www.instagram.com/{username}/"


PAGE_SIZE = 100
SORT_COLUMNS = {
    "username": "username",
    "country": "country_group",
    "validity": "validity",
    "client_status": "client_status",
    "link_status": "link_status",
    "first_seen": "first_seen",
}
LINK_FILTERS = {
    "ok",
    "mismatch",
    "non_instagram",
    "missing_url",
    "post_link",
    "confirmed",
    "missing_username",
}


def load_env_file():
    path = BASE / "config.env"
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


load_env_file()
if os.environ.get("SECRET_KEY"):
    app.secret_key = os.environ["SECRET_KEY"]
USE_MYSQL = bool(os.environ.get("MYSQL_HOST") and os.environ.get("MYSQL_USER") and os.environ.get("MYSQL_DATABASE"))


class QueryResult:
    def __init__(self, rows, lastrowid=None):
        self._rows = rows
        self.lastrowid = lastrowid

    def fetchone(self):
        return self._rows[0] if self._rows else None

    def fetchall(self):
        return list(self._rows)

    def __iter__(self):
        return iter(self._rows)


class DBConn:
    """Same execute/commit surface for the local file and for MySQL."""

    def __init__(self, raw, mysql):
        self.raw = raw
        self.mysql = mysql

    def execute(self, sql, params=()):
        statement = sql.replace("?", "%s") if self.mysql else sql
        cursor = self.raw.cursor()
        cursor.execute(statement, tuple(params) if params else ())
        rows = cursor.fetchall() if cursor.description else []
        result = QueryResult(rows, cursor.lastrowid)
        cursor.close()
        return result

    def commit(self):
        self.raw.commit()

    def close(self):
        self.raw.close()


def connect():
    if USE_MYSQL:
        import pymysql
        from pymysql.cursors import DictCursor

        raw = pymysql.connect(
            host=os.environ["MYSQL_HOST"],
            port=int(os.environ.get("MYSQL_PORT") or 3306),
            user=os.environ["MYSQL_USER"],
            password=os.environ.get("MYSQL_PASSWORD") or "",
            database=os.environ["MYSQL_DATABASE"],
            charset="utf8mb4",
            cursorclass=DictCursor,
            autocommit=False,
        )
        return DBConn(raw, mysql=True)
    raw = sqlite3.connect(DB_PATH)
    raw.row_factory = sqlite3.Row
    raw.execute("PRAGMA journal_mode=WAL")
    raw.execute("PRAGMA synchronous=NORMAL")
    return DBConn(raw, mysql=False)


@contextmanager
def db(write=False):
    if write:
        write_lock.acquire()
    conn = connect()
    try:
        yield conn
        if write:
            conn.commit()
    finally:
        conn.close()
        if write:
            write_lock.release()


MYSQL_SCHEMA = """
CREATE TABLE IF NOT EXISTS creators (
    id INT NOT NULL AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(80) NOT NULL,
    handle_display VARCHAR(120) NOT NULL,
    profile_url TEXT,
    url_username VARCHAR(80) NULL,
    country_raw VARCHAR(160) NULL,
    country_group VARCHAR(160) NULL,
    validity VARCHAR(20) NOT NULL DEFAULT 'unchecked',
    client_status VARCHAR(20) NOT NULL DEFAULT 'unset',
    notes TEXT,
    link_status VARCHAR(40) NOT NULL DEFAULT 'ok',
    first_seen VARCHAR(40) NULL,
    last_updated VARCHAR(40) NULL,
    source_file VARCHAR(255) NULL,
    agent_name VARCHAR(80) NULL,
    UNIQUE KEY uq_creators_username (username),
    KEY idx_creators_url_username (url_username),
    KEY idx_creators_validity (validity),
    KEY idx_creators_client (client_status),
    KEY idx_creators_country (country_group),
    KEY idx_creators_link (link_status)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
"""


def copy_local_file_into_mysql(conn):
    if not DB_PATH.exists():
        return 0
    source = sqlite3.connect(DB_PATH)
    source.row_factory = sqlite3.Row
    try:
        rows = source.execute("SELECT * FROM creators").fetchall()
    except sqlite3.OperationalError:
        return 0
    finally:
        source.close()
    for row in rows:
        conn.execute(
            """
            INSERT INTO creators (
                username, handle_display, profile_url, url_username, country_raw, country_group,
                validity, client_status, notes, link_status, first_seen, last_updated, source_file, agent_name
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                row["username"],
                row["handle_display"],
                row["profile_url"],
                row["url_username"],
                row["country_raw"],
                row["country_group"],
                row["validity"],
                row["client_status"],
                row["notes"] or "",
                row["link_status"],
                row["first_seen"],
                row["last_updated"],
                row["source_file"],
                row["agent_name"] if "agent_name" in row.keys() else "",
            ),
        )
    return len(rows)


def ensure_agent_column(conn):
    if conn.mysql:
        names = {row["Field"] for row in conn.execute("SHOW COLUMNS FROM creators")}
        if "agent_name" not in names:
            conn.execute("ALTER TABLE creators ADD COLUMN agent_name VARCHAR(80) NULL")
        return
    names = {row["name"] for row in conn.execute("PRAGMA table_info(creators)")}
    if "agent_name" not in names:
        conn.execute("ALTER TABLE creators ADD COLUMN agent_name TEXT DEFAULT ''")


def init_db():
    DATA.mkdir(exist_ok=True)
    with db(write=True) as conn:
        if USE_MYSQL:
            conn.execute(MYSQL_SCHEMA)
            print(
                f"Database: MySQL {os.environ['MYSQL_DATABASE']} on {os.environ['MYSQL_HOST']}"
            )
        else:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS creators (
                    id INTEGER PRIMARY KEY,
                    username TEXT NOT NULL UNIQUE,
                    handle_display TEXT NOT NULL,
                    profile_url TEXT,
                    url_username TEXT,
                    country_raw TEXT,
                    country_group TEXT,
                    validity TEXT NOT NULL DEFAULT 'unchecked',
                    client_status TEXT NOT NULL DEFAULT 'unset',
                    notes TEXT DEFAULT '',
                    link_status TEXT NOT NULL DEFAULT 'ok',
                    first_seen TEXT,
                    last_updated TEXT,
                    source_file TEXT,
                    agent_name TEXT DEFAULT ''
                )
                """
            )
            conn.execute("CREATE INDEX IF NOT EXISTS idx_creators_url_username ON creators(url_username)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_creators_validity ON creators(validity)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_creators_client ON creators(client_status)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_creators_country ON creators(country_group)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_creators_link ON creators(link_status)")
            print(f"Database: local file {DB_PATH.name}")
        ensure_agent_column(conn)
        count = conn.execute("SELECT COUNT(*) AS n FROM creators").fetchone()["n"]
        if count == 0 and USE_MYSQL:
            copied = copy_local_file_into_mysql(conn)
            if copied:
                print(f"Copied {copied} creators from the local database file into MySQL.")
                count = copied
        if count == 0:
            seed_file = find_seed_file()
            if seed_file:
                records, warnings = read_table(seed_file.name, seed_file.read_bytes())
                stamp = datetime.fromtimestamp(seed_file.stat().st_mtime).date().isoformat()
                plan_upload(conn, records, seed_file.name, commit=True, seen_on=stamp)
                print(f"Loaded {seed_file.name} ({len(records)} rows).")
                for warning in warnings:
                    print(" ", warning)


def find_seed_file():
    preferred = sorted(BASE.glob("Duplicate_Check*.xlsx"))
    preferred = [path for path in preferred if not path.name.startswith("~$")]
    if preferred:
        return preferred[-1]
    files = [path for path in BASE.glob("*.xlsx") if not path.name.startswith("~$")]
    if len(files) == 1:
        return files[0]
    return None


def creator_public(row) -> dict:
    return {
        "id": row["id"],
        "username": row["username"],
        "handle": row["handle_display"],
        "profile_url": row["profile_url"] or "",
        "url_username": row["url_username"] or "",
        "country": row["country_raw"] or "",
        "country_group": row["country_group"] or "No country",
        "validity": row["validity"],
        "client_status": row["client_status"],
        "notes": row["notes"] or "",
        "link_status": row["link_status"],
        "first_seen": row["first_seen"] or "",
        "last_updated": row["last_updated"] or "",
        "source_file": row["source_file"] or "",
        "agent_name": row["agent_name"] or "" if "agent_name" in row.keys() else "",
        "open_url": canonical_profile_url(row["username"]),
    }


def outcome_for(creator) -> str:
    if creator["validity"] == "invalid":
        return "invalid"
    if creator["client_status"] == "rejected":
        return "rejected"
    if creator["client_status"] == "accepted":
        return "accepted"
    if creator["validity"] == "valid":
        return "pending_client"
    return "unchecked"


def find_creators(conn, *names):
    keys = [name for name in dict.fromkeys(names) if name]
    found = {}
    for key in keys:
        rows = conn.execute(
            """
            SELECT * FROM creators
            WHERE username = ? OR (url_username IS NOT NULL AND url_username = ?)
            """,
            (key, key),
        ).fetchall()
        for row in rows:
            found[row["id"]] = row
    return list(found.values())


def header_key(label: str) -> str:
    return re.sub(r"\s+", " ", label.strip().lower())


def official_header_map(headers):
    expected = {header_key(label): field for label, field in FILE_COLUMNS}
    mapping = {}
    unknown = []
    for index, header in enumerate(headers):
        key = header_key(header)
        if not key:
            continue
        field = expected.get(key)
        if field:
            if field in mapping.values():
                raise improper_file(f"The column “{header}” is repeated.")
            mapping[field] = index
        elif key in IGNORED_FILE_HEADERS:
            continue
        else:
            unknown.append(str(header))
    if unknown:
        names = ", ".join(unknown)
        raise improper_file(f"It has extra columns: {names}.")
    if "username" not in mapping and "url" not in mapping:
        raise improper_file("The first row must be the header from the template.")
    return mapping


def read_table(filename: str, data: bytes):
    lower = filename.lower()
    if lower.endswith(".csv"):
        text = None
        for encoding in ("utf-8-sig", "cp1252"):
            try:
                text = data.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            text = data.decode("utf-8", errors="replace")
        reader = csv.reader(io.StringIO(text))
        grid = [[cell_str(cell) for cell in row] for row in reader]
    elif lower.endswith(".xlsx"):
        workbook = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        if "Creators" in workbook.sheetnames:
            sheet = workbook["Creators"]
        elif "Main" in workbook.sheetnames:
            sheet = workbook["Main"]
        else:
            sheet = workbook.active
        grid = []
        for row in sheet.iter_rows(values_only=True):
            grid.append([cell_str(value) for value in row])
    else:
        raise improper_file("Save it as an Excel file (.xlsx) from the template.")

    while grid and not any(grid[-1]):
        grid.pop()
    if not grid:
        raise improper_file("The file is empty.")
    mapping = official_header_map(grid[0])
    body = grid[1:]
    if len(body) > 20000:
        raise improper_file("It has more than 20,000 rows.")

    records = []
    for offset, cells in enumerate(body, start=2):
        if not any(cells):
            continue
        record = record_from_cells(cells, mapping, offset)
        if not cell_str(record["handle"]) and not cell_str(record["url"]):
            raise improper_file(
                f"Row {offset} needs an Instagram Handle or an Instagram Profile URL."
            )
        records.append(record)
    return records, []


def checked_choice(raw, parser, row_number, column_name, allowed):
    text = cell_str(raw)
    if not text:
        return ""
    parsed = parser(text)
    if not parsed:
        words = list(allowed)
        choices = words[0] if len(words) == 1 else ", ".join(words[:-1]) + ", or " + words[-1]
        raise improper_file(f"Row {row_number}: {column_name} must be {choices}.")
    return parsed


def parse_found_date(raw, row_number=0):
    text = cell_str(raw).replace("/", "-").replace(".", "-")
    if not text:
        return ""
    text = text.split(" ", 1)[0]
    parsed = None
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d-%m-%y"):
        try:
            parsed = datetime.strptime(text, fmt).date()
            break
        except ValueError:
            continue
    if parsed is None:
        if row_number:
            raise improper_file(f"Row {row_number}: Date must look like 17-11-2026.")
        raise ValueError("Date must look like 17-11-2026.")
    return parsed.isoformat()


def clean_agent_name(raw, row_number=0):
    text = cell_str(raw)
    if len(text) > 80:
        if row_number:
            raise improper_file(f"Row {row_number}: Agent name is too long.")
        text = text[:80]
    return text


def record_from_cells(cells, mapping, row_number=0):
    if mapping:
        def take(field):
            index = mapping.get(field)
            if index is None or index >= len(cells):
                return ""
            return cells[index]

        handle = take("username")
        url = take("url")
        country = take("country")
        validity = checked_choice(
            take("validity"),
            parse_validity,
            row_number,
            "Validity",
            ("Not checked yet", "Valid", "Invalid"),
        )
        client = checked_choice(
            take("client"),
            parse_client,
            row_number,
            "Client decision",
            ("Not decided", "Accepted", "Rejected"),
        )
        notes = take("notes")
        found_on = parse_found_date(take("found_on"), row_number)
        agent_name = clean_agent_name(take("agent_name"), row_number)
    else:
        handle = url = country = validity = client = notes = found_on = agent_name = ""
        extras = []
        for cell in cells:
            if not cell:
                continue
            lower = cell.lower()
            if not url and (
                "instagram.com" in lower
                or lower.startswith("http://")
                or lower.startswith("https://")
                or lower.startswith("www.")
            ):
                url = cell
            elif not validity and parse_validity(cell):
                validity = cell
            elif not client and parse_client(cell):
                client = cell
            else:
                extras.append(cell)
        if extras:
            handle = extras[0]
        if len(extras) > 1:
            country = extras[1]
    return {
        "handle": handle,
        "url": url,
        "country": country,
        "validity": parse_validity(validity) or "",
        "client": parse_client(client) or "",
        "notes": notes,
        "found_on": found_on,
        "agent_name": agent_name,
    }


def parse_freeform(text: str):
    records = []
    for line in text.splitlines():
        raw = line.strip()
        if not raw:
            continue
        if "\t" in raw:
            cells = [part.strip() for part in raw.split("\t")]
        elif "," in raw and "http" not in raw.split(",")[0].lower():
            cells = [part.strip() for part in raw.split(",")]
        else:
            cells = [raw]
        record = record_from_cells(cells, None)
        record["input"] = raw
        record["row_number"] = len(records) + 1
        records.append(record)
    return records


def names_for(record):
    handle_username = identity_from_text(record.get("handle"))
    url_username, url_kind = inspect_url(record.get("url"))
    if handle_username is None and record.get("handle"):
        extracted, kind = inspect_url(record.get("handle"))
        if kind == "ok":
            handle_username = extracted
    primary = handle_username or url_username
    link_status = classify_link(handle_username or primary, record.get("url"), url_username, url_kind)
    display = cell_str(record.get("handle"))
    if not display or "instagram.com" in display.lower() or display.lower().startswith("http"):
        display = primary or ""
    display = display.replace("*", "").strip().lstrip("@")
    return {
        "primary": primary,
        "handle_username": handle_username,
        "url_username": url_username,
        "url_kind": url_kind,
        "link_status": link_status,
        "display": display or (primary or ""),
    }


def plan_upload(conn, records, source_name, commit, seen_on=None):
    seen_on = seen_on or today_stamp()
    stamp = now_stamp()
    pending = {}
    results = []

    for record in records:
        parsed = names_for(record)
        primary = parsed["primary"]
        message = ""
        action = "skip"
        changes = []
        target_id = None

        if not primary:
            if parsed["url_kind"] == "non_instagram":
                action = "skip"
                message = "Skipped — the link is not an Instagram profile"
            else:
                action = "skip"
                message = "Skipped — no Instagram username"
        else:
            keys = [parsed["handle_username"], parsed["url_username"], primary]
            db_hits = find_creators(conn, *keys)
            pending_hits = []
            for key in {item for item in keys if item}:
                for username, planned in pending.items():
                    names = {username, planned.get("url_username")}
                    if key in names and planned not in pending_hits:
                        pending_hits.append(planned)
            if len(db_hits) > 1:
                action = "ambiguous"
                message = "Skipped — this row matches more than one saved creator"
            elif len(db_hits) == 1:
                current = db_hits[0]
                target_id = current["id"]
                changes = field_changes(current, record, parsed)
                if changes:
                    action = "update"
                    message = "Update " + ", ".join(changes)
                else:
                    action = "unchanged"
                    message = "Already on the list — no change"
            elif pending_hits:
                action = "duplicate"
                message = "Repeated in this file"
                first = pending_hits[0]
                merge_pending(first, record, parsed)
            else:
                action = "add"
                if parsed["link_status"] == "mismatch":
                    message = "Add to the master list, and flag that the handle and link differ"
                else:
                    message = "Add to the master list"

        result = {
            "row_number": record.get("row_number"),
            "input": record.get("input") or parsed["display"],
            "handle": parsed["display"],
            "username": primary or "",
            "url_username": parsed["url_username"] or "",
            "profile_url": cell_str(record.get("url")),
            "country": cell_str(record.get("country")),
            "validity": record.get("validity") or "",
            "client": record.get("client") or "",
            "notes": record.get("notes") or "",
            "found_on": record.get("found_on") or "",
            "agent_name": record.get("agent_name") or "",
            "link_status": parsed["link_status"],
            "action": action,
            "message": message,
            "changes": changes,
            "target_id": target_id,
        }
        results.append(result)
        if action == "add" and primary:
            pending[primary] = result

    if commit:
        for result in results:
            if result["action"] == "add":
                insert_creator(conn, result, source_name, result.get("found_on") or seen_on, stamp)
            elif result["action"] == "update":
                update_creator(conn, result, source_name, stamp)
        conn.commit()

    summary = {
        "add": sum(1 for row in results if row["action"] == "add"),
        "update": sum(1 for row in results if row["action"] == "update"),
        "unchanged": sum(1 for row in results if row["action"] == "unchanged"),
        "duplicate": sum(1 for row in results if row["action"] == "duplicate"),
        "skip": sum(1 for row in results if row["action"] == "skip"),
        "ambiguous": sum(1 for row in results if row["action"] == "ambiguous"),
    }
    return {"summary": summary, "rows": results}


def field_changes(current, record, parsed):
    changes = []
    new_url = cell_str(record.get("url"))
    if new_url and new_url != (current["profile_url"] or ""):
        changes.append("profile link")
    new_country = cell_str(record.get("country"))
    if new_country and new_country != (current["country_raw"] or ""):
        changes.append("country")
    if record.get("validity") and record["validity"] != current["validity"]:
        changes.append("validity")
    if record.get("client") and record["client"] != current["client_status"]:
        changes.append("client decision")
    if record.get("notes") and record["notes"] != (current["notes"] or ""):
        changes.append("notes")
    if parsed["link_status"] == "mismatch" and current["link_status"] == "ok":
        if "profile link" not in changes:
            changes.append("link check")
    return changes


def merge_pending(first, record, parsed):
    if record.get("validity") and not first["validity"]:
        first["validity"] = record["validity"]
    if record.get("client") and not first["client"]:
        first["client"] = record["client"]
    if record.get("notes") and not first["notes"]:
        first["notes"] = record["notes"]
    if cell_str(record.get("country")) and not first["country"]:
        first["country"] = cell_str(record.get("country"))
    if cell_str(record.get("url")) and not first["profile_url"]:
        first["profile_url"] = cell_str(record.get("url"))
        first["url_username"] = parsed["url_username"] or ""
        first["link_status"] = parsed["link_status"]


def insert_creator(conn, result, source_name, seen_on, stamp):
    username = result["username"]
    url_username = result["url_username"] or None
    if result["link_status"] == "ok" and url_username is None:
        url_username = username
    country = result["country"]
    conn.execute(
        """
        INSERT INTO creators (
            username, handle_display, profile_url, url_username, country_raw, country_group,
            validity, client_status, notes, link_status, first_seen, last_updated, source_file, agent_name
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            username,
            result["handle"] or username,
            result["profile_url"],
            url_username,
            country,
            country_group(country),
            result["validity"] or "unchecked",
            result["client"] or "unset",
            result["notes"] or "",
            result["link_status"] or "ok",
            seen_on,
            stamp,
            source_name,
            result.get("agent_name") or "",
        ),
    )


def update_creator(conn, result, source_name, stamp):
    current = conn.execute("SELECT * FROM creators WHERE id = ?", (result["target_id"],)).fetchone()
    if current is None:
        return
    profile_url = current["profile_url"]
    url_username = current["url_username"]
    link_status = current["link_status"]
    country_raw = current["country_raw"]
    validity = current["validity"]
    client_status = current["client_status"]
    notes = current["notes"] or ""
    handle_display = current["handle_display"]
    agent_name = (current["agent_name"] or "") if "agent_name" in current.keys() else ""

    if result["profile_url"]:
        profile_url = result["profile_url"]
        if result["url_username"]:
            url_username = result["url_username"]
        new_status = result["link_status"]
        if link_status == "confirmed" and new_status == "mismatch":
            new_status = "confirmed"
        link_status = new_status
    if result["country"]:
        country_raw = result["country"]
    if result["validity"]:
        validity = result["validity"]
    if result["client"]:
        client_status = result["client"]
    if result["notes"]:
        notes = result["notes"]
    if result["handle"] and "instagram.com" not in result["handle"].lower():
        handle_display = result["handle"]
    if result.get("agent_name") and not agent_name:
        agent_name = result["agent_name"]

    conn.execute(
        """
        UPDATE creators
        SET handle_display = ?, profile_url = ?, url_username = ?, country_raw = ?, country_group = ?,
            validity = ?, client_status = ?, notes = ?, link_status = ?, last_updated = ?, source_file = ?,
            agent_name = ?
        WHERE id = ?
        """,
        (
            handle_display,
            profile_url,
            url_username,
            country_raw,
            country_group(country_raw),
            validity,
            client_status,
            notes,
            link_status,
            stamp,
            source_name or current["source_file"],
            agent_name,
            current["id"],
        ),
    )


def check_records(conn, records):
    results = []
    seen = {}
    for record in records:
        parsed = names_for(record)
        primary = parsed["primary"]
        original = record.get("input") or record.get("handle") or record.get("url") or ""
        if not primary:
            if parsed["url_kind"] == "post_link":
                outcome, note = "not_instagram", "This link is a post, not a profile."
            elif parsed["url_kind"] == "non_instagram":
                outcome, note = "not_instagram", ""
            else:
                outcome, note = "empty", ""
            results.append(result_row(original, parsed, outcome, None, note, record))
            continue
        if primary in seen:
            results.append(
                result_row(
                    original,
                    parsed,
                    "duplicate_in_batch",
                    seen[primary],
                    "This username is already in the list you just pasted.",
                    record,
                )
            )
            continue
        hits = find_creators(conn, parsed["handle_username"], parsed["url_username"], primary)
        if len(hits) > 1:
            names = ", ".join(row["username"] for row in hits)
            results.append(
                result_row(original, parsed, "ambiguous", None, "Saved creators: " + names, record)
            )
            continue
        if len(hits) == 0:
            results.append(result_row(original, parsed, "new", None, "", record))
            seen[primary] = None
            continue
        creator = creator_public(hits[0])
        seen[primary] = creator
        note = match_note(parsed, creator)
        results.append(result_row(original, parsed, outcome_for(creator), creator, note, record))
    summary = {}
    for row in results:
        summary[row["outcome"]] = summary.get(row["outcome"], 0) + 1
    return {"summary": summary, "rows": results}


def match_note(parsed, creator):
    handle_name = parsed["handle_username"]
    link_name = parsed["url_username"]
    saved_link = creator["url_username"]
    if handle_name and handle_name != creator["username"] and handle_name == saved_link:
        return f"Matched on the name inside the profile link. Saved handle is {creator['handle']}."
    if link_name and saved_link and link_name == saved_link and saved_link != creator["username"]:
        return f"The link name {saved_link} is saved under the handle {creator['handle']}."
    if parsed["primary"] != creator["username"]:
        return f"Matched to saved handle {creator['handle']}."
    if creator["link_status"] == "mismatch":
        return f"Saved handle is {creator['handle']}. The stored profile link says {saved_link}."
    return ""


def result_row(original, parsed, outcome, creator, note, record=None):
    record = record or {}
    country = cell_str(record.get("country"))
    if not country and creator:
        country = creator["country"]
    profile_url = cell_str(record.get("url"))
    if not profile_url and creator:
        profile_url = creator["profile_url"]
    return {
        "input": original,
        "username": (creator or {}).get("username") or parsed["primary"] or "",
        "handle": (creator or {}).get("handle") or parsed["display"],
        "url_username": parsed["url_username"] or "",
        "profile_url": profile_url,
        "country": country,
        "link_status": (creator or {}).get("link_status") or parsed["link_status"],
        "found_on": record.get("found_on") or ((creator or {}).get("first_seen") or "")[:10],
        "agent_name": record.get("agent_name") or (creator or {}).get("agent_name") or "",
        "outcome": outcome,
        "label": OUTCOME_LABEL[outcome],
        "note": note,
        "creator": creator,
        "open_url": canonical_profile_url((creator or {}).get("username") or parsed["primary"])
        if (parsed["primary"] or (creator or {}).get("username"))
        else "",
    }


def export_values(row):
    keys = row.keys()
    client_key = "client_status" if "client_status" in keys else "client"
    return [
        row["handle_display"] if "handle_display" in keys else row["handle"],
        row["profile_url"] or "",
        row["url_username"] or "",
        row["country_raw"] if "country_raw" in keys else row["country"],
        VALIDITY_LABEL.get(row["validity"], row["validity"] or ""),
        CLIENT_LABEL.get(row[client_key], ""),
        row["notes"] or "",
        LINK_LABEL.get(row["link_status"], row["link_status"] or ""),
        (row["first_seen"] or "")[:10],
        (row["agent_name"] if "agent_name" in row.keys() else "") or "",
        (row["last_updated"] or "")[:10],
        row["source_file"] or "",
    ]


def style_header(sheet, headers):
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1B3A34")
    for column, header in enumerate(headers, start=1):
        cell = sheet.cell(1, column, header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center")
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"


def add_choice_validation(sheet, column, choices):
    check = DataValidation(
        type="list",
        formula1='"' + ",".join(choices) + '"',
        allow_blank=True,
    )
    check.error = "Choose one of: " + ", ".join(choices)
    check.errorStyle = "stop"
    check.add(f"{column}2:{column}5000")
    sheet.add_data_validation(check)


def template_workbook():
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Creators"
    headers = [label for label, _field in FILE_COLUMNS]
    style_header(sheet, headers)
    widths = [28, 46, 22, 20, 20, 36, 16, 24]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    add_choice_validation(sheet, "D", ["Not checked yet", "Valid", "Invalid"])
    add_choice_validation(sheet, "E", ["Not decided", "Accepted", "Rejected"])
    guide = workbook.create_sheet("How to fill")
    guide["A1"] = "Use the Creators sheet. Keep this header row and do not add columns."
    guide["A2"] = "Instagram Handle: the username, without @."
    guide["A3"] = "Instagram Profile URL: the profile link, for example https://www.instagram.com/name/"
    guide["A4"] = "Country: the country you already use, such as USA or United Kingdom."
    guide["A5"] = "Validity: Not checked yet, Valid, or Invalid. Leave it blank when it is not checked yet."
    guide["A6"] = "Client decision: Not decided, Accepted, or Rejected. Leave it blank when the client has not decided."
    guide["A7"] = "Notes: optional."
    guide["A8"] = "Date: the day this creator was found, like 17-11-2026. Leave it blank to use today."
    guide["A9"] = "Agent name: the person who found this creator."
    guide["A10"] = "One creator on each row. A new username is saved on the master list as soon as it is checked."
    guide.column_dimensions["A"].width = 110
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer


def check_workbook(rows):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Creators"
    headers = [label for label, _field in FILE_COLUMNS] + ["Result"]
    style_header(sheet, headers)
    for index, row in enumerate(rows, start=2):
        values = [
            row.get("handle") or "",
            row.get("url") or "",
            row.get("country") or "",
            VALIDITY_LABEL.get(row.get("validity") or "", row.get("validity") or ""),
            CLIENT_LABEL.get(row.get("client") or "", row.get("client") or ""),
            row.get("notes") or "",
            row.get("found_on") or "",
            row.get("agent_name") or "",
            row.get("result") or "",
        ]
        for column, value in enumerate(values, start=1):
            sheet.cell(index, column, value)
    widths = [28, 46, 22, 20, 20, 36, 16, 24, 42]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    add_choice_validation(sheet, "D", ["Not checked yet", "Valid", "Invalid"])
    add_choice_validation(sheet, "E", ["Not decided", "Accepted", "Rejected"])
    last_row = max(sheet.max_row, 1)
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{last_row}"
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer


def workbook_for(rows, styled=True):
    if not styled:
        workbook = openpyxl.Workbook(write_only=True)
        sheet = workbook.create_sheet("Creators")
        sheet.append(EXPORT_HEADERS)
        for row in rows:
            sheet.append(export_values(row))
        buffer = io.BytesIO()
        workbook.save(buffer)
        buffer.seek(0)
        return buffer
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Creators"
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1B3A34")
    for column, header in enumerate(EXPORT_HEADERS, start=1):
        cell = sheet.cell(1, column, header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center")
    for index, row in enumerate(rows, start=2):
        for column, value in enumerate(export_values(row), start=1):
            sheet.cell(index, column, value)
    widths = [28, 46, 26, 22, 18, 18, 32, 28, 14, 14, 32]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.auto_filter.ref = sheet.dimensions
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(EXPORT_HEADERS))}{max(sheet.max_row, 1)}"
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer


def load_creator(conn, creator_id):
    row = conn.execute("SELECT * FROM creators WHERE id = ?", (creator_id,)).fetchone()
    return row


@app.get("/")
def home():
    return render_template("index.html")


@app.get("/api/template.xlsx")
def download_template():
    return send_file(
        template_workbook(),
        as_attachment=True,
        download_name="creator-template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/api/export-check")
def export_check():
    payload = request.get_json(silent=True) or {}
    rows = payload.get("rows") or []
    if not rows:
        return jsonify(error="There is no result to download."), 400
    return send_file(
        check_workbook(rows),
        as_attachment=True,
        download_name=f"creator-check-{today_stamp()}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def filters_from_mapping(source):
    def pick(*keys):
        for key in keys:
            value = source.get(key)
            if value:
                return str(value).strip()
        return ""

    validity = pick("validity")
    client = pick("client", "client_status")
    link = pick("link", "link_status")
    return {
        "q": pick("q"),
        "country": pick("country"),
        "validity": validity if validity in {"unchecked", "valid", "invalid"} else "",
        "client": client if client in {"unset", "accepted", "rejected"} else "",
        "link": link if link in LINK_FILTERS else "",
    }


def creator_where(filters):
    clauses = []
    params = []
    if filters["validity"]:
        clauses.append("validity = ?")
        params.append(filters["validity"])
    if filters["client"]:
        clauses.append("client_status = ?")
        params.append(filters["client"])
    if filters["link"]:
        clauses.append("link_status = ?")
        params.append(filters["link"])
    if filters["country"]:
        clauses.append("country_group = ?")
        params.append(filters["country"])
    if filters["q"]:
        like = f"%{filters['q']}%"
        clauses.append(
            """(
                username LIKE ? OR handle_display LIKE ? OR IFNULL(url_username, '') LIKE ?
                OR IFNULL(country_raw, '') LIKE ? OR country_group LIKE ?
                OR IFNULL(notes, '') LIKE ? OR IFNULL(profile_url, '') LIKE ?
            )"""
        )
        params.extend([like] * 7)
    return (" AND ".join(clauses) if clauses else "1 = 1"), params


def database_summary(conn):
    row = conn.execute(
        """
        SELECT
            COUNT(*) AS total,
            SUM(CASE WHEN validity = 'unchecked' THEN 1 ELSE 0 END) AS unchecked,
            SUM(CASE WHEN validity = 'valid' THEN 1 ELSE 0 END) AS valid,
            SUM(CASE WHEN validity = 'invalid' THEN 1 ELSE 0 END) AS invalid,
            SUM(CASE WHEN client_status = 'accepted' THEN 1 ELSE 0 END) AS accepted,
            SUM(CASE WHEN client_status = 'rejected' THEN 1 ELSE 0 END) AS rejected,
            SUM(CASE WHEN link_status = 'mismatch' THEN 1 ELSE 0 END) AS mismatch
        FROM creators
        """
    ).fetchone()
    countries = []
    for item in conn.execute(
        "SELECT country_group, COUNT(*) AS n FROM creators GROUP BY country_group"
    ):
        countries.append({"name": item["country_group"] or "No country", "count": item["n"]})
    countries.sort(key=lambda item: (item["name"] == "No country", item["name"].lower()))
    return {
        "total": row["total"] or 0,
        "unchecked": row["unchecked"] or 0,
        "valid": row["valid"] or 0,
        "invalid": row["invalid"] or 0,
        "accepted": row["accepted"] or 0,
        "rejected": row["rejected"] or 0,
        "mismatch": row["mismatch"] or 0,
        "countries": countries,
    }


@app.get("/api/creators")
def list_creators():
    filters = filters_from_mapping(request.args)
    sort = SORT_COLUMNS.get(request.args.get("sort") or "first_seen", "first_seen")
    direction = "DESC" if (request.args.get("dir") or "").lower() == "desc" else "ASC"
    try:
        page = max(1, int(request.args.get("page") or 1))
    except ValueError:
        page = 1
    where, params = creator_where(filters)
    with db() as conn:
        summary = database_summary(conn)
        matched = conn.execute(
            f"SELECT COUNT(*) AS n FROM creators WHERE {where}", params
        ).fetchone()["n"]
        pages = max(1, (matched + PAGE_SIZE - 1) // PAGE_SIZE) if matched else 1
        page = min(page, pages)
        rows = conn.execute(
            f"""
            SELECT * FROM creators
            WHERE {where}
            ORDER BY {sort} {direction}, username ASC
            LIMIT ? OFFSET ?
            """,
            [*params, PAGE_SIZE, (page - 1) * PAGE_SIZE],
        ).fetchall()
    stats = {key: summary[key] for key in ("total", "unchecked", "valid", "invalid", "accepted", "rejected", "mismatch")}
    return jsonify(
        creators=[creator_public(row) for row in rows],
        total=summary["total"],
        matched=matched,
        page=page,
        pages=pages,
        page_size=PAGE_SIZE,
        stats=stats,
        countries=summary["countries"],
    )


@app.patch("/api/creators/<int:creator_id>")
def patch_creator(creator_id):
    payload = request.get_json(silent=True) or {}
    with db(write=True) as conn:
        current = load_creator(conn, creator_id)
        if current is None:
            return jsonify(error="That creator is no longer on the list."), 404
        validity = current["validity"]
        client_status = current["client_status"]
        notes = current["notes"] or ""
        country_raw = current["country_raw"] or ""
        profile_url = current["profile_url"] or ""
        url_username = current["url_username"]
        link_status = current["link_status"]
        if "validity" in payload:
            parsed = parse_validity(payload.get("validity"))
            if not parsed:
                return jsonify(error="Validity must be Not checked yet, Valid, or Invalid."), 400
            validity = parsed
        if "client_status" in payload:
            parsed = parse_client(payload.get("client_status"))
            if not parsed:
                return jsonify(error="Client decision must be Not decided, Accepted, or Rejected."), 400
            client_status = parsed
        if "notes" in payload:
            notes = cell_str(payload.get("notes"))
        if "country" in payload:
            country_raw = cell_str(payload.get("country"))
        if "profile_url" in payload:
            profile_url = cell_str(payload.get("profile_url"))
            url_username, url_kind = inspect_url(profile_url)
            link_status = classify_link(current["username"], profile_url, url_username, url_kind)
            if url_kind == "ok" and link_status == "ok" and not url_username:
                url_username = current["username"]
        conn.execute(
            """
            UPDATE creators
            SET validity = ?, client_status = ?, notes = ?, country_raw = ?, country_group = ?,
                profile_url = ?, url_username = ?, link_status = ?, last_updated = ?
            WHERE id = ?
            """,
            (
                validity,
                client_status,
                notes,
                country_raw,
                country_group(country_raw),
                profile_url,
                url_username,
                link_status,
                now_stamp(),
                creator_id,
            ),
        )
        updated = load_creator(conn, creator_id)
    return jsonify(creator=creator_public(updated))


@app.post("/api/creators/<int:creator_id>/use-link-name")
def use_link_name(creator_id):
    with db(write=True) as conn:
        current = load_creator(conn, creator_id)
        if current is None:
            return jsonify(error="That creator is no longer on the list."), 404
        new_name = current["url_username"] or ""
        if not new_name or new_name == current["username"]:
            return jsonify(error="The profile link does not contain a different username."), 400
        taken = conn.execute(
            "SELECT id, handle_display FROM creators WHERE username = ? AND id != ?",
            (new_name, creator_id),
        ).fetchone()
        if taken:
            return jsonify(
                error=f"@{new_name} is already saved as its own row. Decide which row to keep."
            ), 409
        conn.execute(
            """
            UPDATE creators
            SET username = ?, handle_display = ?, url_username = ?, link_status = 'ok', last_updated = ?
            WHERE id = ?
            """,
            (new_name, new_name, new_name, now_stamp(), creator_id),
        )
        updated = load_creator(conn, creator_id)
    return jsonify(creator=creator_public(updated))


@app.post("/api/creators/<int:creator_id>/keep-handle")
def keep_handle(creator_id):
    with db(write=True) as conn:
        current = load_creator(conn, creator_id)
        if current is None:
            return jsonify(error="That creator is no longer on the list."), 404
        conn.execute(
            "UPDATE creators SET link_status = 'confirmed', last_updated = ? WHERE id = ?",
            (now_stamp(), creator_id),
        )
        updated = load_creator(conn, creator_id)
    return jsonify(creator=creator_public(updated))


@app.delete("/api/creators/<int:creator_id>")
def delete_creator(creator_id):
    with db(write=True) as conn:
        current = load_creator(conn, creator_id)
        if current is None:
            return jsonify(error="That creator is no longer on the list."), 404
        conn.execute("DELETE FROM creators WHERE id = ?", (creator_id,))
    return jsonify(ok=True)


def one_lookup_record(payload):
    handle = cell_str(payload.get("handle"))
    url = cell_str(payload.get("url"))
    if not handle and not url:
        raise ValueError("Enter an Instagram handle or a profile link.")
    return {
        "handle": handle,
        "url": url,
        "country": "",
        "input": handle or url,
        "row_number": 1,
    }


@app.post("/api/agent/match")
def agent_match():
    payload = request.get_json(silent=True) or {}
    try:
        record = one_lookup_record(payload)
    except ValueError as error:
        return jsonify(error=str(error)), 400
    with db() as conn:
        row = check_records(conn, [record])["rows"][0]
    return jsonify(row=row)


@app.post("/api/agent/update")
def agent_update():
    payload = request.get_json(silent=True) or {}
    try:
        record = one_lookup_record(payload)
    except ValueError as error:
        return jsonify(error=str(error)), 400
    validity = parse_validity(payload.get("validity"))
    client_status = parse_client(payload.get("client_status"))
    if not validity or not client_status:
        return jsonify(error="Choose Valid or Invalid, and whether the client approved."), 400
    with db(write=True) as conn:
        row = check_records(conn, [record])["rows"][0]
        if row["outcome"] == "ambiguous":
            return jsonify(error="This matches more than one creator. An admin needs to sort that out."), 409
        creator = row["creator"]
        if not creator:
            return jsonify(error="This creator is not on the master list, so there is nothing to update."), 404
        conn.execute(
            "UPDATE creators SET validity = ?, client_status = ?, last_updated = ? WHERE id = ?",
            (validity, client_status, now_stamp(), creator["id"]),
        )
        updated = load_creator(conn, creator["id"])
    return jsonify(creator=creator_public(updated))


def remember_new_creators(conn, report, source_name, default_agent):
    today = today_stamp()
    saved = 0
    for row in report["rows"]:
        if row["outcome"] != "new" or not row.get("username"):
            continue
        result = {
            "username": row["username"],
            "handle": row.get("handle") or row["username"],
            "url_username": row.get("url_username") or "",
            "profile_url": row.get("profile_url") or "",
            "country": row.get("country") or "",
            "validity": "",
            "client": "",
            "notes": "",
            "link_status": row.get("link_status") or "ok",
            "agent_name": row.get("agent_name") or default_agent,
        }
        try:
            insert_creator(conn, result, source_name, row.get("found_on") or today, now_stamp())
        except Exception as error:
            message = str(error).lower()
            if "unique" in message or "duplicate" in message:
                row["outcome"] = "unchecked"
                row["label"] = OUTCOME_LABEL["unchecked"]
                continue
            raise
        row["label"] = "New — saved to the master list"
        row["agent_name"] = result["agent_name"]
        row["found_on"] = row.get("found_on") or today
        row["saved"] = True
        saved += 1
    summary = {}
    for row in report["rows"]:
        summary[row["outcome"]] = summary.get(row["outcome"], 0) + 1
    report["summary"] = summary
    report["saved"] = saved


def settle_record(conn, record, source_name):
    parsed = names_for(record)
    original = cell_str(record.get("input") or record.get("handle") or record.get("url"))
    country = cell_str(record.get("country"))
    if parsed["handle_username"] and parsed["url_username"] and parsed["handle_username"] != parsed["url_username"]:
        return {
            "input": original,
            "status": "mismatch",
            "message": "Handle and link are different names",
            "handle": parsed["display"] or "",
            "profile_url": cell_str(record.get("url")),
            "country": country,
        }
    if not parsed["primary"]:
        if parsed["url_kind"] in {"non_instagram", "post_link"}:
            message = "Not an Instagram profile link"
        else:
            message = "No username on this line"
        return {
            "input": original,
            "status": "invalid",
            "message": message,
            "handle": "",
            "profile_url": "",
            "country": country,
        }
    profile_url = cell_str(record.get("url"))
    if parsed["url_kind"] != "ok":
        profile_url = canonical_profile_url(parsed["primary"])
    hits = find_creators(conn, parsed["handle_username"], parsed["url_username"], parsed["primary"])
    if hits:
        creator = creator_public(hits[0])
        return {
            "input": original,
            "status": "exists",
            "message": "Already exists",
            "handle": creator["handle"],
            "profile_url": creator["profile_url"] or creator["open_url"],
            "country": creator["country"] or country,
        }
    result = {
        "username": parsed["primary"],
        "handle": parsed["display"] or parsed["primary"],
        "url_username": parsed["url_username"] or parsed["primary"],
        "profile_url": profile_url,
        "country": country,
        "validity": "",
        "client": "",
        "notes": "",
        "link_status": "ok",
        "agent_name": "",
    }
    try:
        insert_creator(conn, result, source_name, today_stamp(), now_stamp())
    except Exception as error:
        message = str(error).lower()
        if "unique" not in message and "duplicate" not in message:
            raise
        return {
            "input": original,
            "status": "exists",
            "message": "Already exists",
            "handle": result["handle"],
            "profile_url": profile_url,
            "country": country,
        }
    return {
        "input": original,
        "status": "added",
        "message": "Added to the master list",
        "handle": result["handle"],
        "profile_url": profile_url,
        "country": country,
    }


@app.post("/api/check-one")
def check_one():
    payload = request.get_json(silent=True) or {}
    record = {"handle": payload.get("handle") or "", "url": payload.get("url") or "", "input": ""}
    with db(write=True) as conn:
        row = settle_record(conn, record, "Checked on the page")
    return jsonify(row)


@app.post("/api/check-list")
def check_list():
    warnings = []
    if request.files.get("file"):
        upload = request.files["file"]
        filename = secure_filename(upload.filename or "upload.xlsx") or "upload.xlsx"
        try:
            records, warnings = read_table(filename, upload.read())
        except ValueError as error:
            return jsonify(error=str(error)), 400
        except Exception:
            traceback.print_exc()
            return jsonify(error=str(improper_file("The file could not be opened."))), 400
        source_name = filename
    else:
        payload = request.get_json(silent=True) or {}
        records = parse_freeform(payload.get("text") or "")
        source_name = "Checked from a list"
    if not records:
        return jsonify(error="Paste a handle, an @name, or a profile link."), 400
    with db(write=True) as conn:
        rows = [settle_record(conn, record, source_name) for record in records]
    added = sum(row["status"] == "added" for row in rows)
    exists = sum(row["status"] == "exists" for row in rows)
    return jsonify(rows=rows, added=added, exists=exists, warnings=warnings)


@app.post("/api/check")
def check():
    if request.files.get("file"):
        upload = request.files["file"]
        filename = secure_filename(upload.filename or "upload.xlsx") or "upload.xlsx"
        default_agent = clean_agent_name(request.form.get("agent_name"))
        try:
            records, warnings = read_table(filename, upload.read())
        except ValueError as error:
            return jsonify(error=str(error)), 400
        except Exception:
            traceback.print_exc()
            return jsonify(error=str(improper_file("The file could not be opened."))), 400
        source_name = filename
    else:
        payload = request.get_json(silent=True) or {}
        default_agent = clean_agent_name(payload.get("agent_name"))
        records = parse_freeform(payload.get("text") or "")
        warnings = []
        source_name = "Checked from a list"
    if not default_agent and any(not record.get("agent_name") for record in records):
        return jsonify(error="Enter the agent name. It is saved with every new creator."), 400
    for record in records:
        record["input"] = record.get("input") or " ".join(
            part for part in (record.get("handle"), record.get("url")) if part
        )
    with db(write=True) as conn:
        report = check_records(conn, records)
        remember_new_creators(conn, report, source_name, default_agent)
    report["warnings"] = warnings
    return jsonify(report)


@app.post("/api/creators/bulk")
def bulk_add():
    payload = request.get_json(silent=True) or {}
    items = payload.get("items") or []
    added = []
    with db(write=True) as conn:
        for item in items:
            record = {
                "handle": item.get("handle") or item.get("username") or "",
                "url": item.get("profile_url") or "",
                "country": item.get("country") or "",
                "validity": "",
                "client": "",
                "notes": item.get("notes") or "",
                "row_number": 0,
            }
            parsed = names_for(record)
            if not parsed["primary"]:
                continue
            hits = find_creators(conn, parsed["handle_username"], parsed["url_username"], parsed["primary"])
            if hits:
                continue
            result = {
                "username": parsed["primary"],
                "handle": parsed["display"],
                "url_username": parsed["url_username"] or "",
                "profile_url": cell_str(record["url"]),
                "country": cell_str(record["country"]),
                "validity": "",
                "client": "",
                "notes": record["notes"],
                "link_status": parsed["link_status"],
            }
            insert_creator(conn, result, "Added from a check", today_stamp(), now_stamp())
            row = conn.execute(
                "SELECT * FROM creators WHERE username = ?", (parsed["primary"],)
            ).fetchone()
            added.append(creator_public(row))
    return jsonify(added=len(added), creators=added)


@app.post("/api/upload")
def upload():
    upload_file = request.files.get("file")
    if upload_file is None:
        return jsonify(error="Choose an .xlsx or .csv file first."), 400
    filename = secure_filename(upload_file.filename or "upload.xlsx") or "upload.xlsx"
    commit = request.form.get("commit") == "1"
    try:
        records, warnings = read_table(filename, upload_file.read())
    except ValueError as error:
        return jsonify(error=str(error)), 400
    except Exception:
        traceback.print_exc()
        return jsonify(error=str(improper_file("The file could not be opened."))), 400
    with db(write=commit) as conn:
        report = plan_upload(conn, records, filename, commit=commit)
    report["warnings"] = warnings
    report["filename"] = filename
    return jsonify(report)


@app.post("/api/export")
def export():
    payload = request.get_json(silent=True) or {}
    filters = filters_from_mapping(payload)
    filtered = any(filters.values())
    where, params = creator_where(filters)
    with db() as conn:
        matched = conn.execute(
            f"SELECT COUNT(*) AS n FROM creators WHERE {where}", params
        ).fetchone()["n"]
        if filtered and matched == 0:
            return jsonify(error="There are no rows to download."), 400
        rows = conn.execute(
            f"SELECT * FROM creators WHERE {where} ORDER BY first_seen ASC, username ASC",
            params,
        )
        workbook = workbook_for(rows, styled=matched <= 8000)
    name = f"creator-filtered-{today_stamp()}.xlsx" if filtered else f"creator-master-{today_stamp()}.xlsx"
    return send_file(
        workbook,
        as_attachment=True,
        download_name=name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.errorhandler(413)
def too_large(_error):
    return jsonify(error="That file is larger than 10 MB."), 413


init_db()


def main():
    parser = argparse.ArgumentParser(description="Creator duplicate checker")
    parser.add_argument("--open", action="store_true", help="Open the checker in the browser")
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()
    url = f"http://127.0.0.1:{args.port}"
    print(f"Creator duplicate checker is running at {url}")
    if args.open:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=args.port, debug=False, threaded=True)


if __name__ == "__main__":
    main()
