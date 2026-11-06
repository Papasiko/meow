from __future__ import annotations

import io
import os
import secrets
import sqlite3
from datetime import datetime, timedelta
from functools import wraps
from pathlib import Path
from typing import Any, Callable

from flask import Flask, g, jsonify, render_template, request, session, send_file
from werkzeug.security import check_password_hash, generate_password_hash
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

BUILD_VERSION = "2026.11.02-requirements-update-v1"
BASE_DIR = Path(__file__).resolve().parent
INSTANCE_DIR = BASE_DIR / "instance"
INSTANCE_DIR.mkdir(exist_ok=True)
DB_PATH = Path(os.environ.get("EVALUATION_DB_PATH", str(INSTANCE_DIR / "evaluation.db")))
SECRET_PATH = INSTANCE_DIR / ".secret_key"


def gregorian_to_jalali(gy: int, gm: int, gd: int) -> tuple[int, int, int]:
    """Convert a Gregorian date to Jalali without external dependencies."""
    g_day_month = (0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334)
    gy2 = gy + 1 if gm > 2 else gy
    days = (355666 + 365 * gy + (gy2 + 3) // 4 - (gy2 + 99) // 100
            + (gy2 + 399) // 400 + gd + g_day_month[gm - 1])
    jy = -1595 + 33 * (days // 12053)
    days %= 12053
    jy += 4 * (days // 1461)
    days %= 1461
    if days > 365:
        jy += (days - 1) // 365
        days = (days - 1) % 365
    if days < 186:
        jm, jd = 1 + days // 31, 1 + days % 31
    else:
        jm, jd = 7 + (days - 186) // 30, 1 + (days - 186) % 30
    return jy, jm, jd


def jalali_to_gregorian(jy: int, jm: int, jd: int) -> tuple[int, int, int]:
    """Convert a Jalali date to Gregorian without external dependencies."""
    jy += 1595
    days = -355668 + 365 * jy + (jy // 33) * 8 + ((jy % 33 + 3) // 4) + jd
    days += (jm - 1) * 31 if jm < 7 else (jm - 7) * 30 + 186
    gy = 400 * (days // 146097)
    days %= 146097
    if days > 36524:
        gy += 100 * ((days - 1) // 36524)
        days = (days - 1) % 36524
        if days >= 365:
            days += 1
    gy += 4 * (days // 1461)
    days %= 1461
    if days > 365:
        gy += (days - 1) // 365
        days = (days - 1) % 365
    gd = days + 1
    leap = gy % 4 == 0 and (gy % 100 != 0 or gy % 400 == 0)
    month_days = (31, 29 if leap else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    gm = 1
    while gm <= 12 and gd > month_days[gm - 1]:
        gd -= month_days[gm - 1]
        gm += 1
    return gy, gm, gd


def persian_digits(value: str) -> str:
    return str(value).translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))


def jalali_date_text(value: str | datetime | None, include_time: bool = False) -> str:
    """Format an ISO/Gregorian date or timestamp as a Persian-calendar string."""
    if not value:
        return "—"
    raw = value.strftime("%Y-%m-%d %H:%M:%S") if isinstance(value, datetime) else str(value).strip()
    try:
        date_part = raw[:10]
        gy, gm, gd = (int(part) for part in date_part.split("-"))
        jy, jm, jd = gregorian_to_jalali(gy, gm, gd)
        result = f"{jy:04d}/{jm:02d}/{jd:02d}"
        if include_time and len(raw) >= 16:
            time_part = raw[11:16]
            result += f" {time_part}"
        return persian_digits(result)
    except (ValueError, TypeError, IndexError):
        return str(value)


def get_jalali_month_days(jy: int, jm: int) -> int:
    if jm <= 6:
        return 31
    if jm <= 11:
        return 30
    start = jalali_to_gregorian(jy, 1, 1)
    next_start = jalali_to_gregorian(jy + 1, 1, 1)
    from datetime import date as _date
    return 30 if (_date(*next_start) - _date(*start)).days == 366 else 29


def get_or_create_secret() -> str:
    secret = os.environ.get("FLASK_SECRET_KEY")
    if secret:
        return secret
    if SECRET_PATH.exists():
        return SECRET_PATH.read_text(encoding="utf-8").strip()
    secret = secrets.token_urlsafe(48)
    SECRET_PATH.write_text(secret, encoding="utf-8")
    try:
        SECRET_PATH.chmod(0o600)
    except OSError:
        pass
    return secret


app = Flask(__name__, template_folder="templates")
app.secret_key = get_or_create_secret()
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("SESSION_COOKIE_SECURE", "0") == "1",
    PERMANENT_SESSION_LIFETIME=timedelta(hours=8),
    MAX_CONTENT_LENGTH=3 * 1024 * 1024,
)


def connect_db() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA busy_timeout = 15000")
    return conn


def table_exists(conn: sqlite3.Connection, table: str) -> bool:
    return conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone() is not None


def table_columns(conn: sqlite3.Connection, table: str) -> set[str]:
    if not table_exists(conn, table):
        return set()
    return {row["name"] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()}


def log_event(conn: sqlite3.Connection, actor_id: int | None, action: str, entity: str, entity_id: int | None, detail: str = "") -> None:
    conn.execute(
        "INSERT INTO audit_log (actor_id, action, entity, entity_id, detail, created_at) VALUES (?, ?, ?, ?, ?, datetime('now','localtime'))",
        (actor_id, action, entity, entity_id, detail[:500]),
    )


def rebuild_legacy_users(conn: sqlite3.Connection) -> None:
    cols = table_columns(conn, "users")
    if not cols or "max_points" in cols:
        return
    # Rebuild the first-version table so additional evaluator accounts are no longer limited to IDs 1..7.
    conn.execute("""
        CREATE TABLE users_new (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE COLLATE NOCASE,
            display_name TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL CHECK (role IN ('admin','evaluator')),
            evaluator_no INTEGER,
            max_points REAL NOT NULL DEFAULT 0 CHECK (max_points >= 0 AND max_points <= 100),
            section_name TEXT NOT NULL DEFAULT '',
            active INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0,1)),
            created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
        )
    """)
    source_cols = cols
    rows = conn.execute("SELECT * FROM users").fetchall()
    for row in rows:
        data = dict(row)
        n = data.get("evaluator_no")
        max_points = 0 if data.get("role") == "admin" else (10 if n == 1 else 15)
        section = "" if data.get("role") == "admin" else f"بخش ارزیابی {n or ''}".strip()
        conn.execute(
            "INSERT INTO users_new (id,username,display_name,password_hash,role,evaluator_no,max_points,section_name,active,created_at) VALUES (?,?,?,?,?,?,?,?,?,?)",
            (data.get("id"), data.get("username"), data.get("display_name"), data.get("password_hash"), data.get("role", "evaluator"), n, max_points, section, data.get("active", 1), data.get("created_at") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )
    conn.execute("DROP TABLE users")
    conn.execute("ALTER TABLE users_new RENAME TO users")


def rebuild_legacy_questions(conn: sqlite3.Connection) -> None:
    cols = table_columns(conn, "questions")
    if not cols:
        conn.execute("""
            CREATE TABLE questions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                question_text TEXT NOT NULL,
                rubric TEXT NOT NULL DEFAULT '',
                max_score REAL NOT NULL CHECK (max_score > 0),
                evaluator_no INTEGER,
                assigned_user_id INTEGER REFERENCES users(id) ON DELETE RESTRICT,
                section_name TEXT NOT NULL DEFAULT '',
                version INTEGER NOT NULL DEFAULT 1,
                active INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0,1)),
                created_at TEXT NOT NULL DEFAULT (datetime('now','localtime')),
                updated_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            )
        """)
        return
    if "assigned_user_id" in cols and "section_name" in cols:
        return
    rows = conn.execute("SELECT * FROM questions").fetchall()
    conn.execute("""
        CREATE TABLE questions_new (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            question_text TEXT NOT NULL,
            rubric TEXT NOT NULL DEFAULT '',
            max_score REAL NOT NULL CHECK (max_score > 0),
            evaluator_no INTEGER,
            assigned_user_id INTEGER REFERENCES users(id) ON DELETE RESTRICT,
            section_name TEXT NOT NULL DEFAULT '',
            version INTEGER NOT NULL DEFAULT 1,
            active INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0,1)),
            created_at TEXT NOT NULL DEFAULT (datetime('now','localtime')),
            updated_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
        )
    """)
    user_by_no = {row["evaluator_no"]: row["id"] for row in conn.execute("SELECT id,evaluator_no FROM users WHERE evaluator_no IS NOT NULL").fetchall()}
    section_by_id = {row["id"]: row["section_name"] for row in conn.execute("SELECT id,section_name FROM users").fetchall()}
    for row in rows:
        d = dict(row)
        assigned_id = user_by_no.get(d.get("evaluator_no"))
        conn.execute(
            "INSERT INTO questions_new (id,question_text,rubric,max_score,evaluator_no,assigned_user_id,section_name,version,active,created_at,updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (d.get("id"), d.get("question_text", ""), d.get("rubric", ""), d.get("max_score", 1), d.get("evaluator_no"), assigned_id, section_by_id.get(assigned_id, ""), d.get("version", 1), d.get("active", 1), d.get("created_at") or datetime.now().strftime("%Y-%m-%d %H:%M:%S"), d.get("updated_at") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )
    conn.execute("DROP TABLE questions")
    conn.execute("ALTER TABLE questions_new RENAME TO questions")


def rebuild_legacy_scores(conn: sqlite3.Connection) -> None:
    cols = table_columns(conn, "scores")
    if not cols:
        conn.execute("""
            CREATE TABLE scores (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                period_id INTEGER NOT NULL REFERENCES evaluation_periods(id),
                unit_id INTEGER NOT NULL REFERENCES units(id) ON DELETE RESTRICT,
                question_id INTEGER REFERENCES questions(id) ON DELETE SET NULL,
                evaluator_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
                evaluator_no INTEGER,
                question_text_snapshot TEXT NOT NULL,
                max_score_snapshot REAL NOT NULL CHECK (max_score_snapshot > 0),
                score REAL NOT NULL CHECK (score >= 0 AND score <= max_score_snapshot),
                comment TEXT NOT NULL DEFAULT '',
                updated_by INTEGER REFERENCES users(id),
                updated_at TEXT NOT NULL DEFAULT (datetime('now','localtime')),
                UNIQUE(period_id, unit_id, question_id)
            )
        """)
        return
    if "evaluator_user_id" in cols:
        return
    rows = conn.execute("SELECT * FROM scores").fetchall()
    conn.execute("""
        CREATE TABLE scores_new (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period_id INTEGER NOT NULL REFERENCES evaluation_periods(id),
            unit_id INTEGER NOT NULL REFERENCES units(id) ON DELETE RESTRICT,
            question_id INTEGER REFERENCES questions(id) ON DELETE SET NULL,
            evaluator_user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
            evaluator_no INTEGER,
            question_text_snapshot TEXT NOT NULL,
            max_score_snapshot REAL NOT NULL CHECK (max_score_snapshot > 0),
            score REAL NOT NULL CHECK (score >= 0 AND score <= max_score_snapshot),
            comment TEXT NOT NULL DEFAULT '',
            updated_by INTEGER REFERENCES users(id),
            updated_at TEXT NOT NULL DEFAULT (datetime('now','localtime')),
            UNIQUE(period_id, unit_id, question_id)
        )
    """)
    for row in rows:
        d = dict(row)
        qrow = conn.execute("SELECT assigned_user_id FROM questions WHERE id=?", (d.get("question_id"),)).fetchone() if d.get("question_id") else None
        evaluator_user_id = qrow[0] if qrow else None
        if evaluator_user_id is None and d.get("evaluator_no") is not None:
            urow = conn.execute("SELECT id FROM users WHERE evaluator_no=?", (d.get("evaluator_no"),)).fetchone()
            evaluator_user_id = urow[0] if urow else None
        conn.execute(
            "INSERT INTO scores_new (id,period_id,unit_id,question_id,evaluator_user_id,evaluator_no,question_text_snapshot,max_score_snapshot,score,comment,updated_by,updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
            (d.get("id"), d.get("period_id"), d.get("unit_id"), d.get("question_id"), evaluator_user_id, d.get("evaluator_no"), d.get("question_text_snapshot", "سؤال قدیمی"), d.get("max_score_snapshot", 1), d.get("score", 0), d.get("comment", ""), d.get("updated_by"), d.get("updated_at") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )
    conn.execute("DROP TABLE scores")
    conn.execute("ALTER TABLE scores_new RENAME TO scores")


def init_db() -> None:
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = OFF")
    try:
        # Create foundational tables before migrating references to them.
        conn.execute("""
            CREATE TABLE IF NOT EXISTS units (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT NOT NULL UNIQUE COLLATE NOCASE,
                name TEXT NOT NULL UNIQUE COLLATE NOCASE,
                status TEXT NOT NULL DEFAULT 'pending' CHECK (status IN ('done','progress','pending')),
                created_at TEXT NOT NULL DEFAULT (datetime('now','localtime')),
                updated_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS evaluation_periods (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                starts_on TEXT,
                ends_on TEXT,
                is_active INTEGER NOT NULL DEFAULT 0 CHECK (is_active IN (0,1)),
                is_archived INTEGER NOT NULL DEFAULT 0 CHECK (is_archived IN (0,1)),
                created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            )
        """)
        if not table_exists(conn, "users"):
            conn.execute("""
                CREATE TABLE users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT NOT NULL UNIQUE COLLATE NOCASE,
                    display_name TEXT NOT NULL,
                    password_hash TEXT NOT NULL,
                    role TEXT NOT NULL CHECK (role IN ('admin','evaluator')),
                    evaluator_no INTEGER,
                    max_points REAL NOT NULL DEFAULT 0 CHECK (max_points >= 0 AND max_points <= 100),
                    section_name TEXT NOT NULL DEFAULT '',
                    active INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0,1)),
                    created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
                )
            """)
        else:
            rebuild_legacy_users(conn)
        # Store each user's optional profile photo in SQLite so it travels with the database.
        user_cols = table_columns(conn, "users")
        if "deleted_at" not in user_cols:
            conn.execute("ALTER TABLE users ADD COLUMN deleted_at TEXT")
        if "profile_image" not in user_cols:
            conn.execute("ALTER TABLE users ADD COLUMN profile_image BLOB")
        if "profile_image_mime" not in user_cols:
            conn.execute("ALTER TABLE users ADD COLUMN profile_image_mime TEXT NOT NULL DEFAULT ''")
        rebuild_legacy_questions(conn)
        rebuild_legacy_scores(conn)
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS evaluation_submissions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                period_id INTEGER NOT NULL REFERENCES evaluation_periods(id) ON DELETE RESTRICT,
                unit_id INTEGER NOT NULL REFERENCES units(id) ON DELETE RESTRICT,
                user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE RESTRICT,
                status TEXT NOT NULL DEFAULT 'draft' CHECK (status IN ('draft','submitted')),
                submitted_at TEXT,
                updated_at TEXT NOT NULL DEFAULT (datetime('now','localtime')),
                last_modified_by INTEGER REFERENCES users(id),
                UNIQUE(period_id, unit_id, user_id)
            );
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                actor_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
                action TEXT NOT NULL,
                entity TEXT NOT NULL,
                entity_id INTEGER,
                detail TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            );
            CREATE INDEX IF NOT EXISTS idx_questions_owner_active ON questions(assigned_user_id, active);
            CREATE INDEX IF NOT EXISTS idx_scores_unit_period ON scores(unit_id, period_id);
            CREATE INDEX IF NOT EXISTS idx_submissions_period_unit ON evaluation_submissions(period_id, unit_id, status);
            CREATE INDEX IF NOT EXISTS idx_submissions_date ON evaluation_submissions(submitted_at);
        """)
        # Repair legacy question ownership based on its previous evaluator number.
        if table_exists(conn, "questions"):
            conn.execute("""
                UPDATE questions SET assigned_user_id = (
                    SELECT users.id FROM users WHERE users.evaluator_no = questions.evaluator_no LIMIT 1
                ) WHERE assigned_user_id IS NULL AND evaluator_no IS NOT NULL
            """)
            conn.execute("""
                UPDATE questions SET section_name = COALESCE(NULLIF(section_name,''),
                    (SELECT users.section_name FROM users WHERE users.id=questions.assigned_user_id), '')
                WHERE section_name IS NULL OR section_name=''
            """)
        user_count = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        if user_count == 0:
            initial_users = [("admin", "مدیر سامانه", generate_password_hash(os.environ.get("INITIAL_ADMIN_PASSWORD", "admin123")), "admin", None, 0, "")]
            for n in range(1, 8):
                points = 10 if n == 1 else 15
                initial_users.append((f"evaluator{n}", f"ارزیاب شماره {n}", generate_password_hash(os.environ.get("INITIAL_EVALUATOR_PASSWORD", "eval123")), "evaluator", n, points, f"بخش ارزیابی {n}"))
            conn.executemany(
                "INSERT INTO users (username,display_name,password_hash,role,evaluator_no,max_points,section_name) VALUES (?,?,?,?,?,?,?)",
                initial_users,
            )
            log_event(conn, None, "seed_initial_accounts", "users", None, "ایجاد حساب‌های اولیه سامانه")
        # Dedicated primary administrator account; password is hashed like every other account.
        if not conn.execute("SELECT 1 FROM users WHERE username=? COLLATE NOCASE", ("momo",)).fetchone():
            conn.execute("INSERT INTO users(username,display_name,password_hash,role,evaluator_no,max_points,section_name,active) VALUES(?,?,?,?,NULL,0,'',1)",
                         ("momo", "مدیر اصلی سامانه", generate_password_hash(os.environ.get("PRIMARY_ADMIN_PASSWORD", "momo2734")), "admin"))
            log_event(conn, None, "seed_primary_admin", "users", None, "ایجاد حساب مدیر اصلی")
        if conn.execute("SELECT COUNT(*) FROM evaluation_periods").fetchone()[0] == 0:
            conn.execute("INSERT INTO evaluation_periods (name, starts_on, is_active) VALUES (?,?,1)", ("دوره جاری", datetime.now().date().isoformat()))
        conn.commit()
    finally:
        conn.close()
    # Re-open with FK checks on so normal requests always validate relational integrity.
    conn = connect_db()
    try:
        conn.execute("PRAGMA foreign_key_check")
    finally:
        conn.close()


@app.before_request
def load_current_user() -> None:
    g.user = None
    user_id = session.get("user_id")
    if user_id is None:
        return
    conn = connect_db()
    try:
        g.user = conn.execute(
            "SELECT id,username,display_name,role,evaluator_no,max_points,section_name,active,deleted_at,profile_image FROM users WHERE id=? AND deleted_at IS NULL",
            (user_id,),
        ).fetchone()
    finally:
        conn.close()
    if g.user is None or not g.user["active"]:
        session.clear()
        g.user = None


def user_payload(user: sqlite3.Row | dict[str, Any]) -> dict[str, Any]:
    role = user["role"]
    if role == "admin":
        role_label = "مدیر سامانه"
        avatar = (user["display_name"] or "م")[:1]
    else:
        role_label = (user["section_name"] or "ارزیاب")
        avatar = (user["display_name"] or user["username"] or "ا")[:1]
    return {
        "id": user["id"], "username": user["username"], "display_name": user["display_name"],
        "role": role, "is_primary_admin": bool(role == "admin" and user["username"].lower() == "momo"), "evaluator_no": user["evaluator_no"], "max_points": float(user["max_points"] or 0),
        "section_name": user["section_name"] or "", "active": bool(user["active"]), "deleted": bool(user["deleted_at"]),
        "role_label": role_label, "avatar": avatar,
        "has_profile_image": bool(user["profile_image"]),
    }


def login_required(fn: Callable[..., Any]) -> Callable[..., Any]:
    @wraps(fn)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        if g.user is None:
            return jsonify({"ok": False, "error": "برای ادامه وارد سامانه شوید."}), 401
        return fn(*args, **kwargs)
    return wrapper


def admin_required(fn: Callable[..., Any]) -> Callable[..., Any]:
    @wraps(fn)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        if g.user is None:
            return jsonify({"ok": False, "error": "برای ادامه وارد سامانه شوید."}), 401
        if g.user["role"] != "admin":
            return jsonify({"ok": False, "error": "این عملیات فقط برای مدیر سامانه مجاز است."}), 403
        return fn(*args, **kwargs)
    return wrapper


def primary_admin_required(fn: Callable[..., Any]) -> Callable[..., Any]:
    @wraps(fn)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        if g.user is None:
            return jsonify({"ok": False, "error": "برای ادامه وارد سامانه شوید."}), 401
        if g.user["role"] != "admin" or str(g.user["username"]).lower() != "momo":
            return jsonify({"ok": False, "error": "این عملیات فقط برای مدیر اصلی سامانه مجاز است."}), 403
        return fn(*args, **kwargs)
    return wrapper


def json_error(message: str, status: int = 400):
    return jsonify({"ok": False, "error": message}), status


def get_active_period(conn: sqlite3.Connection) -> sqlite3.Row | None:
    return conn.execute("SELECT * FROM evaluation_periods WHERE is_active=1 AND is_archived=0 ORDER BY id DESC LIMIT 1").fetchone()


def get_target_evaluator(conn: sqlite3.Connection, requested_id: Any) -> sqlite3.Row | None:
    if g.user["role"] != "admin":
        return conn.execute("SELECT * FROM users WHERE id=? AND role='evaluator'", (g.user["id"],)).fetchone()
    if requested_id in (None, "", 0, "0"):
        return None
    try:
        target_id = int(requested_id)
    except (ValueError, TypeError):
        return None
    return conn.execute("SELECT * FROM users WHERE id=? AND role='evaluator'", (target_id,)).fetchone()


def active_evaluator_ids(conn: sqlite3.Connection) -> list[int]:
    # Every active evaluator with a positive allocation is required to submit each unit.
    return [int(r[0]) for r in conn.execute("SELECT id FROM users WHERE role='evaluator' AND active=1 AND max_points>0 ORDER BY id").fetchall()]


def unit_summary(conn: sqlite3.Connection, unit: sqlite3.Row, period: sqlite3.Row | None) -> dict[str, Any]:
    unit_data = dict(unit)
    if not period:
        unit_data.update({"score": None, "status_label": "دوره‌ای تعریف نشده", "completed_by": [], "submission_count": 0, "required_count": 0, "total_max": 0})
        return unit_data
    required_ids = active_evaluator_ids(conn)
    submitted_rows = conn.execute(
        "SELECT s.user_id,u.display_name,s.submitted_at FROM evaluation_submissions s JOIN users u ON u.id=s.user_id WHERE s.period_id=? AND s.unit_id=? AND s.status='submitted' ORDER BY s.submitted_at",
        (period["id"], unit["id"]),
    ).fetchall()
    submitted = {r["user_id"] for r in submitted_rows}
    completed_by = [r["display_name"] for r in submitted_rows]
    any_score = conn.execute("SELECT 1 FROM scores WHERE period_id=? AND unit_id=? LIMIT 1", (period["id"], unit["id"])).fetchone() is not None
    complete = bool(required_ids) and set(required_ids).issubset(submitted)
    total = conn.execute(
        "SELECT COALESCE(SUM(sc.score),0) FROM scores sc JOIN evaluation_submissions sub ON sub.period_id=sc.period_id AND sub.unit_id=sc.unit_id AND sub.user_id=sc.evaluator_user_id AND sub.status='submitted' WHERE sc.period_id=? AND sc.unit_id=?",
        (period["id"], unit["id"]),
    ).fetchone()[0]
    if complete:
        status, label = "done", "تکمیل‌شده"
        score: float | None = round(float(total), 2)
    elif submitted or any_score:
        status, label = "progress", "در حال ارزیابی"
        score = None
    else:
        status, label = "pending", "در انتظار شروع"
        score = None
    unit_data.update({"status": status, "status_label": label, "score": score, "completed_by": completed_by,
                      "submission_count": len(submitted), "required_count": len(required_ids), "total_max": 100})
    return unit_data


@app.get("/")
def index() -> str:
    return render_template("index.html")


@app.post("/api/login")
def api_login() -> Any:
    payload = request.get_json(silent=True) or {}
    username = str(payload.get("username", "")).strip()
    password = str(payload.get("password", ""))
    if not username or not password:
        return json_error("نام کاربری و رمز عبور را وارد کنید.")
    conn = connect_db()
    try:
        user = conn.execute("SELECT * FROM users WHERE username=? COLLATE NOCASE", (username,)).fetchone()
        if user is None or user["deleted_at"] or not user["active"] or not check_password_hash(user["password_hash"], password):
            log_event(conn, user["id"] if user else None, "login_failed", "auth", None, f"نام کاربری: {username[:80]}")
            conn.commit()
            return json_error("نام کاربری یا رمز عبور صحیح نیست.", 401)
        session.clear()
        session["user_id"] = user["id"]
        session.permanent = True
        log_event(conn, user["id"], "login_success", "auth", user["id"], "ورود موفق")
        conn.commit()
        return jsonify({"ok": True, "user": user_payload(user)})
    finally:
        conn.close()


@app.get("/api/me")
def api_me() -> Any:
    if g.user is None:
        return jsonify({"authenticated": False})
    return jsonify({"authenticated": True, "user": user_payload(g.user)})


@app.post("/api/logout")
def api_logout() -> Any:
    if g.user is not None:
        conn = connect_db()
        try:
            log_event(conn, g.user["id"], "logout", "auth", g.user["id"], "خروج از سامانه")
            conn.commit()
        finally:
            conn.close()
    session.clear()
    return jsonify({"ok": True})


@app.get("/api/periods")
@login_required
def get_periods() -> Any:
    conn = connect_db()
    try:
        rows = conn.execute("SELECT * FROM evaluation_periods WHERE is_archived=0 ORDER BY id DESC").fetchall()
        active = get_active_period(conn)
        return jsonify({"ok": True, "periods": [dict(r) for r in rows], "active_period_id": active["id"] if active else None})
    finally:
        conn.close()


@app.post("/api/periods")
@admin_required
def create_period() -> Any:
    payload = request.get_json(silent=True) or {}
    name = str(payload.get("name", "")).strip()
    starts_on = str(payload.get("starts_on", "")).strip() or None
    ends_on = str(payload.get("ends_on", "")).strip() or None
    if not name or len(name) > 100:
        return json_error("نام دوره الزامی است و نباید بیش از ۱۰۰ نویسه باشد.")
    if starts_on and ends_on and starts_on > ends_on:
        return json_error("تاریخ پایان نمی‌تواند قبل از تاریخ شروع باشد.")
    conn = connect_db()
    try:
        conn.execute("UPDATE evaluation_periods SET is_active=0")
        cur = conn.execute("INSERT INTO evaluation_periods (name,starts_on,ends_on,is_active) VALUES (?,?,?,1)", (name, starts_on, ends_on))
        period_id = cur.lastrowid
        log_event(conn, g.user["id"], "create", "period", period_id, f"ایجاد و فعال‌سازی دوره {name}")
        conn.commit()
        return jsonify({"ok": True, "period_id": period_id}), 201
    finally:
        conn.close()


@app.post("/api/periods/<int:period_id>/activate")
@admin_required
def activate_period(period_id: int) -> Any:
    conn = connect_db()
    try:
        row = conn.execute("SELECT * FROM evaluation_periods WHERE id=? AND is_archived=0", (period_id,)).fetchone()
        if not row:
            return json_error("دوره پیدا نشد.", 404)
        conn.execute("UPDATE evaluation_periods SET is_active=0")
        conn.execute("UPDATE evaluation_periods SET is_active=1 WHERE id=?", (period_id,))
        log_event(conn, g.user["id"], "activate", "period", period_id, f"فعال‌سازی دوره {row['name']}")
        conn.commit()
        return jsonify({"ok": True})
    finally:
        conn.close()


@app.put("/api/periods/<int:period_id>")
@admin_required
def update_period(period_id: int) -> Any:
    payload = request.get_json(silent=True) or {}
    name = str(payload.get("name", "")).strip()
    starts_on = str(payload.get("starts_on", "")).strip() or None
    ends_on = str(payload.get("ends_on", "")).strip() or None
    if not name or len(name) > 100:
        return json_error("نام دوره الزامی است و نباید بیش از ۱۰۰ نویسه باشد.")
    if starts_on and ends_on and starts_on > ends_on:
        return json_error("تاریخ پایان نمی‌تواند قبل از تاریخ شروع باشد.")
    conn = connect_db()
    try:
        row = conn.execute("SELECT * FROM evaluation_periods WHERE id=? AND is_archived=0", (period_id,)).fetchone()
        if not row: return json_error("دوره پیدا نشد.", 404)
        conn.execute("UPDATE evaluation_periods SET name=?,starts_on=?,ends_on=? WHERE id=?", (name, starts_on, ends_on, period_id))
        log_event(conn, g.user["id"], "update", "period", period_id, f"ویرایش دوره {name}")
        conn.commit()
        return jsonify({"ok": True})
    finally: conn.close()


@app.delete("/api/periods/<int:period_id>")
@admin_required
def delete_period(period_id: int) -> Any:
    conn = connect_db()
    try:
        row = conn.execute("SELECT * FROM evaluation_periods WHERE id=? AND is_archived=0", (period_id,)).fetchone()
        if not row: return json_error("دوره پیدا نشد.", 404)
        if row["is_active"]: return json_error("ابتدا یک دوره دیگر را فعال کنید؛ دوره فعال قابل حذف نیست.", 409)
        conn.execute("UPDATE evaluation_periods SET is_archived=1 WHERE id=?", (period_id,))
        log_event(conn, g.user["id"], "archive", "period", period_id, f"حذف منطقی دوره {row['name']}")
        conn.commit()
        return jsonify({"ok": True})
    finally: conn.close()


@app.get("/api/units")
@login_required
def get_units() -> Any:
    period_id = request.args.get("period_id", type=int)
    conn = connect_db()
    try:
        period = conn.execute("SELECT * FROM evaluation_periods WHERE id=? AND is_archived=0", (period_id,)).fetchone() if period_id else get_active_period(conn)
        rows = conn.execute("SELECT * FROM units ORDER BY name COLLATE NOCASE").fetchall()
        if g.user["role"] == "admin":
            units = [unit_summary(conn, row, period) for row in rows]
        else:
            units = []
            for row in rows:
                item = dict(row)
                sub = conn.execute("SELECT status,submitted_at FROM evaluation_submissions WHERE period_id=? AND unit_id=? AND user_id=?", (period["id"], row["id"], g.user["id"])).fetchone() if period else None
                score = conn.execute("SELECT SUM(score) FROM scores WHERE period_id=? AND unit_id=? AND evaluator_user_id=?", (period["id"], row["id"], g.user["id"])).fetchone()[0] if period else None
                item.update({"status":"done" if sub and sub["status"]=="submitted" else ("progress" if sub else "pending"),
                             "status_label":"ثبت نهایی شما" if sub and sub["status"]=="submitted" else ("ذخیره موقت شما" if sub else "در انتظار ارزیابی شما"),
                             "score": round(float(score or 0),2) if sub and sub["status"]=="submitted" else None,
                             "completed_by": [], "submission_count": 1 if sub and sub["status"]=="submitted" else 0,
                             "required_count": 1, "total_max": float(g.user["max_points"] or 0)})
                units.append(item)
        return jsonify({"ok": True, "units": units, "period_id": period["id"] if period else None})
    finally:
        conn.close()


@app.post("/api/units")
@admin_required
def create_unit() -> Any:
    payload = request.get_json(silent=True) or {}
    name = str(payload.get("name", "")).strip()
    code = str(payload.get("code", "")).strip()
    if not name:
        return json_error("نام زیررده الزامی است.")
    if len(name) > 80 or len(code) > 20:
        return json_error("نام یا کد بیش از حد طولانی است.")
    conn = connect_db()
    try:
        if not code:
            next_num = conn.execute("SELECT COALESCE(MAX(id),0)+1 FROM units").fetchone()[0]
            code = f"U-{next_num:03d}"
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur = conn.execute("INSERT INTO units(code,name,status,created_at,updated_at) VALUES(?,?, 'pending', ?,?)", (code, name, now, now))
        unit_id = cur.lastrowid
        log_event(conn, g.user["id"], "create", "unit", unit_id, f"ایجاد زیررده {name}")
        conn.commit()
        unit = conn.execute("SELECT * FROM units WHERE id=?", (unit_id,)).fetchone()
        return jsonify({"ok": True, "unit": dict(unit)}), 201
    except sqlite3.IntegrityError:
        conn.rollback()
        return json_error("نام زیررده یا کد آن تکراری است.", 409)
    finally:
        conn.close()


@app.delete("/api/units/<int:unit_id>")
@admin_required
def delete_unit(unit_id: int) -> Any:
    conn = connect_db()
    try:
        unit = conn.execute("SELECT * FROM units WHERE id=?", (unit_id,)).fetchone()
        if not unit:
            return json_error("زیررده پیدا نشد.", 404)
        try:
            conn.execute("DELETE FROM units WHERE id=?", (unit_id,))
            log_event(conn, g.user["id"], "delete", "unit", unit_id, f"حذف زیررده {unit['name']}")
            conn.commit()
        except sqlite3.IntegrityError:
            conn.rollback()
            return json_error("برای این زیررده سابقه ارزیابی ثبت شده؛ حذف آن مجاز نیست.", 409)
        return jsonify({"ok": True})
    finally:
        conn.close()




@app.put("/api/units/<int:unit_id>")
@admin_required
def update_unit(unit_id: int) -> Any:
    payload = request.get_json(silent=True) or {}
    name = str(payload.get("name", "")).strip()
    if not name or len(name) > 80:
        return json_error("نام زیررده الزامی است و نباید بیش از ۸۰ نویسه باشد.")
    conn = connect_db()
    try:
        unit = conn.execute("SELECT * FROM units WHERE id=?", (unit_id,)).fetchone()
        if not unit:
            return json_error("زیررده پیدا نشد.", 404)
        conn.execute("UPDATE units SET name=?,updated_at=datetime('now','localtime') WHERE id=?", (name, unit_id))
        log_event(conn, g.user["id"], "rename", "unit", unit_id, f"تغییر نام زیررده از «{unit['name']}» به «{name}»؛ شناسه حفظ شد")
        conn.commit()
        return jsonify({"ok": True, "unit": dict(conn.execute("SELECT * FROM units WHERE id=?", (unit_id,)).fetchone())})
    except sqlite3.IntegrityError:
        conn.rollback()
        return json_error("این نام زیررده قبلاً استفاده شده است.", 409)
    finally:
        conn.close()


@app.get("/api/my/question-settings")
@login_required
def get_my_question_settings() -> Any:
    if g.user["role"] != "evaluator":
        return json_error("این بخش فقط برای ارزیاب‌هاست.", 403)
    conn = connect_db()
    try:
        rows = conn.execute("SELECT id,question_text,rubric,max_score,section_name,version FROM questions WHERE assigned_user_id=? AND active=1 ORDER BY id", (g.user["id"],)).fetchall()
        total = sum(float(r["max_score"]) for r in rows)
        return jsonify({"ok": True, "questions": [dict(r) for r in rows], "total": total,
                        "max_points": float(g.user["max_points"] or 0),
                        "matches": abs(total-float(g.user["max_points"] or 0)) < 0.00001})
    finally:
        conn.close()


@app.put("/api/my/question-settings")
@login_required
def update_my_question_settings() -> Any:
    if g.user["role"] != "evaluator":
        return json_error("این بخش فقط برای ارزیاب‌هاست.", 403)
    payload = request.get_json(silent=True) or {}
    items = payload.get("questions")
    if not isinstance(items, list) or not items:
        return json_error("فهرست بارم سؤال‌ها معتبر نیست.")
    conn = connect_db()
    try:
        conn.execute("BEGIN IMMEDIATE")
        current = conn.execute("SELECT id,max_score FROM questions WHERE assigned_user_id=? AND active=1 ORDER BY id", (g.user["id"],)).fetchall()
        current_ids = {int(r["id"]) for r in current}
        updates: dict[int, float] = {}
        for item in items:
            if not isinstance(item, dict):
                return json_error("ساختار یکی از بارم‌ها معتبر نیست.")
            try:
                qid, points = int(item.get("id")), float(item.get("max_score"))
            except (ValueError, TypeError):
                return json_error("شناسه سؤال یا بارم معتبر نیست.")
            if qid not in current_ids or qid in updates:
                return json_error("سؤال نامعتبر یا تکراری است.", 403)
            if not (0 < points <= 100):
                return json_error("بارم هر سؤال باید بیشتر از صفر و حداکثر ۱۰۰ باشد.")
            updates[qid] = points
        if set(updates) != current_ids:
            return json_error("بارم تمام سؤال‌های فعال باید یکجا ارسال شود.")
        quota = float(g.user["max_points"] or 0)
        active_period = get_active_period(conn)
        if active_period:
            submitted = conn.execute(
                "SELECT 1 FROM evaluation_submissions WHERE period_id=? AND user_id=? AND status='submitted' LIMIT 1",
                (active_period["id"], g.user["id"])
            ).fetchone()
            changed_ids = [qid for qid, points in updates.items()
                           if abs(float(next(r["max_score"] for r in current if int(r["id"]) == qid))-points) > 0.00001]
            if changed_ids and submitted:
                return json_error("پس از ثبت نهایی ارزیابی در دوره فعال، تغییر بارم سؤال‌ها تا دوره بعد مجاز نیست.", 409)
            if changed_ids:
                marks = ",".join("?" for _ in changed_ids)
                has_scores = conn.execute(f"SELECT 1 FROM scores WHERE period_id=? AND evaluator_user_id=? AND question_id IN ({marks}) LIMIT 1",
                                          (active_period["id"], g.user["id"], *changed_ids)).fetchone()
                if has_scores:
                    return json_error("برای سؤال‌هایی که در دوره فعال نمره دارند، ابتدا ارزیابی‌های جاری را تعیین تکلیف کنید؛ تغییر بارم به دوره بعد موکول شود.", 409)
        total = sum(updates.values())
        if abs(total-quota) > 0.00001:
            return json_error(f"مجموع بارم‌ها باید دقیقاً برابر سقف تخصیص‌یافته ({quota:g}) باشد؛ مجموع واردشده {total:g} است.")
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for qid, points in updates.items():
            old = next(r for r in current if int(r["id"]) == qid)
            if abs(float(old["max_score"])-points) > 0.00001:
                conn.execute("UPDATE questions SET max_score=?,version=version+1,updated_at=? WHERE id=? AND assigned_user_id=? AND active=1",
                             (points, now, qid, g.user["id"]))
                active_period = get_active_period(conn)
                period_context = f"؛ دوره={active_period['name']}" if active_period else ""
                log_event(conn, g.user["id"], "change_question_rubric", "question", qid,
                          f"تغییر بارم سؤال؛ مقدار قبلی={float(old['max_score']):g}؛ مقدار جدید={points:g}{period_context}")
        conn.commit()
        return jsonify({"ok": True, "total": total, "message": "بارم سؤال‌ها با موفقیت ذخیره شد."})
    except sqlite3.Error:
        conn.rollback()
        return json_error("ذخیره بارم‌ها انجام نشد؛ هیچ تغییر ناقصی اعمال نشد.", 500)
    finally:
        conn.close()


@app.get("/api/admin/audit")
@primary_admin_required
def get_admin_audit() -> Any:
    period_id = request.args.get("period_id", type=int)
    actor_id = request.args.get("user_id", type=int)
    since = str(request.args.get("since", "")).strip()
    until = str(request.args.get("until", "")).strip()
    conn = connect_db()
    try:
        where, args = ["1=1"], []
        if actor_id:
            where.append("a.actor_id=?"); args.append(actor_id)
        if since:
            where.append("a.created_at>=?"); args.append(since)
        if until:
            where.append("a.created_at<=?"); args.append(until + " 23:59:59")
        if period_id:
            # Period names are stored in the audit detail for relevant evaluation changes.
            pname = conn.execute("SELECT name FROM evaluation_periods WHERE id=?", (period_id,)).fetchone()
            where.append("(a.entity NOT IN ('evaluation','question') OR a.detail LIKE ?)")
            args.append(f"%دوره={pname['name']}%" if pname else "%__no_period__%")
        rows = conn.execute(f"""
            SELECT a.id,a.actor_id,COALESCE(u.display_name,'کاربر حذف‌شده') AS actor_name,
                   a.action,a.entity,a.entity_id,a.detail,a.created_at
            FROM audit_log a LEFT JOIN users u ON u.id=a.actor_id
            WHERE {' AND '.join(where)}
            ORDER BY a.id DESC LIMIT 500
        """, tuple(args)).fetchall()
        return jsonify({"ok": True, "events": [dict(r) for r in rows]})
    finally:
        conn.close()


@app.get("/api/users")
@admin_required
def get_users() -> Any:
    conn = connect_db()
    try:
        rows = conn.execute("""
            SELECT u.*, COALESCE((SELECT SUM(q.max_score) FROM questions q WHERE q.assigned_user_id=u.id AND q.active=1),0) AS question_points,
                   (SELECT COUNT(*) FROM questions q WHERE q.assigned_user_id=u.id AND q.active=1) AS question_count
            FROM users u ORDER BY CASE u.role WHEN 'admin' THEN 0 ELSE 1 END, u.display_name COLLATE NOCASE
        """).fetchall()
        allocation = conn.execute("SELECT COALESCE(SUM(max_points),0) FROM users WHERE role='evaluator' AND active=1").fetchone()[0]
        return jsonify({"ok": True, "users": [{**user_payload(r), "question_points": float(r["question_points"] or 0), "question_count": r["question_count"]} for r in rows], "allocation_total": float(allocation)})
    finally:
        conn.close()


@app.post("/api/users")
@admin_required
def create_user() -> Any:
    payload = request.get_json(silent=True) or {}
    username = str(payload.get("username", "")).strip()
    display_name = str(payload.get("display_name", "")).strip()
    password = str(payload.get("password", ""))
    role = str(payload.get("role", "evaluator"))
    section_name = str(payload.get("section_name", "")).strip()
    active = 1 if payload.get("active", True) else 0
    try:
        max_points = float(payload.get("max_points", 0) or 0) if role == "evaluator" else 0.0
    except (ValueError, TypeError):
        return json_error("سقف نمره باید عددی باشد.")
    if len(username) < 3 or len(username) > 50 or any(ch.isspace() for ch in username):
        return json_error("نام کاربری باید ۳ تا ۵۰ نویسه و بدون فاصله باشد.")
    if not display_name or len(display_name) > 100:
        return json_error("نام نمایشی را وارد کنید.")
    if len(password) < 8:
        return json_error("رمز عبور باید حداقل ۸ نویسه داشته باشد.")
    if role not in ("admin", "evaluator"):
        return json_error("نوع حساب معتبر نیست.")
    if not (0 <= max_points <= 100):
        return json_error("سقف نمره باید بین صفر تا ۱۰۰ باشد.")
    if role == "evaluator" and not section_name:
        section_name = display_name

    conn = connect_db()
    try:
        # A deleted username remains unique in SQLite by design. Reuse its original
        # row instead of inserting a duplicate, preserving all historical foreign keys.
        existing = conn.execute(
            "SELECT * FROM users WHERE username=? COLLATE NOCASE", (username,)
        ).fetchone()
        if existing and not existing["deleted_at"]:
            return json_error("این نام کاربری قبلاً ثبت شده است. نام دیگری انتخاب کنید.", 409)

        restoring_id = int(existing["id"]) if existing else None
        allocation = float(conn.execute(
            "SELECT COALESCE(SUM(max_points),0) FROM users "
            "WHERE role='evaluator' AND active=1 AND deleted_at IS NULL AND id<>COALESCE(?, -1)",
            (restoring_id,)
        ).fetchone()[0] or 0)
        if role == "evaluator" and active and allocation + max_points > 100.00001:
            return json_error(
                f"مجموع سقف نمرهٔ ارزیاب‌های فعال از ۱۰۰ بیشتر می‌شود. "
                f"سقف آزاد فعلی: {max(0,100-allocation):g}."
            )

        if restoring_id is not None:
            active_questions = int(conn.execute(
                "SELECT COUNT(*) FROM questions WHERE assigned_user_id=? AND active=1",
                (restoring_id,)
            ).fetchone()[0])
            if role != "evaluator" and active_questions:
                return json_error(
                    "این کاربر سؤال فعال دارد؛ برای بازیابی با نقش مدیر، ابتدا سؤال‌های فعال او را منتقل یا غیرفعال کنید."
                )
            if role == "evaluator":
                evaluator_no = existing["evaluator_no"]
                if evaluator_no is None:
                    evaluator_no = conn.execute(
                        "SELECT COALESCE(MAX(evaluator_no),0)+1 FROM users WHERE role='evaluator'"
                    ).fetchone()[0]
            else:
                evaluator_no = None
            conn.execute("""
                UPDATE users
                SET username=?, display_name=?, password_hash=?, role=?, evaluator_no=?,
                    max_points=?, section_name=?, active=?, deleted_at=NULL
                WHERE id=?
            """, (
                username, display_name, generate_password_hash(password), role,
                evaluator_no, max_points, section_name, active, restoring_id
            ))
            user_id = restoring_id
            action = "restore"
            description = f"بازیابی و ثبت مجدد کاربر حذف‌شده {username} با حفظ سوابق"
        else:
            current_max = conn.execute(
                "SELECT COALESCE(MAX(evaluator_no),0)+1 FROM users WHERE role='evaluator'"
            ).fetchone()[0]
            cur = conn.execute(
                "INSERT INTO users(username,display_name,password_hash,role,evaluator_no,max_points,section_name,active) "
                "VALUES(?,?,?,?,?,?,?,?)",
                (username, display_name, generate_password_hash(password), role,
                 current_max if role == "evaluator" else None, max_points, section_name, active),
            )
            user_id = cur.lastrowid
            action = "create"
            description = f"ایجاد کاربر {username} با نقش {role}"

        if role == "evaluator" and active:
            rebalance_question_points(conn, user_id)
        log_event(conn, g.user["id"], action, "user", user_id, description)
        conn.commit()
        row = conn.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
        return jsonify({"ok": True, "user": user_payload(row), "restored": restoring_id is not None}), (200 if restoring_id is not None else 201)
    except sqlite3.IntegrityError:
        conn.rollback()
        return json_error("ثبت کاربر به دلیل تداخل نام کاربری یا اطلاعات پایگاه داده انجام نشد.", 409)
    finally:
        conn.close()


@app.put("/api/users/<int:user_id>")
@admin_required
def update_user(user_id: int) -> Any:
    payload = request.get_json(silent=True) or {}
    conn = connect_db()
    try:
        old = conn.execute("SELECT * FROM users WHERE id=? AND deleted_at IS NULL", (user_id,)).fetchone()
        if not old:
            return json_error("کاربر پیدا نشد.", 404)
        role = str(payload.get("role", old["role"]))
        username = str(payload.get("username", old["username"])).strip()
        display_name = str(payload.get("display_name", old["display_name"])).strip()
        section_name = str(payload.get("section_name", old["section_name"] or "")).strip()
        active = 1 if payload.get("active", bool(old["active"])) else 0
        password = str(payload.get("password", ""))
        try:
            max_points = float(payload.get("max_points", old["max_points"]) or 0) if role == "evaluator" else 0
        except (ValueError, TypeError):
            return json_error("سقف نمره باید عددی باشد.")
        if user_id == g.user["id"] and (role != "admin" or not active):
            return json_error("نمی‌توانید نقش یا وضعیت حساب مدیرِ در حال استفاده را تغییر دهید.", 400)
        if len(username) < 3 or any(ch.isspace() for ch in username) or not display_name:
            return json_error("نام کاربری یا نام نمایشی معتبر نیست.")
        if password and len(password) < 8:
            return json_error("رمز عبور جدید باید حداقل ۸ نویسه داشته باشد.")
        if role not in ("admin", "evaluator") or not 0 <= max_points <= 100:
            return json_error("نقش یا سقف نمره معتبر نیست.")
        if role == "evaluator" and not section_name:
            section_name = display_name
        allocation = float(conn.execute("SELECT COALESCE(SUM(max_points),0) FROM users WHERE role='evaluator' AND active=1 AND id<>?", (user_id,)).fetchone()[0])
        if role == "evaluator" and active and allocation + max_points > 100.00001:
            return json_error(f"مجموع سقف نمرهٔ ارزیاب‌های فعال از ۱۰۰ بیشتر می‌شود. سقف آزاد فعلی: {max(0,100-allocation):g}.")
        question_count = int(conn.execute("SELECT COUNT(*) FROM questions WHERE assigned_user_id=? AND active=1", (user_id,)).fetchone()[0])
        if role != "evaluator" and question_count > 0:
            return json_error("این کاربر سؤال فعال دارد؛ ابتدا سؤال‌های او را به ارزیاب دیگری منتقل یا غیرفعال کنید.")
        fields = [username, display_name, role, max_points, section_name, active]
        sql = "UPDATE users SET username=?,display_name=?,role=?,max_points=?,section_name=?,active=?"
        if password:
            sql += ",password_hash=?"
            fields.append(generate_password_hash(password))
        sql += " WHERE id=?"
        fields.append(user_id)
        conn.execute(sql, tuple(fields))
        if role == "evaluator" and active:
            rebalance_question_points(conn, user_id)
        log_event(conn, g.user["id"], "update", "user", user_id, f"ویرایش کاربر {username}")
        conn.commit()
        row = conn.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
        return jsonify({"ok": True, "user": user_payload(row)})
    except sqlite3.IntegrityError:
        conn.rollback()
        return json_error("نام کاربری تکراری است.", 409)
    finally:
        conn.close()


@app.delete("/api/users/<int:user_id>")
@admin_required
def delete_user(user_id: int) -> Any:
    if user_id == g.user["id"]:
        return json_error("نمی‌توانید حساب مدیرِ در حال استفاده را حذف کنید.", 400)
    conn = connect_db()
    try:
        row = conn.execute("SELECT * FROM users WHERE id=? AND deleted_at IS NULL", (user_id,)).fetchone()
        if not row: return json_error("کاربر پیدا نشد.", 404)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn.execute("UPDATE users SET active=0,deleted_at=? WHERE id=?", (now, user_id))
        log_event(conn, g.user["id"], "soft_delete", "user", user_id, f"حذف منطقی کاربر {row['username']}")
        conn.commit()
        return jsonify({"ok": True})
    finally: conn.close()




@app.post("/api/users/<int:user_id>/restore")
@admin_required
def restore_user(user_id: int) -> Any:
    """Restore a soft-deleted account in-place so its historical references remain valid."""
    payload = request.get_json(silent=True) or {}
    password = str(payload.get("password", ""))
    if len(password) < 8:
        return json_error("برای بازیابی، رمز عبور جدید باید حداقل ۸ نویسه داشته باشد.")
    conn = connect_db()
    try:
        row = conn.execute("SELECT * FROM users WHERE id=? AND deleted_at IS NOT NULL", (user_id,)).fetchone()
        if not row:
            return json_error("کاربر حذف‌شده پیدا نشد یا قبلاً بازیابی شده است.", 404)
        role = row["role"]
        active = 1
        max_points = float(row["max_points"] or 0) if role == "evaluator" else 0.0
        if role == "evaluator":
            allocation = float(conn.execute(
                "SELECT COALESCE(SUM(max_points),0) FROM users "
                "WHERE role='evaluator' AND active=1 AND deleted_at IS NULL AND id<>?",
                (user_id,)
            ).fetchone()[0] or 0)
            if allocation + max_points > 100.00001:
                return json_error(
                    f"کاربر قابل بازیابی است، اما سقف امتیاز آزاد کافی نیست. "
                    f"سقف آزاد فعلی: {max(0, 100-allocation):g}؛ سقف این ارزیاب: {max_points:g}. "
                    "ابتدا سهم ارزیابان فعال را تنظیم کنید."
                )
            evaluator_no = row["evaluator_no"]
            if evaluator_no is None:
                evaluator_no = conn.execute(
                    "SELECT COALESCE(MAX(evaluator_no),0)+1 FROM users WHERE role='evaluator'"
                ).fetchone()[0]
        else:
            evaluator_no = None
            active_questions = int(conn.execute(
                "SELECT COUNT(*) FROM questions WHERE assigned_user_id=? AND active=1", (user_id,)
            ).fetchone()[0])
            if active_questions:
                return json_error("این حساب سؤال فعال دارد و امکان بازیابی آن با نقش مدیر وجود ندارد.")
        conn.execute("""
            UPDATE users
            SET deleted_at=NULL, active=?, password_hash=?, evaluator_no=?
            WHERE id=?
        """, (active, generate_password_hash(password), evaluator_no, user_id))
        if role == "evaluator":
            rebalance_question_points(conn, user_id)
        log_event(conn, g.user["id"], "restore", "user", user_id,
                  f"بازیابی حساب حذف‌شده {row['username']} با حفظ سوابق")
        conn.commit()
        restored = conn.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
        return jsonify({"ok": True, "user": user_payload(restored)})
    except sqlite3.IntegrityError:
        conn.rollback()
        return json_error("بازیابی کاربر به دلیل تداخل اطلاعات انجام نشد.", 409)
    finally:
        conn.close()

@app.post("/api/users/<int:user_id>/photo")
@admin_required
def upload_user_photo(user_id: int) -> Any:
    uploaded = request.files.get("photo")
    if not uploaded or not uploaded.filename:
        return json_error("ابتدا تصویر کاربر را انتخاب کنید.")
    mime = (uploaded.mimetype or "").lower()
    allowed = {"image/png": ".png", "image/jpeg": ".jpg", "image/webp": ".webp"}
    if mime not in allowed:
        return json_error("فرمت تصویر باید PNG، JPG یا WEBP باشد.")
    data = uploaded.read(2 * 1024 * 1024 + 1)
    if not data or len(data) > 2 * 1024 * 1024:
        return json_error("حجم تصویر باید کمتر از ۲ مگابایت باشد.")
    # Verify basic file signatures rather than trusting the browser-provided MIME type.
    valid = ((mime == "image/png" and data.startswith(b"\x89PNG\r\n\x1a\n")) or
             (mime == "image/jpeg" and data.startswith(b"\xff\xd8\xff")) or
             (mime == "image/webp" and len(data) > 12 and data[:4] == b"RIFF" and data[8:12] == b"WEBP"))
    if not valid:
        return json_error("محتوای فایل با فرمت تصویر اعلام‌شده مطابقت ندارد.")
    conn = connect_db()
    try:
        if not conn.execute("SELECT 1 FROM users WHERE id=?", (user_id,)).fetchone():
            return json_error("کاربر پیدا نشد.", 404)
        conn.execute("UPDATE users SET profile_image=?, profile_image_mime=? WHERE id=?", (sqlite3.Binary(data), mime, user_id))
        log_event(conn, g.user["id"], "upload_photo", "user", user_id, "بارگذاری تصویر پروفایل")
        conn.commit()
        return jsonify({"ok": True, "message": "تصویر کاربر ذخیره شد."})
    finally:
        conn.close()


@app.get("/api/users/<int:user_id>/photo")
@login_required
def get_user_photo(user_id: int) -> Any:
    # Photos are visible only when the signed-in account is the photo's owner.
    if int(g.user["id"]) != int(user_id):
        return json_error("دسترسی به تصویر این کاربر مجاز نیست.", 403)
    conn = connect_db()
    try:
        row = conn.execute("SELECT profile_image, profile_image_mime FROM users WHERE id=?", (user_id,)).fetchone()
        if not row or not row["profile_image"]:
            return json_error("تصویری برای این کاربر ثبت نشده است.", 404)
        from flask import Response
        return Response(bytes(row["profile_image"]), mimetype=row["profile_image_mime"] or "image/png", headers={"Cache-Control": "no-store"})
    finally:
        conn.close()


@app.get("/api/questions")
@login_required
def get_questions() -> Any:
    conn = connect_db()
    try:
        where = "" if g.user["role"] == "admin" else "WHERE q.assigned_user_id=?"
        args: tuple[Any, ...] = () if g.user["role"] == "admin" else (g.user["id"],)
        rows = conn.execute(f"""
            SELECT q.*, u.display_name AS assigned_name, u.username AS assigned_username, u.max_points AS user_max_points
            FROM questions q LEFT JOIN users u ON u.id=q.assigned_user_id
            {where}
            ORDER BY q.active DESC, q.section_name COLLATE NOCASE, q.id DESC
        """, args).fetchall()
        return jsonify({"ok": True, "questions": [dict(r) for r in rows]})
    finally:
        conn.close()


def question_payload_valid(payload: dict[str, Any]) -> tuple[str, str, str, int, float] | tuple[None, None, None, None, None]:
    text = str(payload.get("question_text", "")).strip()
    rubric = str(payload.get("rubric", "")).strip()
    section = str(payload.get("section_name", "")).strip()
    try:
        user_id = int(payload.get("assigned_user_id"))
        max_score = float(payload.get("max_score"))
    except (ValueError, TypeError):
        return None, None, None, None, None
    return text, rubric, section, user_id, max_score


def rebalance_question_points(conn: sqlite3.Connection, user_id: int) -> None:
    """Split the evaluator quota equally; version scored questions to preserve history."""
    user = conn.execute("SELECT max_points FROM users WHERE id=? AND role='evaluator' AND active=1", (user_id,)).fetchone()
    if not user:
        return
    active = conn.execute("SELECT * FROM questions WHERE assigned_user_id=? AND active=1 ORDER BY id", (user_id,)).fetchall()
    for question in active:
        if conn.execute("SELECT 1 FROM scores WHERE question_id=? LIMIT 1", (question["id"],)).fetchone():
            conn.execute("UPDATE questions SET active=0,updated_at=datetime('now','localtime') WHERE id=?", (question["id"],))
            conn.execute("INSERT INTO questions(question_text,rubric,max_score,evaluator_no,assigned_user_id,section_name,version,active,created_at,updated_at) VALUES(?,?,?,?,?,?,?,1,datetime('now','localtime'),datetime('now','localtime'))",
                         (question["question_text"], question["rubric"], question["max_score"], question["evaluator_no"], user_id, question["section_name"], int(question["version"] or 1)+1))
    active = conn.execute("SELECT id FROM questions WHERE assigned_user_id=? AND active=1 ORDER BY id", (user_id,)).fetchall()
    if not active:
        return
    quota = float(user["max_points"] or 0)
    each = quota / len(active)
    allocated = 0.0
    for i, question in enumerate(active):
        points = quota - allocated if i == len(active)-1 else each
        allocated += points
        conn.execute("UPDATE questions SET max_score=?,updated_at=datetime('now','localtime') WHERE id=?", (points, question["id"]))


def validate_question_allocation(conn: sqlite3.Connection, user_id: int, max_score: float, exclude_id: int | None = None) -> str | None:
    user = conn.execute("SELECT * FROM users WHERE id=? AND role='evaluator' AND active=1", (user_id,)).fetchone()
    if not user:
        return "باید یک ارزیاب فعال را به سؤال اختصاص دهید."
    if max_score <= 0 or max_score > 100:
        return "بارم سؤال باید بیشتر از صفر و حداکثر ۱۰۰ باشد."
    sql = "SELECT COALESCE(SUM(max_score),0) FROM questions WHERE assigned_user_id=? AND active=1"
    args: list[Any] = [user_id]
    if exclude_id is not None:
        sql += " AND id<>?"
        args.append(exclude_id)
    used = float(conn.execute(sql, tuple(args)).fetchone()[0])
    if used + max_score > float(user["max_points"]) + 0.00001:
        return f"مجموع بارم سؤال‌های این کاربر از سقف {float(user['max_points']):g} امتیازی او بیشتر می‌شود. بارم آزاد: {max(0,float(user['max_points'])-used):g}."
    return None


@app.post("/api/questions")
@admin_required
def create_question() -> Any:
    payload = request.get_json(silent=True) or {}
    text, rubric, section, user_id, max_score = question_payload_valid(payload)
    if text is None:
        return json_error("متن سؤال، ارزیاب و بارم معتبر را وارد کنید.")
    if not text or len(text) > 1000 or len(rubric) > 4000 or len(section) > 120:
        return json_error("متن سؤال یا معیار نمره‌دهی خالی است یا بیش از حد طولانی است.")
    conn = connect_db()
    try:
        user = conn.execute("SELECT evaluator_no,section_name,max_points,active,role FROM users WHERE id=?", (user_id,)).fetchone()
        if not user or user["role"] != "evaluator" or not user["active"] or float(user["max_points"] or 0) <= 0:
            return json_error("یک ارزیاب فعال با سقف امتیاز مثبت انتخاب کنید.")
        section = section or user["section_name"] or "بخش ارزیابی"
        cur = conn.execute("INSERT INTO questions(question_text,rubric,max_score,evaluator_no,assigned_user_id,section_name,version,active) VALUES(?,?,?,?,?,?,1,1)",
                           (text, rubric, max_score, user["evaluator_no"], user_id, section))
        qid = cur.lastrowid
        rebalance_question_points(conn, user_id)
        active_period = get_active_period(conn)
        period_context = f"؛ دوره={active_period['name']}" if active_period else ""
        log_event(conn, g.user["id"], "create", "question", qid, f"ایجاد سؤال برای {user_id}: {text[:100]}{period_context}")
        conn.commit()
        return jsonify({"ok": True, "question_id": qid, "auto_allocation": True}), 201
    finally:
        conn.close()


@app.put("/api/questions/<int:question_id>")
@admin_required
def update_question(question_id: int) -> Any:
    payload = request.get_json(silent=True) or {}
    text, rubric, section, user_id, max_score = question_payload_valid(payload)
    if text is None or not text or len(text) > 1000 or len(rubric) > 4000 or len(section) > 120:
        return json_error("اطلاعات سؤال معتبر نیست.")
    conn = connect_db()
    try:
        old = conn.execute("SELECT * FROM questions WHERE id=?", (question_id,)).fetchone()
        if not old:
            return json_error("سؤال پیدا نشد.", 404)
        user = conn.execute("SELECT evaluator_no,section_name,max_points,active,role FROM users WHERE id=?", (user_id,)).fetchone()
        if not user or user["role"] != "evaluator" or not user["active"] or float(user["max_points"] or 0) <= 0:
            return json_error("یک ارزیاب فعال با سقف امتیاز مثبت انتخاب کنید.")
        section = section or user["section_name"] or "بخش ارزیابی"
        active_period = get_active_period(conn)
        has_active_period_scores = bool(active_period and conn.execute("SELECT 1 FROM scores WHERE question_id=? AND period_id=? LIMIT 1", (question_id, active_period["id"])).fetchone())
        if has_active_period_scores:
            return json_error("برای این سؤال در دوره فعال نمره ثبت شده است. برای حفظ دقت سوابق، ابتدا دوره قبلی را نهایی کنید و تغییر را در دوره بعد انجام دهید.", 409)
        has_scores = conn.execute("SELECT 1 FROM scores WHERE question_id=? LIMIT 1", (question_id,)).fetchone() is not None
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if has_scores:
            # Scored questions are versioned by replacing the current question; previous score snapshots remain intact.
            conn.execute("UPDATE questions SET active=0,updated_at=? WHERE id=?", (now, question_id))
            cur = conn.execute("INSERT INTO questions(question_text,rubric,max_score,evaluator_no,assigned_user_id,section_name,version,active,created_at,updated_at) VALUES(?,?,?,?,?,?,?,1,?,?)",
                               (text, rubric, max_score, user["evaluator_no"], user_id, section, int(old["version"] or 1)+1, now, now))
            new_id = cur.lastrowid
            detail = f"نسخه‌بندی سؤال {question_id} به {new_id} برای حفظ سوابق نمره"
        else:
            conn.execute("UPDATE questions SET question_text=?,rubric=?,max_score=?,evaluator_no=?,assigned_user_id=?,section_name=?,version=version+1,updated_at=? WHERE id=?",
                         (text, rubric, max_score, user["evaluator_no"], user_id, section, now, question_id))
            new_id = question_id
            detail = f"ویرایش سؤال {question_id}"
        rebalance_question_points(conn, user_id)
        active_period = get_active_period(conn)
        period_context = f"؛ دوره={active_period['name']}" if active_period else ""
        log_event(conn, g.user["id"], "update", "question", question_id, detail + period_context)
        conn.commit()
        return jsonify({"ok": True, "question_id": new_id, "versioned": has_scores, "auto_allocation": True})
    finally:
        conn.close()


@app.post("/api/questions/<int:question_id>/toggle")
@admin_required
def toggle_question(question_id: int) -> Any:
    conn = connect_db()
    try:
        q = conn.execute("SELECT * FROM questions WHERE id=?", (question_id,)).fetchone()
        if not q:
            return json_error("سؤال پیدا نشد.", 404)
        active = 0 if q["active"] else 1
        active_period = get_active_period(conn)
        if not active and active_period and conn.execute("SELECT 1 FROM scores WHERE question_id=? AND period_id=? LIMIT 1", (question_id, active_period["id"])).fetchone():
            return json_error("برای این سؤال در دوره فعال نمره ثبت شده است. ابتدا ارزیابی‌های جاری را تعیین تکلیف کنید.", 409)
        if active:
            owner = conn.execute("SELECT role,active,max_points FROM users WHERE id=?", (q["assigned_user_id"],)).fetchone()
            if not owner or owner["role"] != "evaluator" or not owner["active"] or float(owner["max_points"] or 0) <= 0:
                return json_error("برای فعال‌کردن سؤال، ارزیاب باید فعال و دارای سقف امتیاز مثبت باشد.")
        conn.execute("UPDATE questions SET active=?,updated_at=datetime('now','localtime') WHERE id=?", (active, question_id))
        rebalance_question_points(conn, int(q["assigned_user_id"]))
        active_period = get_active_period(conn)
        period_context = f"؛ دوره={active_period['name']}" if active_period else ""
        log_event(conn, g.user["id"], "activate" if active else "deactivate", "question", question_id, "تغییر وضعیت سؤال" + period_context)
        conn.commit()
        return jsonify({"ok": True, "active": bool(active)})
    finally:
        conn.close()


@app.get("/api/evaluation")
@login_required
def get_evaluation() -> Any:
    unit_id = request.args.get("unit_id", type=int)
    period_id = request.args.get("period_id", type=int)
    requested_user = request.args.get("evaluator_id")
    conn = connect_db()
    try:
        unit = conn.execute("SELECT * FROM units WHERE id=?", (unit_id,)).fetchone()
        period = conn.execute("SELECT * FROM evaluation_periods WHERE id=? AND is_archived=0", (period_id,)).fetchone() if period_id else get_active_period(conn)
        if not unit or not period:
            return json_error("زیررده یا دوره انتخاب‌شده پیدا نشد.", 404)
        target = get_target_evaluator(conn, requested_user)
        if g.user["role"] == "admin" and not target:
            return jsonify({"ok": True, "unit": dict(unit), "period": dict(period), "evaluators": [user_payload(r) for r in conn.execute("SELECT * FROM users WHERE role='evaluator' ORDER BY display_name").fetchall()], "needs_evaluator": True})
        if not target:
            return json_error("ارزیاب انتخاب‌شده معتبر نیست.", 400)
        submission = conn.execute("SELECT * FROM evaluation_submissions WHERE period_id=? AND unit_id=? AND user_id=?", (period["id"], unit["id"], target["id"])).fetchone()
        if g.user["role"] == "admin":
            questions = conn.execute("""
                SELECT q.* FROM questions q WHERE q.assigned_user_id=? AND (
                    q.active=1 OR EXISTS (SELECT 1 FROM scores sc WHERE sc.question_id=q.id AND sc.period_id=? AND sc.unit_id=? AND sc.evaluator_user_id=?)
                ) ORDER BY q.active DESC, q.id
            """, (target["id"], period["id"], unit["id"], target["id"])).fetchall()
        elif submission and submission["status"] == "submitted":
            # Show the exact question versions that were scored at final submission.
            # Active replacement questions must not appear as blank historical results.
            questions = conn.execute("""
                SELECT q.* FROM questions q
                JOIN scores sc ON sc.question_id=q.id
                WHERE q.assigned_user_id=? AND sc.period_id=? AND sc.unit_id=? AND sc.evaluator_user_id=?
                ORDER BY q.id
            """, (target["id"], period["id"], unit["id"], target["id"])).fetchall()
        else:
            questions = conn.execute("SELECT * FROM questions WHERE assigned_user_id=? AND active=1 ORDER BY id", (target["id"],)).fetchall()
        scores = conn.execute("SELECT * FROM scores WHERE period_id=? AND unit_id=? AND evaluator_user_id=?", (period["id"], unit["id"], target["id"])).fetchall()
        score_map = {r["question_id"]: r for r in scores if r["question_id"] is not None}
        required_sum = float(conn.execute(
            "SELECT COALESCE(SUM(max_score),0) FROM questions WHERE assigned_user_id=? AND active=1",
            (target["id"],)
        ).fetchone()[0])
        user_max = float(target["max_points"] or 0)
        active_evaluators = conn.execute("SELECT id,max_points FROM users WHERE role='evaluator' AND active=1 AND max_points>0").fetchall()
        global_allocation_total = float(conn.execute("SELECT COALESCE(SUM(max_points),0) FROM users WHERE role='evaluator' AND active=1").fetchone()[0])
        all_allocations_ready = abs(global_allocation_total-100) < 0.00001
        for evaluator_row in active_evaluators:
            question_points = float(conn.execute("SELECT COALESCE(SUM(max_score),0) FROM questions WHERE assigned_user_id=? AND active=1", (evaluator_row["id"],)).fetchone()[0])
            if abs(question_points-float(evaluator_row["max_points"])) > 0.00001:
                all_allocations_ready = False
                break
        lock = bool(submission and submission["status"] == "submitted" and g.user["role"] != "admin")
        inactive_period_lock = bool(not period["is_active"] and g.user["role"] != "admin")
        q_payload = []
        for q in questions:
            s = score_map.get(q["id"])
            q_payload.append({**dict(q), "score": float(s["score"]) if s else None, "comment": s["comment"] if s else ""})
        return jsonify({"ok": True, "unit": dict(unit), "period": dict(period), "evaluator": user_payload(target),
                        "questions": q_payload, "submission": dict(submission) if submission else None,
                        "locked": lock or inactive_period_lock, "is_admin": g.user["role"] == "admin",
                        "question_points": round(required_sum, 2), "user_max_points": user_max,
                        "allocation_ready": abs(required_sum-user_max) < 0.00001,
                        "global_allocation_total": global_allocation_total, "final_ready": all_allocations_ready})
    finally:
        conn.close()


@app.post("/api/evaluation/save")
@login_required
def save_evaluation() -> Any:
    payload = request.get_json(silent=True) or {}
    unit_id = payload.get("unit_id")
    period_id = payload.get("period_id")
    action = str(payload.get("action", "draft"))
    requested_user = payload.get("evaluator_id")
    try:
        unit_id, period_id = int(unit_id), int(period_id)
    except (ValueError, TypeError):
        return json_error("زیررده یا دوره معتبر نیست.")
    if action not in ("draft", "submit", "admin_update"):
        return json_error("نوع ثبت معتبر نیست.")
    if action == "admin_update" and g.user["role"] != "admin":
        return json_error("اصلاح نمره پس از ثبت نهایی فقط برای مدیر مجاز است.", 403)
    conn = connect_db()
    try:
        unit = conn.execute("SELECT * FROM units WHERE id=?", (unit_id,)).fetchone()
        period = conn.execute("SELECT * FROM evaluation_periods WHERE id=? AND is_archived=0", (period_id,)).fetchone()
        target = get_target_evaluator(conn, requested_user)
        if not unit or not period or not target:
            return json_error("زیررده، دوره یا ارزیاب معتبر نیست.", 404)
        if g.user["role"] != "admin" and not period["is_active"]:
            return json_error("ثبت نمره در دوره غیرفعال مجاز نیست.", 403)
        existing = conn.execute("SELECT * FROM evaluation_submissions WHERE period_id=? AND unit_id=? AND user_id=?", (period_id, unit_id, target["id"])).fetchone()
        if existing and existing["status"] == "submitted" and g.user["role"] != "admin":
            return json_error("این ارزیابی ثبت نهایی شده است؛ برای تغییر نمره با مدیر سامانه تماس بگیرید.", 403)
        questions = conn.execute("SELECT * FROM questions WHERE assigned_user_id=? AND active=1 ORDER BY id", (target["id"],)).fetchall()
        if not questions:
            return json_error("برای این کاربر هنوز سؤال فعالی تعریف نشده است.")
        editable_questions = list(questions)
        if g.user["role"] == "admin":
            scored_in_period = conn.execute("""
                SELECT q.* FROM questions q WHERE q.assigned_user_id=? AND q.active=0 AND EXISTS (
                    SELECT 1 FROM scores sc WHERE sc.question_id=q.id AND sc.period_id=? AND sc.unit_id=? AND sc.evaluator_user_id=?
                ) ORDER BY q.id
            """, (target["id"], period_id, unit_id, target["id"])).fetchall()
            editable_questions.extend(scored_in_period)
        points_total = sum(float(q["max_score"]) for q in questions)
        user_max = float(target["max_points"] or 0)
        allocation_ready = abs(points_total-user_max) < 0.00001
        if action == "submit" and not allocation_ready:
            return json_error(f"برای ثبت نهایی، مجموع بارم سؤال‌های فعال ({points_total:g}) باید دقیقاً برابر سقف نمره این کاربر ({user_max:g}) باشد.")
        if action == "submit":
            allocation_total = float(conn.execute("SELECT COALESCE(SUM(max_points),0) FROM users WHERE role='evaluator' AND active=1").fetchone()[0])
            if abs(allocation_total - 100) > 0.00001:
                return json_error(f"قبل از ثبت نهایی، مجموع سقف نمره ارزیاب‌های فعال باید دقیقاً ۱۰۰ باشد؛ مقدار فعلی {allocation_total:g} است.")
            all_evaluators = conn.execute("SELECT id,max_points FROM users WHERE role='evaluator' AND active=1 AND max_points>0").fetchall()
            for evaluator_row in all_evaluators:
                active_points = float(conn.execute("SELECT COALESCE(SUM(max_score),0) FROM questions WHERE assigned_user_id=? AND active=1", (evaluator_row["id"],)).fetchone()[0])
                if abs(active_points - float(evaluator_row["max_points"])) > 0.00001:
                    return json_error(f"جمع بارم سؤال‌های فعال ارزیاب «{evaluator_row['id']}» باید برابر سقف نمره اختصاص‌یافته او باشد؛ فعلاً {active_points:g} از {float(evaluator_row['max_points']):g} است.")
        by_id = {q["id"]: q for q in editable_questions}
        active_question_ids = {q["id"] for q in questions}
        scores_in = payload.get("scores", [])
        if not isinstance(scores_in, list):
            return json_error("ساختار نمره‌ها معتبر نیست.")
        seen: set[int] = set()
        values: list[tuple[int, float, str]] = []
        for item in scores_in:
            if not isinstance(item, dict):
                return json_error("اطلاعات یک نمره معتبر نیست.")
            try:
                qid = int(item.get("question_id"))
            except (TypeError, ValueError):
                return json_error("شناسه سؤال معتبر نیست.")
            q = by_id.get(qid)
            if not q or qid in seen:
                return json_error("یکی از سؤال‌ها به این کاربر اختصاص ندارد یا تکراری است.", 403)
            seen.add(qid)
            val = item.get("score")
            if val in (None, ""):
                # Inactive historical questions can be shown to administrators for reference;
                # they must not block a final submission for the current active question set.
                if action == "submit" and qid in active_question_ids:
                    return json_error(f"برای ثبت نهایی باید برای سؤال فعال «{q['question_text']}» نمره وارد کنید.")
                continue
            try:
                score_value = float(val)
            except (TypeError, ValueError):
                return json_error("نمره باید عددی باشد.")
            if score_value < 0 or score_value > float(q["max_score"]):
                return json_error(f"نمره سؤال «{q['question_text']}» باید بین صفر و {float(q['max_score']):g} باشد.")
            comment = str(item.get("comment", "")).strip()[:1000]
            values.append((qid, score_value, comment))
        if action == "submit":
            scored_ids = {qid for qid, _score, _comment in values}
            missing_ids = active_question_ids - scored_ids
            if missing_ids:
                missing_q = next((q for q in questions if q["id"] in missing_ids), None)
                label = missing_q["question_text"] if missing_q else "سؤال فعال"
                return json_error(f"برای ثبت نهایی باید برای سؤال فعال «{label}» نمره وارد کنید.")
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        evaluator_no = target["evaluator_no"]
        for qid, score_value, comment in values:
            q = by_id[qid]
            conn.execute("""
                INSERT INTO scores(period_id,unit_id,question_id,evaluator_user_id,evaluator_no,question_text_snapshot,max_score_snapshot,score,comment,updated_by,updated_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(period_id,unit_id,question_id) DO UPDATE SET
                    evaluator_user_id=excluded.evaluator_user_id, evaluator_no=excluded.evaluator_no,
                    score=excluded.score, comment=excluded.comment, updated_by=excluded.updated_by, updated_at=excluded.updated_at
            """, (period_id, unit_id, qid, target["id"], evaluator_no, q["question_text"], float(q["max_score"]), score_value, comment, g.user["id"], now))
        new_status = "submitted" if action == "submit" else (existing["status"] if existing and existing["status"] == "submitted" else "draft")
        submitted_at = now if action == "submit" else (existing["submitted_at"] if existing else None)
        if existing:
            conn.execute("UPDATE evaluation_submissions SET status=?,submitted_at=?,updated_at=?,last_modified_by=? WHERE id=?",
                         (new_status, submitted_at, now, g.user["id"], existing["id"]))
        else:
            conn.execute("INSERT INTO evaluation_submissions(period_id,unit_id,user_id,status,submitted_at,updated_at,last_modified_by) VALUES(?,?,?,?,?,?,?)",
                         (period_id, unit_id, target["id"], new_status, submitted_at, now, g.user["id"]))
        log_action = "final_submit" if action == "submit" else ("admin_score_override" if action == "admin_update" else "save_draft")
        log_event(conn, g.user["id"], log_action, "evaluation", unit_id, f"زیررده={unit['name']}؛ ارزیاب={target['username']}؛ دوره={period['name']}")
        conn.execute("UPDATE units SET updated_at=? WHERE id=?", (now, unit_id))
        conn.commit()
        return jsonify({"ok": True, "status": new_status, "submitted_at": submitted_at, "message": "ثبت نهایی انجام شد." if new_status == "submitted" else "ذخیره موقت انجام شد."})
    except sqlite3.IntegrityError as exc:
        conn.rollback()
        return json_error("ذخیره نمره انجام نشد؛ یکپارچگی داده‌ها را بررسی کنید.", 409)
    finally:
        conn.close()


@app.get("/api/dashboard")
@login_required
def get_dashboard() -> Any:
    period_id = request.args.get("period_id", type=int)
    conn = connect_db()
    try:
        period = conn.execute("SELECT * FROM evaluation_periods WHERE id=? AND is_archived=0", (period_id,)).fetchone() if period_id else get_active_period(conn)
        if period is None:
            return jsonify({"ok": True, "monthly": [], "progress": [], "period": None})
        # Last 12 Jalali calendar months: show average final score for units completed in each month.
        monthly: list[dict[str, Any]] = []
        now = datetime.now()
        current_jy, current_jm, _ = gregorian_to_jalali(now.year, now.month, now.day)
        month_starts: list[tuple[int, int, str, str]] = []
        for offset in range(11, -1, -1):
            serial_month = current_jy * 12 + (current_jm - 1) - offset
            jy, month0 = divmod(serial_month, 12)
            jm = month0 + 1
            next_serial_month = serial_month + 1
            next_jy, next_month0 = divmod(next_serial_month, 12)
            next_jm = next_month0 + 1
            gsy, gsm, gsd = jalali_to_gregorian(jy, jm, 1)
            gey, gem, ged = jalali_to_gregorian(next_jy, next_jm, 1)
            start_text = f"{gsy:04d}-{gsm:02d}-{gsd:02d} 00:00:00"
            end_text = f"{gey:04d}-{gem:02d}-{ged:02d} 00:00:00"
            month_starts.append((jy, jm, start_text, end_text))
        required_ids = active_evaluator_ids(conn) if g.user["role"] == "admin" else [int(g.user["id"])]
        for jy, jm, start, end in month_starts:
            completed_unit_scores: list[float] = []
            periods = [{"id": period["id"]}]
            for p in periods:
                units = conn.execute("SELECT id FROM units").fetchall()
                for un in units:
                    if not required_ids:
                        continue
                    subs = conn.execute("SELECT user_id,submitted_at FROM evaluation_submissions WHERE period_id=? AND unit_id=? AND status='submitted'", (p["id"], un["id"])).fetchall()
                    sub_ids = {sub["user_id"] for sub in subs}
                    if not set(required_ids).issubset(sub_ids):
                        continue
                    completion_date = max((sub["submitted_at"] or "" for sub in subs if sub["user_id"] in required_ids), default="")
                    if not (start <= completion_date < end):
                        continue
                    value = conn.execute(
                        "SELECT COALESCE(SUM(sc.score),0) FROM scores sc "
                        "WHERE sc.period_id=? AND sc.unit_id=? AND sc.evaluator_user_id IN ({})".format(",".join("?" for _ in required_ids)),
                        (p["id"], un["id"], *required_ids),
                    ).fetchone()[0]
                    completed_unit_scores.append(float(value))
            monthly.append({"month": f"{jy:04d}-{jm:02d}", "score": round(sum(completed_unit_scores)/len(completed_unit_scores), 2) if completed_unit_scores else 0,
                            "count": len(completed_unit_scores)})
        progress = []
        if g.user["role"] == "admin":
            evaluator_rows = conn.execute("SELECT * FROM users WHERE role='evaluator' AND active=1 ORDER BY display_name").fetchall()
            total_units = conn.execute("SELECT COUNT(*) FROM units").fetchone()[0]
            for evaluator in evaluator_rows:
                q_count = conn.execute("SELECT COUNT(*) FROM questions WHERE assigned_user_id=? AND active=1", (evaluator["id"],)).fetchone()[0]
                if not q_count:
                    continue
                done = conn.execute("SELECT COUNT(*) FROM evaluation_submissions WHERE period_id=? AND user_id=? AND status='submitted'", (period["id"], evaluator["id"])).fetchone()[0]
                progress.append({"user_id": evaluator["id"], "name": evaluator["display_name"], "done": done, "total": total_units,
                                 "percent": round(done/total_units*100) if total_units else 0,
                                 "question_points": float(conn.execute("SELECT COALESCE(SUM(max_score),0) FROM questions WHERE assigned_user_id=? AND active=1", (evaluator["id"],)).fetchone()[0]),
                                 "max_points": float(evaluator["max_points"] or 0)})
        ranked_units = []
        for unit_row in conn.execute("SELECT * FROM units").fetchall():
            summary = unit_summary(conn, unit_row, period)
            if summary.get("status") == "done" and summary.get("score") is not None:
                ranked_units.append({"id": int(unit_row["id"]), "name": unit_row["name"], "score": float(summary["score"])})
        ranked_units.sort(key=lambda item: (-item["score"], item["name"]))
        if g.user["role"] == "admin":
            required_count = len(active_evaluator_ids(conn))
            detail_rows = conn.execute("""
                SELECT un.id,un.name,COALESCE(SUM(CASE WHEN sub.status='submitted' THEN sc.score ELSE 0 END),0) score,
                       COALESCE(SUM(CASE WHEN sub.status='submitted' THEN sc.max_score_snapshot ELSE 0 END),0) maximum,
                       COUNT(DISTINCT CASE WHEN sub.status='submitted' THEN sc.evaluator_user_id END) submitted_count
                FROM units un LEFT JOIN scores sc ON sc.unit_id=un.id AND sc.period_id=?
                LEFT JOIN evaluation_submissions sub ON sub.period_id=sc.period_id AND sub.unit_id=sc.unit_id AND sub.user_id=sc.evaluator_user_id
                GROUP BY un.id ORDER BY un.name
            """, (period["id"],)).fetchall()
            performance = [{"id":r["id"],"name":r["name"],"score":round(float(r["score"]),2),
                            "percent":round(float(r["score"])/float(r["maximum"])*100,2) if float(r["maximum"] or 0)>0 and int(r["submitted_count"])==required_count else None}
                           for r in detail_rows]
        else:
            detail_rows = conn.execute("""
                SELECT un.id,un.name,sub.status,COALESCE(SUM(sc.score),0) score
                FROM units un LEFT JOIN evaluation_submissions sub ON sub.unit_id=un.id AND sub.period_id=? AND sub.user_id=?
                LEFT JOIN scores sc ON sc.unit_id=un.id AND sc.period_id=? AND sc.evaluator_user_id=?
                GROUP BY un.id ORDER BY un.name
            """, (period["id"],g.user["id"],period["id"],g.user["id"])).fetchall()
            quota=float(g.user["max_points"] or 0)
            performance=[{"id":r["id"],"name":r["name"],"score":round(float(r["score"]),2),
                          "percent":round(float(r["score"])/quota*100,2) if quota>0 and r["status"]=="submitted" else None}
                         for r in detail_rows]
        section_performance = []
        if g.user["role"] == "admin":
            section_rows = conn.execute("""
                SELECT COALESCE(NULLIF(q.section_name,''),NULLIF(ev.section_name,''),'بدون بخش') AS section_label,
                       SUM(sc.score) score,SUM(sc.max_score_snapshot) maximum
                FROM scores sc
                JOIN evaluation_submissions sub ON sub.period_id=sc.period_id AND sub.unit_id=sc.unit_id AND sub.user_id=sc.evaluator_user_id AND sub.status='submitted'
                LEFT JOIN questions q ON q.id=sc.question_id
                LEFT JOIN users ev ON ev.id=sc.evaluator_user_id
                WHERE sc.period_id=?
                GROUP BY 1 ORDER BY 1
            """, (period["id"],)).fetchall()
            section_performance=[{"name":r["section_label"],"percent":round(float(r["score"])/float(r["maximum"])*100,2) if float(r["maximum"] or 0)>0 else None} for r in section_rows]
        assessed=[x["percent"] for x in performance if x["percent"] is not None]
        distribution={"very_low":sum(0<=x<=20 for x in assessed),"low":sum(21<=x<=40 for x in assessed),
                      "medium":sum(41<=x<=80 for x in assessed),"good":sum(81<=x<=100 for x in assessed)}
        stats={"assessed":len(assessed),"pending":len(performance)-len(assessed),
               "average":round(sum(assessed)/len(assessed),2) if assessed else None,
               "highest":max(assessed) if assessed else None,"lowest":min(assessed) if assessed else None}
        return jsonify({"ok": True, "period": dict(period), "monthly": monthly, "progress": progress,
                        "ranked_units": ranked_units if g.user["role"]=="admin" else [],
                        "performance":performance,"section_performance":section_performance,"distribution":distribution,"stats":stats,
                        "is_admin": g.user["role"] == "admin"})
    finally:
        conn.close()


def month_cutoff(months: int) -> str:
    # Rolling interval based on Jalali calendar months, ending at the same local clock time.
    now = datetime.now()
    jy, jm, jd = gregorian_to_jalali(now.year, now.month, now.day)
    serial_month = jy * 12 + (jm - 1) - months
    target_jy, target_month0 = divmod(serial_month, 12)
    target_jm = target_month0 + 1
    target_jd = min(jd, get_jalali_month_days(target_jy, target_jm))
    gy, gm, gd = jalali_to_gregorian(target_jy, target_jm, target_jd)
    return now.replace(year=gy, month=gm, day=gd).strftime("%Y-%m-%d %H:%M:%S")



@app.get("/api/reports/overview.xlsx")
@primary_admin_required
def export_overview_report() -> Any:
    """Export a polished RTL executive dashboard and its supporting detail table."""
    period_id = request.args.get("period_id", type=int)
    conn = connect_db()
    try:
        period = conn.execute(
            "SELECT * FROM evaluation_periods WHERE id=?", (period_id,)
        ).fetchone() if period_id else get_active_period(conn)
        if not period:
            return json_error("دوره ارزیابی انتخاب‌شده پیدا نشد.", 404)

        evaluators = conn.execute("""
            SELECT id, display_name, max_points
            FROM users
            WHERE role='evaluator' AND active=1 AND max_points>0
            ORDER BY display_name
        """).fetchall()
        units = conn.execute("SELECT id, code, name FROM units ORDER BY name COLLATE NOCASE").fetchall()
        if not evaluators:
            return json_error("برای تولید گزارش، ارزیاب فعال تعریف نشده است.")
        if not units:
            return json_error("برای تولید گزارش، هیچ واحدی تعریف نشده است.")

        evaluator_ids = {int(r["id"]) for r in evaluators}
        completed_units = []
        for unit in units:
            submitted = {
                int(r[0]) for r in conn.execute("""
                    SELECT user_id FROM evaluation_submissions
                    WHERE unit_id=? AND period_id=? AND status='submitted'
                """, (unit["id"], period["id"]))
            }
            if evaluator_ids.issubset(submitted):
                completed_units.append(unit)
        if not completed_units:
            return json_error("هنوز هیچ واحدی ارزیابی خود را به‌طور کامل ثبت نکرده است.")

        evaluator_count = len(evaluators)
        maximum_points = sum(float(r["max_points"] or 0) for r in evaluators)
        if maximum_points <= 0:
            return json_error("سقف امتیاز ارزیابان معتبر نیست.")

        report_rows = []
        for unit in completed_units:
            score = float(conn.execute("""
                SELECT COALESCE(SUM(score),0)
                FROM scores WHERE period_id=? AND unit_id=?
            """, (period["id"], unit["id"])).fetchone()[0] or 0)
            percent = (score / maximum_points * 100) if maximum_points else 0
            section_rows = conn.execute("""
                SELECT COALESCE(NULLIF(q.section_name,''), NULLIF(ev.section_name,''),'بدون بخش') AS section_name,
                       COALESCE(SUM(sc.score),0) AS score
                FROM scores sc
                LEFT JOIN questions q ON q.id=sc.question_id
                LEFT JOIN users ev ON ev.id=sc.evaluator_user_id
                WHERE sc.period_id=? AND sc.unit_id=?
                GROUP BY COALESCE(NULLIF(q.section_name,''), NULLIF(ev.section_name,''),'بدون بخش')
                ORDER BY section_name COLLATE NOCASE
            """, (period["id"], unit["id"])).fetchall()
            report_rows.append({
                "code": unit["code"], "name": unit["name"],
                "score": score, "percent": percent,
                "sections": "؛ ".join(
                    f"{r['section_name']}: {float(r['score']):g}" for r in section_rows
                ) or "بدون نمره"
            })

        report_rows.sort(key=lambda r: (-r["score"], str(r["name"])))
        for rank, row in enumerate(report_rows, start=1):
            row["rank"] = rank

        assessed_count = len(report_rows)
        average_score = sum(r["score"] for r in report_rows) / assessed_count
        average_percent = sum(r["percent"] for r in report_rows) / assessed_count

        wb = Workbook()
        ws = wb.active
        ws.title = "خلاصه مدیریتی"
        ws.sheet_view.rightToLeft = True
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A8"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.page_margins.left = 0.25
        ws.page_margins.right = 0.25
        ws.page_margins.top = 0.45
        ws.page_margins.bottom = 0.45
        ws.print_area = "A1:N35"

        navy = "173650"
        navy2 = "244B68"
        teal = "168C8C"
        gold = "D5A642"
        pale = "EAF1F5"
        pale_teal = "E4F3F1"
        white = "FFFFFF"
        ink = "243447"
        muted = "687987"
        line = "D7E0E6"
        green = "DFF0E5"
        yellow = "FFF2CC"
        red = "FCE4D6"

        thin_border = Border(
            left=Side(style="thin", color=line), right=Side(style="thin", color=line),
            top=Side(style="thin", color=line), bottom=Side(style="thin", color=line)
        )
        title_font = Font(name="B Titr", size=18, bold=True, color=white)
        subtitle_font = Font(name="B Nazanin", size=11, color=white)
        header_font = Font(name="B Titr", size=11, bold=True, color=white)
        body_font = Font(name="B Nazanin", size=12, color=ink)
        small_font = Font(name="B Nazanin", size=10, color=muted)
        metric_font = Font(name="B Titr", size=20, bold=True, color=navy)

        for col in range(1, 15):
            ws.column_dimensions[get_column_letter(col)].width = 13
        ws.column_dimensions["A"].width = 5
        ws.column_dimensions["B"].width = 22
        ws.column_dimensions["C"].width = 13
        ws.column_dimensions["D"].width = 13
        ws.column_dimensions["E"].width = 13
        ws.column_dimensions["F"].width = 13
        ws.column_dimensions["G"].width = 3
        ws.column_dimensions["H"].width = 13
        ws.column_dimensions["I"].width = 13
        ws.column_dimensions["J"].width = 13
        ws.column_dimensions["K"].width = 13
        ws.column_dimensions["L"].width = 13
        ws.column_dimensions["M"].width = 13
        ws.column_dimensions["N"].width = 13

        ws.merge_cells("A1:N2")
        ws["A1"] = "داشبورد مدیریتی ارزیابی عملکرد"
        ws["A1"].font = title_font
        ws["A1"].fill = PatternFill("solid", fgColor=navy)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        for row in ws["A1:N2"]:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=navy)
        ws.row_dimensions[1].height = 27
        ws.row_dimensions[2].height = 22

        ws.merge_cells("A3:N3")
        ws["A3"] = f"مرکز منابع انسانی سرباز سپاه بیت‌المقدس کردستان  |  دوره: {period['name']}"
        ws["A3"].font = subtitle_font
        ws["A3"].fill = PatternFill("solid", fgColor=navy2)
        ws["A3"].alignment = Alignment(horizontal="center", vertical="center")
        for cell in ws[3]:
            cell.fill = PatternFill("solid", fgColor=navy2)
        ws.row_dimensions[3].height = 25

        ws.merge_cells("A4:N4")
        ws["A4"] = (
            f"تاریخ تهیه گزارش: {jalali_date_text(datetime.now())}   |   "
            f"سقف امتیاز هر واحد: {maximum_points:g}   |   "
            "مبنای درصد تحقق: نمره کسب‌شده ÷ سقف امتیاز × ۱۰۰"
        )
        ws["A4"].font = small_font
        ws["A4"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[4].height = 23

        # Four KPI cards.
        cards = [
            ("A6:C6", "A7:C9", "تعداد ارزیابان فعال", evaluator_count, navy),
            ("D6:F6", "D7:F9", "واحدهای ارزیابی‌شده", assessed_count, teal),
            ("H6:J6", "H7:J9", "میانگین امتیاز", round(average_score, 2), navy2),
            ("K6:N6", "K7:N9", "میانگین درصد تحقق", round(average_percent, 2) / 100, gold),
        ]
        for label_range, value_range, label, value, color in cards:
            ws.merge_cells(label_range)
            ws.merge_cells(value_range)
            lc = ws[label_range.split(":")[0]]
            vc = ws[value_range.split(":")[0]]
            lc.value = label
            lc.font = header_font
            lc.fill = PatternFill("solid", fgColor=color)
            lc.alignment = Alignment(horizontal="center", vertical="center")
            vc.value = value
            vc.font = metric_font
            vc.fill = PatternFill("solid", fgColor=pale_teal if color == teal else pale)
            vc.alignment = Alignment(horizontal="center", vertical="center")
            if label == "میانگین درصد تحقق":
                vc.number_format = "0.00%"
            for rng in (label_range, value_range):
                for row_cells in ws[rng]:
                    for cell in row_cells:
                        cell.border = thin_border
            for r in range(6, 10):
                ws.row_dimensions[r].height = 22

        ws.merge_cells("A11:F11")
        ws["A11"] = "رتبه‌بندی واحدها بر اساس امتیاز کسب‌شده"
        ws["A11"].font = Font(name="B Titr", size=13, bold=True, color=white)
        ws["A11"].fill = PatternFill("solid", fgColor=navy)
        ws["A11"].alignment = Alignment(horizontal="center", vertical="center")
        for row_cells in ws["A11:F11"]:
            for cell in row_cells:
                cell.fill = PatternFill("solid", fgColor=navy)
        headers = ["رتبه", "نام واحد", "کد واحد", "امتیاز", "درصد تحقق", "وضعیت"]
        for col, value in enumerate(headers, start=1):
            c = ws.cell(12, col, value)
            c.font = header_font
            c.fill = PatternFill("solid", fgColor=navy2)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = thin_border
        for ridx, row in enumerate(report_rows, start=13):
            values = [
                row["rank"], row["name"], row["code"], round(row["score"], 2),
                row["percent"] / 100,
                "مطلوب" if row["percent"] < 50 else ("متوسط" if row["percent"] < 75 else "مطلوب/خوب")
            ]
            for col, value in enumerate(values, start=1):
                c = ws.cell(ridx, col, value)
                c.font = body_font
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                c.border = thin_border
                c.fill = PatternFill("solid", fgColor=white if ridx % 2 else pale)
            ws.cell(ridx, 5).number_format = "0.00%"
            status_cell = ws.cell(ridx, 6)
            status_cell.fill = PatternFill("solid", fgColor=red if row["percent"] < 50 else (yellow if row["percent"] < 75 else green))
            ws.row_dimensions[ridx].height = 24

        # Comparison chart beside the ranking table.
        from openpyxl.chart import BarChart, Reference
        chart = BarChart()
        chart.type = "bar"
        chart.style = 10
        chart.title = "مقایسه امتیاز واحدها"
        chart.y_axis.title = "واحد ارزیابی"
        chart.x_axis.title = "امتیاز"
        chart.height = 9.2
        chart.width = 18
        chart.legend = None
        chart.varyColors = True
        data_end = 12 + len(report_rows)
        chart.add_data(Reference(ws, min_col=4, min_row=12, max_row=data_end), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=2, min_row=13, max_row=data_end))
        chart.x_axis.scaling.min = 0
        chart.x_axis.scaling.max = maximum_points
        chart.dataLabels = None
        ws.add_chart(chart, "H11")

        section_start = max(15, 14 + len(report_rows))
        ws.merge_cells(start_row=section_start, start_column=1, end_row=section_start, end_column=14)
        note = ws.cell(section_start, 1)
        note.value = "راهنمای گزارش و ملاحظات محاسباتی"
        note.font = Font(name="B Titr", size=12, bold=True, color=white)
        note.fill = PatternFill("solid", fgColor=navy)
        note.alignment = Alignment(horizontal="center", vertical="center")
        for row_cells in ws.iter_rows(min_row=section_start, max_row=section_start, min_col=1, max_col=14):
            for cell in row_cells:
                cell.fill = PatternFill("solid", fgColor=navy)
        notes = [
            f"تعداد ارزیابان فعال دارای سقف امتیاز: {evaluator_count}",
            f"تعداد واحدهایی که تمام ارزیابان برای آن‌ها ثبت نهایی انجام داده‌اند: {assessed_count}",
            f"سقف امتیاز هر واحد برابر مجموع سقف امتیاز ارزیابان فعال است: {maximum_points:g}",
            "درصد تحقق هر واحد از تقسیم مجموع نمرات ثبت‌شده بر سقف امتیاز محاسبه شده است.",
            "رتبه‌بندی فقط بین واحدهای دارای ارزیابی کامل در همین دوره انجام شده است."
        ]
        for offset, text_value in enumerate(notes, start=1):
            r = section_start + offset
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
            c = ws.cell(r, 1, "• " + text_value)
            c.font = small_font
            c.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
            c.fill = PatternFill("solid", fgColor="F7F9FB" if offset % 2 else white)
            ws.row_dimensions[r].height = 21

        # Detailed source sheet for auditability and filtering.
        detail = wb.create_sheet("جزئیات واحدها")
        detail.sheet_view.rightToLeft = True
        detail.sheet_view.showGridLines = False
        detail.freeze_panes = "A5"
        detail.auto_filter.ref = f"A4:G{4 + len(report_rows)}"
        detail.merge_cells("A1:G2")
        detail["A1"] = f"جزئیات عملکرد واحدها | {period['name']}"
        detail["A1"].font = title_font
        detail["A1"].fill = PatternFill("solid", fgColor=navy)
        detail["A1"].alignment = Alignment(horizontal="center", vertical="center")
        for row_cells in detail["A1:G2"]:
            for cell in row_cells:
                cell.fill = PatternFill("solid", fgColor=navy)
        detail.merge_cells("A3:G3")
        detail["A3"] = f"سقف امتیاز: {maximum_points:g}  |  تعداد ارزیابان فعال: {evaluator_count}"
        detail["A3"].font = small_font
        detail["A3"].alignment = Alignment(horizontal="center", vertical="center")
        detail_headers = ["رتبه", "نام واحد", "کد واحد", "امتیاز کسب‌شده", "سقف امتیاز", "درصد تحقق", "تفکیک امتیاز بخش‌ها"]
        for col, value in enumerate(detail_headers, start=1):
            c = detail.cell(4, col, value)
            c.font = header_font
            c.fill = PatternFill("solid", fgColor=navy2)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = thin_border
        for ridx, row in enumerate(report_rows, start=5):
            values = [row["rank"], row["name"], row["code"], round(row["score"], 2),
                      maximum_points, row["percent"] / 100, row["sections"]]
            for col, value in enumerate(values, start=1):
                c = detail.cell(ridx, col, value)
                c.font = body_font
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                c.border = thin_border
                c.fill = PatternFill("solid", fgColor=white if ridx % 2 else pale)
            detail.cell(ridx, 6).number_format = "0.00%"
            detail.row_dimensions[ridx].height = 36
        for col, width in {"A":9, "B":25, "C":15, "D":18, "E":16, "F":16, "G":70}.items():
            detail.column_dimensions[col].width = width
        detail.sheet_properties.pageSetUpPr.fitToPage = True
        detail.page_setup.orientation = "landscape"
        detail.page_setup.paperSize = detail.PAPERSIZE_A4
        detail.page_setup.fitToWidth = 1
        detail.page_setup.fitToHeight = 0
        detail.print_title_rows = "1:4"

        # Full traceable detail requested for the primary administrator only.
        all_detail = wb.create_sheet("جزئیات ارزیابی تمام بخش‌ها")
        all_detail.sheet_view.rightToLeft = True
        all_detail.freeze_panes = "A2"
        all_headers = ["دوره ارزیابی","بخش","زیررده","نام و نام خانوادگی ارزیاب","متن سؤال","سقف امتیاز سؤال","نمره ثبت‌شده","وضعیت ارزیابی","تاریخ ثبت نهایی"]
        for col, label in enumerate(all_headers, 1):
            c=all_detail.cell(1,col,label); c.font=header_font; c.fill=PatternFill("solid",fgColor=navy2)
            c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        rows = conn.execute("""
            SELECT COALESCE(NULLIF(q.section_name,''),NULLIF(ev.section_name,''),'بدون بخش') section_name,
                   un.name unit_name,ev.display_name evaluator_name,
                   COALESCE(sc.question_text_snapshot,q.question_text) question_text,
                   sc.max_score_snapshot,sc.score,sub.status,sub.submitted_at
            FROM scores sc
            JOIN units un ON un.id=sc.unit_id
            LEFT JOIN questions q ON q.id=sc.question_id
            LEFT JOIN users ev ON ev.id=sc.evaluator_user_id
            LEFT JOIN evaluation_submissions sub ON sub.period_id=sc.period_id AND sub.unit_id=sc.unit_id AND sub.user_id=sc.evaluator_user_id
            WHERE sc.period_id=? ORDER BY section_name,unit_name,evaluator_name,q.id
        """,(period["id"],)).fetchall()
        for ridx,item in enumerate(rows,2):
            vals=[period["name"],item["section_name"],item["unit_name"],item["evaluator_name"] or "—",
                  item["question_text"],item["max_score_snapshot"],item["score"],
                  "ثبت نهایی" if item["status"]=="submitted" else ("موقت" if item["status"]=="draft" else "ناقص"),
                  jalali_date_text(item["submitted_at"],True)]
            for col,value in enumerate(vals,1):
                c=all_detail.cell(ridx,col,value); c.font=body_font
                c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=thin_border
        all_detail.auto_filter.ref=f"A1:I{max(1,len(rows)+1)}"
        for col,width in enumerate([18,24,22,25,55,16,15,16,20],1): all_detail.column_dimensions[get_column_letter(col)].width=width
        all_detail.sheet_properties.pageSetUpPr.fitToPage=True
        all_detail.page_setup.orientation="landscape"; all_detail.page_setup.fitToWidth=1; all_detail.page_setup.fitToHeight=0

        # Workbook metadata and consistent print settings.
        wb.properties.title = f"داشبورد مدیریتی ارزیابی عملکرد - {period['name']}"
        wb.properties.subject = "خلاصه مدیریتی، مقایسه و رتبه‌بندی عملکرد واحدها"
        wb.properties.creator = "سامانه ارزیابی عملکرد"
        wb.active = 0

        watermark_path = BASE_DIR / "static" / "assets" / "2511_watermark.png"
        if watermark_path.exists():
            try:
                from openpyxl.drawing.image import Image as XLImage
                for sheet in wb.worksheets:
                    wm = XLImage(str(watermark_path)); wm.width = 150; wm.height = 230
                    sheet.add_image(wm, "H8")
            except Exception:
                pass
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        log_event(conn, g.user["id"], "export", "overview_report", None,
                  f"داشبورد مدیریتی دوره {period['name']}")
        conn.commit()
        safe = lambda v: " ".join("".join(c for c in str(v) if c not in '<>:/\\|?*"').split())
        filename = f"داشبورد مدیریتی - {safe(period['name'])}.xlsx"
        return send_file(out, as_attachment=True, download_name=filename,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    finally:
        conn.close()


@app.get("/api/units/<int:unit_id>/report.xlsx")
@login_required
def export_unit_report(unit_id: int) -> Any:
    """Create the official unit report for a selected evaluation period."""
    period_id = request.args.get("period_id", type=int)
    conn = connect_db()
    try:
        unit = conn.execute("SELECT * FROM units WHERE id=?", (unit_id,)).fetchone()
        if not unit:
            return json_error("زیررده پیدا نشد.", 404)
        period = conn.execute("SELECT * FROM evaluation_periods WHERE id=?", (period_id,)).fetchone() if period_id else get_active_period(conn)
        if not period:
            return json_error("دوره ارزیابی انتخاب‌شده پیدا نشد.", 404)
        required = conn.execute("SELECT id,display_name FROM users WHERE role='evaluator' AND active=1 AND max_points>0 ORDER BY display_name").fetchall()
        if not required:
            return json_error("هیچ ارزیاب فعالی برای این دوره تعریف نشده است.")
        submitted = {int(r[0]) for r in conn.execute("SELECT user_id FROM evaluation_submissions WHERE unit_id=? AND period_id=? AND status='submitted'", (unit_id, period["id"]))}
        section_rows = conn.execute("""
            SELECT COALESCE(NULLIF(q.section_name,''), NULLIF(ev.section_name,''), 'بدون بخش') AS section_name,
                   COALESCE(SUM(sc.score),0) AS section_score
            FROM scores sc
            LEFT JOIN questions q ON q.id=sc.question_id
            LEFT JOIN users ev ON ev.id=sc.evaluator_user_id
            WHERE sc.unit_id=? AND sc.period_id=?
              AND EXISTS (SELECT 1 FROM evaluation_submissions sx WHERE sx.period_id=sc.period_id AND sx.unit_id=sc.unit_id AND sx.user_id=sc.evaluator_user_id AND sx.status='submitted')
            GROUP BY COALESCE(NULLIF(q.section_name,''), NULLIF(ev.section_name,''), 'بدون بخش')
            ORDER BY section_name COLLATE NOCASE
        """, (unit_id, period["id"])).fetchall()
        if g.user["role"] != "admin":
            own_section = (g.user["section_name"] or "").strip()
            section_rows = [r for r in section_rows if str(r["section_name"]).strip() == own_section]
        total_score = float(conn.execute("SELECT COALESCE(SUM(sc.score),0) FROM scores sc WHERE sc.unit_id=? AND sc.period_id=? AND EXISTS (SELECT 1 FROM evaluation_submissions sx WHERE sx.period_id=sc.period_id AND sx.unit_id=sc.unit_id AND sx.user_id=sc.evaluator_user_id AND sx.status='submitted')", (unit_id, period["id"])).fetchone()[0])
        # Rank among sub-units whose assessments for this same period are complete.
        all_units = conn.execute("SELECT id,name FROM units ORDER BY name COLLATE NOCASE").fetchall()
        rank_values = []
        for other in all_units:
            other_submitted = {int(r[0]) for r in conn.execute("SELECT user_id FROM evaluation_submissions WHERE unit_id=? AND period_id=? AND status='submitted'", (other["id"], period["id"]))}
            if set(int(r["id"]) for r in required).issubset(other_submitted):
                score = float(conn.execute("SELECT COALESCE(SUM(score),0) FROM scores WHERE unit_id=? AND period_id=?", (other["id"], period["id"])).fetchone()[0])
                rank_values.append((int(other["id"]), score))
        rank_values.sort(key=lambda x: x[1], reverse=True)
        rank = next((i for i, (uid, _) in enumerate(rank_values, start=1) if uid == unit_id), None)
        total_ranked = len(rank_values)

        detailed_rows = conn.execute("""
            SELECT ? AS period_name, COALESCE(NULLIF(q.section_name,''),NULLIF(ev.section_name,''),'بدون بخش') AS section_name,
                   ? AS unit_name, COALESCE(sc.question_text_snapshot,q.question_text) AS question_text,
                   sc.max_score_snapshot AS max_score, sc.score, ev.display_name AS evaluator_name,
                   sub.status AS submission_status, sub.submitted_at
            FROM scores sc
            LEFT JOIN questions q ON q.id=sc.question_id
            LEFT JOIN users ev ON ev.id=sc.evaluator_user_id
            LEFT JOIN evaluation_submissions sub ON sub.period_id=sc.period_id AND sub.unit_id=sc.unit_id AND sub.user_id=sc.evaluator_user_id
            WHERE sc.period_id=? AND sc.unit_id=?
            ORDER BY evaluator_name, q.id, sc.id
        """, (period["name"], unit["name"], period["id"], unit_id)).fetchall()
        if g.user["role"] != "admin":
            detailed_rows = [r for r in detailed_rows if str(r["section_name"]).strip() == (g.user["section_name"] or "").strip()]

        template = BASE_DIR / "static" / "assets" / "گزارش عملکرد دوره رده.xlsx"
        if template.exists():
            from openpyxl import load_workbook
            wb = load_workbook(template)
            ws = wb.active
        else:
            wb = Workbook(); ws = wb.active
        ws.sheet_view.rightToLeft = True
        ws.cell(1, 1).value = f"گزارش ارزیابی عملکرد رده {unit['name']} در {period['name']} - مرکز منابع انسانی سرباز سپاه بیت‌المقدس کردستان"
        ws.cell(1, 1).font = Font(name="B Titr", bold=True, size=14, color="173650")
        ws.cell(1, 1).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 42
        try:
            from openpyxl.drawing.image import Image as XLImage
            logo_path = BASE_DIR / "static" / "assets" / "2511.png"
            if logo_path.exists():
                logo = XLImage(str(logo_path)); logo.width = 54; logo.height = 54
                ws.add_image(logo, "N1")
        except Exception:
            pass
        ws.cell(2, 1).value = "در بخش"
        ws.cell(2, 2).value = "نمره کسب شده"
        # Clear sample rows, retaining the supplied template's first two header rows and styles.
        for row in range(3, max(ws.max_row, 3) + 1):
            for col in range(1, max(ws.max_column, 2) + 1):
                ws.cell(row, col).value = None
        # The template's row 3 carries the desired table styling; clone it for additional sections.
        from copy import copy
        base_style_row = 3
        for idx, item in enumerate(section_rows, start=3):
            if idx > 3:
                for col in range(1, max(ws.max_column, 2) + 1):
                    source, target = ws.cell(base_style_row, col), ws.cell(idx, col)
                    if source.has_style:
                        target._style = copy(source._style)
                    if source.number_format:
                        target.number_format = source.number_format
                    target.alignment = copy(source.alignment)
            ws.cell(idx, 1).value = str(item["section_name"])
            ws.cell(idx, 2).value = round(float(item["section_score"]), 2)
            ws.cell(idx, 1).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.cell(idx, 2).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(idx, 1).font = Font(name="B Nazanin", size=13, color="243447")
            ws.cell(idx, 2).font = Font(name="B Nazanin", size=13, color="243447")
        data_end = max(3, 2 + len(section_rows))
        # Put the provincial ranking in a dedicated visible cell beside the chart.
        ws["D2"] = "رتبه این رده در سطح سپاه استان"
        ws["D2"].font = Font(name="B Titr", bold=True, size=11, color="173650")
        ws["D2"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws["E2"] = f"{rank} از {total_ranked}" if rank else "قابل محاسبه نیست"
        ws["E2"].font = Font(name="B Titr", bold=True, size=13, color="173650")
        ws["E2"].alignment = Alignment(horizontal="center", vertical="center")
        if section_rows:
            from openpyxl.chart import BarChart, Reference
            chart = BarChart()
            chart.type = "col"
            chart.style = 10
            chart.title = f"نمره بخش‌های زیررده {unit['name']}"
            chart.y_axis.title = "نمره"
            chart.x_axis.title = "نام بخش"
            chart.height = 8.5
            chart.width = 17
            data = Reference(ws, min_col=2, min_row=2, max_row=data_end)
            cats = Reference(ws, min_col=1, min_row=3, max_row=data_end)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            chart.legend = None
            chart.x_axis.tickLblPos = "low"
            chart.x_axis.txPr = None
            ws.add_chart(chart, "G3")
        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 19
        ws.column_dimensions["D"].width = 25
        ws.column_dimensions["E"].width = 17
        ws.freeze_panes = "A3"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None and cell.font.name not in ("B Titr", "B Nazanin"):
                    cell.font = copy(cell.font)
                    cell.font = Font(name="B Nazanin", size=cell.font.sz or 11, bold=cell.font.bold, italic=cell.font.italic, color=cell.font.color)
        detail = wb.create_sheet("جزئیات ارزیابی ارزیابان")
        detail.sheet_view.rightToLeft = True
        detail.freeze_panes = "A2"
        headers = ["دوره ارزیابی","بخش","زیررده","نام و نام خانوادگی ارزیاب","متن سؤال","سقف امتیاز سؤال","نمره ثبت‌شده","وضعیت ارزیابی","تاریخ ثبت نهایی"]
        for col, label in enumerate(headers, 1):
            c = detail.cell(1, col, label); c.font = Font(name="B Titr", size=10, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="244B68"); c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for ridx, item in enumerate(detailed_rows, 2):
            vals = [item["period_name"], item["section_name"], item["unit_name"], item["evaluator_name"] or "—",
                    item["question_text"], item["max_score"], item["score"],
                    "ثبت نهایی" if item["submission_status"]=="submitted" else ("موقت" if item["submission_status"]=="draft" else "ناقص"),
                    jalali_date_text(item["submitted_at"], True)]
            for col, value in enumerate(vals, 1):
                c=detail.cell(ridx,col,value); c.font=Font(name="B Nazanin",size=11)
                c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        detail.auto_filter.ref=f"A1:I{max(1,len(detailed_rows)+1)}"
        detail.sheet_properties.pageSetUpPr.fitToPage=True
        detail.page_setup.orientation="landscape"; detail.page_setup.fitToWidth=1; detail.page_setup.fitToHeight=0
        for col,width in enumerate([18,24,22,25,55,16,15,16,20],1): detail.column_dimensions[get_column_letter(col)].width=width
        watermark_path = BASE_DIR / "static" / "assets" / "2511_watermark.png"
        if watermark_path.exists():
            try:
                from openpyxl.drawing.image import Image as XLImage
                wm=XLImage(str(watermark_path)); wm.width=180; wm.height=277
                ws.add_image(wm,"F10"); detail.add_image(XLImage(str(watermark_path)),"F10")
            except Exception:
                pass
        output = io.BytesIO(); wb.save(output); output.seek(0)
        log_event(conn, g.user["id"], "export", "unit_report", unit_id, f"گزارش دوره {period['name']} برای {unit['name']}")
        conn.commit()
        safe = lambda value: " ".join("".join(c for c in str(value) if c not in '<>:/\\|?*\"') .split())
        filename = f"گزارش عملکرد {safe(period['name'])} {safe(unit['name'])}.xlsx"
        return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    finally:
        conn.close()

@app.get("/api/health")
def health() -> Any:
    return jsonify({"ok": True, "service": "organizational-evaluation", "database": "sqlite", "version": BUILD_VERSION})


init_db()

if __name__ == "__main__":
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "5000"))
    print("=" * 56)
    print("سامانه ارزیابی عملکرد — نسخه داخلی")
    print(f"Local URL: http://127.0.0.1:{port}")
    print("Initial admin: admin / admin123")
    print("Initial evaluators: evaluator1..evaluator7 / eval123")
    print("Change initial passwords before operational use. Ctrl+C to stop.")
    print("=" * 56)
    app.run(host=host, port=port, debug=False)
