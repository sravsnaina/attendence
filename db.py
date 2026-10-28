"""SQLite storage. Everything the app knows lives in one file: data/attendance.db
(plus data/uploads, data/reports, data/backups). Copy the `data` folder to move or archive it.
"""
from __future__ import annotations

import json
import os
import shutil
import sqlite3
from datetime import datetime
from pathlib import Path

from engine import merge_settings

DATA_DIR = Path(os.environ.get("ATT_DATA_DIR", Path(__file__).parent / "data"))

SCHEMA = """
CREATE TABLE IF NOT EXISTS employees(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT UNIQUE NOT NULL,
  designation TEXT DEFAULT '',
  aliases TEXT DEFAULT '[]',
  emp_type TEXT DEFAULT 'staff',
  active INTEGER DEFAULT 1,
  sort_order INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS daily(
  emp_id INTEGER NOT NULL, date TEXT NOT NULL,
  code TEXT DEFAULT '', in_time TEXT DEFAULT '', out_time TEXT DEFAULT '',
  perm_hours REAL DEFAULT 0, remarks TEXT DEFAULT '',
  source TEXT DEFAULT 'manual', updated_at TEXT,
  PRIMARY KEY(emp_id, date));
CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS imports(
  id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT, sha256 TEXT, stored_path TEXT,
  kind TEXT, rows INTEGER, months TEXT, imported_at TEXT);
CREATE TABLE IF NOT EXISTS audit(
  id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, action TEXT, details TEXT);
CREATE TABLE IF NOT EXISTS reports(
  id INTEGER PRIMARY KEY AUTOINCREMENT, month TEXT, created_at TEXT, file_path TEXT,
  summary_json TEXT, finalized INTEGER DEFAULT 0);
"""

# Seeded from your standard sheet ("Attendance List For Every Month.xlsx") and the October PDF.
SEED = [
    ("Dharanidharan", "Software & Controls Engineer", ["DHARANIDHARAN"], "staff"),
    ("Prem Sundar S", "Junior Robotics Engineer", ["PREM SUNDAR", "PREM SUNDAR S"], "staff"),
    ("Gopal P", "Mechanical Project Lead", ["GOPAL"], "staff"),
    ("Surya S", "Robotics Mechanical Engineer", ["SURYA"], "staff"),
    ("Sarvani Lakshmi M", "Junior Software Engineer", ["SRAVANI", "SARVANI"], "staff"),
    ("Bharath K", "Graduate Engineer Trainee", ["BHARATH"], "staff"),
    ("Bharathi K", "Supply Chain Lead", ["BHARATHI"], "staff"),
    ("Thanigai Malai A", "Intern", ["THANIGAI MALAI A"], "intern"),
    ("Sam R", "Intern", ["SAM R"], "intern"),
    ("Ranjith Kumar S", "Intern", ["RANJITH KUMAR S"], "intern"),
    ("Shalini M", "Intern", ["SHALINI M"], "intern"),
]


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def configure(data_dir):
    global DATA_DIR
    DATA_DIR = Path(data_dir)


def path(*parts):
    p = DATA_DIR.joinpath(*parts)
    return p


def conn():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    c = sqlite3.connect(DATA_DIR / "attendance.db")
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    return c


def init_db():
    for sub in ("uploads", "reports", "backups", "tmp"):
        (DATA_DIR / sub).mkdir(parents=True, exist_ok=True)
    with conn() as c:
        c.executescript(SCHEMA)
        if c.execute("SELECT COUNT(*) FROM employees").fetchone()[0] == 0:
            for i, (n, d, al, t) in enumerate(SEED, start=1):
                c.execute("INSERT INTO employees(name,designation,aliases,emp_type,sort_order) VALUES(?,?,?,?,?)",
                          (n, d, json.dumps(al), t, i))


def audit(c, action, details):
    c.execute("INSERT INTO audit(ts,action,details) VALUES(?,?,?)",
              (now(), action, json.dumps(details, default=str)))


# ---------------------------------------------------------------- settings
def get_settings():
    with conn() as c:
        saved = {r["key"]: json.loads(r["value"]) for r in c.execute("SELECT key,value FROM settings")}
    return merge_settings(saved)


def save_settings(new):
    with conn() as c:
        for k, v in new.items():
            c.execute("INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                      (k, json.dumps(v)))
        audit(c, "settings", new)


# --------------------------------------------------------------- employees
def _emp(r):
    d = dict(r)
    d["aliases"] = json.loads(d.get("aliases") or "[]")
    return d


def list_employees(active_only=True):
    with conn() as c:
        q = "SELECT * FROM employees" + (" WHERE active=1" if active_only else "") + " ORDER BY sort_order, id"
        return [_emp(r) for r in c.execute(q)]


def get_employee(emp_id):
    with conn() as c:
        r = c.execute("SELECT * FROM employees WHERE id=?", (emp_id,)).fetchone()
    return _emp(r) if r else None


def add_employee(name, designation="", aliases=(), emp_type="staff", c=None):
    own = c is None
    c = c or conn()
    try:
        order = c.execute("SELECT COALESCE(MAX(sort_order),0)+1 FROM employees").fetchone()[0]
        cur = c.execute("INSERT INTO employees(name,designation,aliases,emp_type,sort_order) VALUES(?,?,?,?,?)",
                        (name.strip(), designation.strip(), json.dumps(list(aliases)), emp_type, order))
        audit(c, "employee_add", {"name": name})
        if own:
            c.commit()
        return cur.lastrowid
    except Exception:
        if own:
            c.rollback()
        raise
    finally:
        if own:
            c.close()


def update_employee(emp_id, name, designation, aliases, emp_type, active, sort_order):
    with conn() as c:
        c.execute("UPDATE employees SET name=?,designation=?,aliases=?,emp_type=?,active=?,sort_order=? WHERE id=?",
                  (name.strip(), designation.strip(), json.dumps(list(aliases)), emp_type, int(active), int(sort_order), emp_id))
        audit(c, "employee_update", {"id": emp_id, "name": name})


def add_alias(emp_id, alias, c=None):
    own = c is None
    c = c or conn()
    try:
        r = c.execute("SELECT * FROM employees WHERE id=?", (emp_id,)).fetchone()
        alias = alias.strip()
        if r and alias:
            e = _emp(r)
            if alias.lower() not in {a.lower() for a in e["aliases"]} and alias.lower() != e["name"].lower():
                c.execute("UPDATE employees SET aliases=? WHERE id=?", (json.dumps(e["aliases"] + [alias]), emp_id))
        if own:
            c.commit()
    finally:
        if own:
            c.close()


# ------------------------------------------------------------- day records
FIELDS = ("code", "in_time", "out_time", "perm_hours", "remarks")


def save_daily(c, emp_id, iso, rec, source, protect_manual=False):
    """Insert/update one day. Returns inserted | updated | unchanged | skipped_manual."""
    new = {"code": (rec.get("code") or "").strip(), "in_time": rec.get("in_time") or "",
           "out_time": rec.get("out_time") or "", "perm_hours": float(rec.get("perm_hours") or 0),
           "remarks": (rec.get("remarks") or "").strip()}
    old = c.execute("SELECT * FROM daily WHERE emp_id=? AND date=?", (emp_id, iso)).fetchone()
    if old is None:
        c.execute("INSERT INTO daily(emp_id,date,code,in_time,out_time,perm_hours,remarks,source,updated_at) VALUES(?,?,?,?,?,?,?,?,?)",
                  (emp_id, iso, new["code"], new["in_time"], new["out_time"], new["perm_hours"], new["remarks"], source, now()))
        return "inserted"
    if all(old[f] == new[f] for f in FIELDS):
        if old["source"] == "ocr" and source == "manual":   # saving the row = verified by you
            c.execute("UPDATE daily SET source='manual', updated_at=? WHERE emp_id=? AND date=?", (now(), emp_id, iso))
            return "updated"
        return "unchanged"
    if protect_manual and old["source"] == "manual" and source != "manual":
        return "skipped_manual"
    c.execute("UPDATE daily SET code=?,in_time=?,out_time=?,perm_hours=?,remarks=?,source=?,updated_at=? WHERE emp_id=? AND date=?",
              (new["code"], new["in_time"], new["out_time"], new["perm_hours"], new["remarks"], source, now(), emp_id, iso))
    audit(c, "daily_change", {"emp_id": emp_id, "date": iso, "from": {f: old[f] for f in FIELDS}, "to": new, "source": source})
    return "updated"


def get_records(year, month, emp_id=None):
    prefix = f"{year:04d}-{month:02d}-%"
    q, args = "SELECT * FROM daily WHERE date LIKE ?", [prefix]
    if emp_id is not None:
        q += " AND emp_id=?"
        args.append(emp_id)
    with conn() as c:
        return {(r["emp_id"], r["date"]): dict(r) for r in c.execute(q, args)}


def months_with_data():
    with conn() as c:
        return [r[0] for r in c.execute("SELECT DISTINCT substr(date,1,7) m FROM daily ORDER BY m DESC")]


# ------------------------------------------------------------------ others
def log_import(c, filename, sha, stored, kind, rows, months):
    c.execute("INSERT INTO imports(filename,sha256,stored_path,kind,rows,months,imported_at) VALUES(?,?,?,?,?,?,?)",
              (filename, sha, str(stored), kind, rows, ",".join(sorted(months)), now()))


def backup(keep=40):
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = DATA_DIR / "backups" / f"attendance_{stamp}.db"
    src = sqlite3.connect(DATA_DIR / "attendance.db")
    dst = sqlite3.connect(dest)
    with dst:
        src.backup(dst)
    src.close()
    dst.close()
    olds = sorted((DATA_DIR / "backups").glob("attendance_*.db"))
    for o in olds[:-keep]:
        o.unlink()
    return dest


def save_report(month, file_path, summary, finalized=False):
    with conn() as c:
        c.execute("INSERT INTO reports(month,created_at,file_path,summary_json,finalized) VALUES(?,?,?,?,?)",
                  (month, now(), str(file_path), json.dumps(summary, default=str), int(finalized)))


def list_reports(limit=15):
    with conn() as c:
        return [dict(r) for r in c.execute("SELECT * FROM reports ORDER BY id DESC LIMIT ?", (limit,))]


def all_daily_rows():
    with conn() as c:
        return [dict(r) for r in c.execute(
            "SELECT e.name, d.date, d.code, d.in_time, d.out_time, d.perm_hours, d.remarks, d.source, d.updated_at "
            "FROM daily d JOIN employees e ON e.id=d.emp_id ORDER BY d.date, e.sort_order")]
