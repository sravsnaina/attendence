"""Attendance Manager - local web app.   Run:  python app.py   then open http://127.0.0.1:5000

Import the fingerprint PDF (or Excel/CSV/photo), fix what needs a decision, download the
monthly Excel for the CEO. Everything is stored in data/attendance.db.
"""
from __future__ import annotations

import calendar
import csv
import hashlib
import io
import json
import threading
import uuid
import webbrowser
from collections import Counter
from datetime import date, datetime, timedelta

from flask import (Flask, Response, abort, flash, redirect, render_template, request,
                   send_file, url_for)
from werkzeug.utils import secure_filename

import db
import engine
import importers
import report

app = Flask(__name__)
app.secret_key = "local-attendance-manager"
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024


# ------------------------------------------------------------- helpers
def parse_month(arg):
    try:
        y, m = arg.split("-")
        y, m = int(y), int(m)
        if 1 <= m <= 12 and 2000 <= y <= 2100:
            return y, m
    except (ValueError, AttributeError):
        pass
    return None


def current_month():
    got = parse_month(request.args.get("m") or request.form.get("m"))
    if got:
        return got
    have = db.months_with_data()
    if have:
        return parse_month(have[0])
    t = date.today()
    return t.year, t.month


def shift_month(y, m, delta):
    n = y * 12 + (m - 1) + delta
    return n // 12, n % 12 + 1


def compute(y, m):
    emps = db.list_employees()
    st = db.get_settings()
    recs = db.get_records(y, m)
    return engine.compute_month(y, m, emps, recs, st, today=date.today()), st


@app.template_filter("days")
def days_filter(v):
    if not v:
        return "-"
    return f"{v:g} Day" if v == 1 else f"{v:g} Days"


@app.template_filter("hhmm")
def hhmm_filter(v):
    return engine.fmt_min(v) if v else "-"


@app.context_processor
def inject():
    return {"month_name": calendar.month_name, "STATUS_LABEL": engine.STATUS_LABEL,
            "ENTRY_CODES": engine.ENTRY_CODES, "fmt_min": engine.fmt_min}


def mkey(y, m):
    return f"{y:04d}-{m:02d}"


# ----------------------------------------------------------- dashboard
@app.route("/")
def dashboard():
    y, m = current_month()
    results, st = compute(y, m)
    totals = Counter()
    for r in results:
        s = r["summary"]
        totals["errors"] += s["errors"]
        totals["warnings"] += s["warnings"]
        totals["pending"] += s["pending"]
    py, pm = shift_month(y, m, -1)
    ny, nm = shift_month(y, m, 1)
    return render_template("dashboard.html", results=results, st=st, y=y, m=m, mk=mkey(y, m),
                           prev=mkey(py, pm), next=mkey(ny, nm), totals=totals,
                           months=db.months_with_data(), reports=db.list_reports(), today_iso=date.today().isoformat())


def _summary_rows(results):
    return [{"employee": r["emp"]["name"], **{k: v for k, v in r["summary"].items()}} for r in results]


def _build(y, m, finalized=False):
    results, st = compute(y, m)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = db.path("reports", f"Attendance_{mkey(y, m)}_{stamp}.xlsx")
    report.build_workbook(results, y, m, st, path)
    db.save_report(mkey(y, m), path, _summary_rows(results), finalized)
    return path, results


@app.route("/report.xlsx")
def download_report():
    y, m = current_month()
    path, _ = _build(y, m)
    return send_file(path, as_attachment=True,
                     download_name=f"Attendance_{calendar.month_name[m]}_{y}.xlsx")


@app.route("/finalize", methods=["POST"])
def finalize():
    y, m = current_month()
    db.backup()
    path, results = _build(y, m, finalized=True)
    flash(f"Snapshot saved for {calendar.month_name[m]} {y}: {path.name} (and a database backup).", "ok")
    return redirect(url_for("dashboard", m=mkey(y, m)))


@app.route("/report/<int:rid>")
def old_report(rid):
    with db.conn() as c:
        r = c.execute("SELECT * FROM reports WHERE id=?", (rid,)).fetchone()
    if not r or not db.path("reports", r["file_path"].split("/")[-1].split("\\")[-1]).exists():
        abort(404)
    return send_file(db.path("reports", r["file_path"].split("/")[-1].split("\\")[-1]), as_attachment=True)


# -------------------------------------------------------------- import
@app.route("/import", methods=["GET", "POST"])
def import_page():
    if request.method == "GET":
        return render_template("import.html")
    f = request.files.get("file")
    if not f or not f.filename:
        flash("Choose a file first.", "err")
        return redirect(url_for("import_page"))
    raw = f.read()
    name = secure_filename(f.filename) or "upload"
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
    token = uuid.uuid4().hex[:12]
    stored = db.path("uploads", f"{datetime.now():%Y%m%d_%H%M%S}_{name}")
    stored.write_bytes(raw)
    try:
        if ext == "pdf":
            blocks, kind = importers.parse_pdf(io.BytesIO(raw)), "pdf"
        elif ext == "csv":
            blocks, kind = importers.parse_csv(io.BytesIO(raw)), "csv"
        elif ext in ("xlsx", "xlsm"):
            blocks, kind = importers.parse_xlsx(io.BytesIO(raw)), "xlsx"
        elif ext in ("png", "jpg", "jpeg", "bmp", "tif", "tiff", "webp"):
            blocks, kind = importers.parse_image(stored), "ocr"
        else:
            raise importers.ImportProblem("Unsupported file type. Use PDF, Excel (.xlsx), CSV or a photo (png/jpg).")
    except importers.ImportProblem as e:
        flash(str(e), "err")
        return redirect(url_for("import_page"))
    except Exception as e:  # corrupt file etc.
        flash(f"Could not read the file: {e}", "err")
        return redirect(url_for("import_page"))
    payload = {"filename": name, "stored": str(stored), "kind": kind,
               "sha": hashlib.sha256(raw).hexdigest(), "blocks": blocks}
    db.path("tmp", f"{token}.json").write_text(json.dumps(payload))
    return redirect(url_for("import_preview", token=token))


def _load_payload(token):
    p = db.path("tmp", f"{secure_filename(token)}.json")
    if not p.exists():
        abort(404)
    return json.loads(p.read_text())


@app.route("/import/<token>")
def import_preview(token):
    payload = _load_payload(token)
    emps = db.list_employees(active_only=False)
    previews = []
    with db.conn() as c:
        for i, b in enumerate(payload["blocks"]):
            emp, how = importers.match_employee(b["name"], emps)
            counts = Counter()
            if emp:
                for r in b["rows"]:
                    if not (r["code"] or r["in_time"] or r["out_time"]):
                        counts["blank"] += 1
                        continue
                    old = c.execute("SELECT * FROM daily WHERE emp_id=? AND date=?", (emp["id"], r["date"])).fetchone()
                    if old is None:
                        counts["new"] += 1
                    elif all(str(old[k] or "") == str(r.get(k) or "") for k in ("code", "in_time", "out_time")):
                        counts["same"] += 1
                    elif old["source"] == "manual":
                        counts["manual"] += 1
                    else:
                        counts["changed"] += 1
            problems = [f'{r["date"]}: {p}' for r in b["rows"] for p in r.get("problems", [])]
            dates = sorted(r["date"] for r in b["rows"])
            previews.append({"i": i, "b": b, "emp": emp, "how": how, "counts": counts,
                             "problems": problems, "range": (dates[0], dates[-1]) if dates else ("", "")})
    return render_template("preview.html", token=token, payload=payload, previews=previews, emps=emps)


@app.route("/import/<token>/confirm", methods=["POST"])
def import_confirm(token):
    payload = _load_payload(token)
    overwrite = bool(request.form.get("overwrite_manual"))
    source = payload["kind"]
    totals, months = Counter(), Counter()
    db.backup()
    with db.conn() as c:
        for i, b in enumerate(payload["blocks"]):
            choice = request.form.get(f"map_{i}", "skip")
            if choice == "skip":
                continue
            if choice == "new":
                try:
                    emp_id = db.add_employee(b["name"].title() if b["name"].isupper() else b["name"],
                                             b["designation"].title(), [b["name"]],
                                             "intern" if b["designation"].lower() == "intern" else "staff", c=c)
                except Exception:
                    c.rollback()
                    flash(f"Employee '{b['name']}' already exists - choose it from the list. Nothing was saved.", "err")
                    return redirect(url_for("import_preview", token=token))
            else:
                emp_id = int(choice)
                if b["name"]:
                    db.add_alias(emp_id, b["name"], c=c)
            for r in b["rows"]:
                if not (r["code"] or r["in_time"] or r["out_time"] or r.get("perm_hours")):
                    continue  # a blank cell in the file never erases data already saved
                res = db.save_daily(c, emp_id, r["date"], r, source, protect_manual=not overwrite)
                totals[res] += 1
                months[r["date"][:7]] += 1
        db.log_import(c, payload["filename"], payload["sha"], payload["stored"], source,
                      sum(totals.values()), months.keys())
        db.audit(c, "import", {"file": payload["filename"], **totals})
    db.path("tmp", f"{secure_filename(token)}.json").unlink(missing_ok=True)
    msg = (f"Imported {payload['filename']}: {totals['inserted']} new, {totals['updated']} updated, "
           f"{totals['unchanged']} unchanged")
    if totals["skipped_manual"]:
        msg += f", {totals['skipped_manual']} skipped (you edited them manually)"
    flash(msg + ".", "ok")
    target = months.most_common(1)[0][0] if months else None
    return redirect(url_for("dashboard", m=target) if target else url_for("dashboard"))


@app.route("/template.csv")
def template():
    return Response(importers.template_csv(), mimetype="text/csv",
                    headers={"Content-Disposition": "attachment; filename=timesheet_template.csv"})


# ---------------------------------------------------------- daily edit

def _save_form_row(c, emp_id, iso, code, tin_raw, tout_raw, perm_raw, remarks):
    """Validate + save one day typed in a form. Returns (changed: bool, error: str|None)."""
    try:
        tin, tout = engine.norm_time(tin_raw), engine.norm_time(tout_raw)
        perm = float(perm_raw or 0)
        if perm < 0:
            raise ValueError("negative")
    except ValueError:
        return False, f"{iso}: invalid time or permission hours - row not saved."
    rec = {"code": code, "in_time": tin, "out_time": tout, "perm_hours": perm, "remarks": remarks}
    exists = c.execute("SELECT 1 FROM daily WHERE emp_id=? AND date=?", (emp_id, iso)).fetchone()
    if not exists and not (code or tin or tout or perm or remarks):
        return False, None
    return db.save_daily(c, emp_id, iso, rec, "manual") in ("inserted", "updated"), None

@app.route("/edit/<int:emp_id>", methods=["GET", "POST"])
def edit(emp_id):
    emp = db.get_employee(emp_id)
    if not emp:
        abort(404)
    y, m = current_month()
    ndays = calendar.monthrange(y, m)[1]
    if request.method == "POST":
        saved = errors = 0
        with db.conn() as c:
            for i in range(ndays):
                changed, err = _save_form_row(
                    c, emp_id, date(y, m, i + 1).isoformat(), request.form.get(f"code_{i}", ""),
                    request.form.get(f"in_{i}", ""), request.form.get(f"out_{i}", ""),
                    request.form.get(f"perm_{i}"), request.form.get(f"rem_{i}", ""))
                if err:
                    flash(err, "err")
                    errors += 1
                elif changed:
                    saved += 1
        db.backup(keep=40)
        flash(f"Saved {saved} changed day(s) for {emp['name']}." + (f" {errors} row(s) had errors." if errors else ""),
              "ok" if not errors else "err")
        return redirect(url_for("edit", emp_id=emp_id, m=mkey(y, m)))
    st = db.get_settings()
    recs = {iso: r for (eid, iso), r in db.get_records(y, m, emp_id).items()}
    res = engine.compute_employee(emp, y, m, recs, st, today=date.today())
    py, pm = shift_month(y, m, -1)
    ny, nm = shift_month(y, m, 1)
    return render_template("edit.html", emp=emp, res=res, recs=recs, y=y, m=m, mk=mkey(y, m), st=st,
                           prev=mkey(py, pm), next=mkey(ny, nm), employees=db.list_employees())


@app.route("/day/<iso>", methods=["GET", "POST"])
def day_entry(iso):
    """Type one day for all employees (fast way to add the latest fingerprint-machine data)."""
    try:
        d = datetime.strptime(iso, "%Y-%m-%d").date()
    except ValueError:
        abort(404)
    emps = db.list_employees()
    if request.method == "POST":
        saved = errors = 0
        with db.conn() as c:
            for e in emps:
                k = e["id"]
                changed, err = _save_form_row(c, k, iso, request.form.get(f"code_{k}", ""),
                                              request.form.get(f"in_{k}", ""), request.form.get(f"out_{k}", ""),
                                              request.form.get(f"perm_{k}"), request.form.get(f"rem_{k}", ""))
                if err:
                    flash(f"{e['name']}: {err}", "err")
                    errors += 1
                elif changed:
                    saved += 1
        db.backup(keep=40)
        flash(f"Saved {saved} change(s) for {d:%d %b %Y}." + (f" {errors} row(s) had errors." if errors else ""),
              "ok" if not errors else "err")
        return redirect(url_for("day_entry", iso=iso))
    st = db.get_settings()
    recs = db.get_records(d.year, d.month)
    rows = []
    for e in emps:
        rec = recs.get((e["id"], iso), {})
        rows.append({"emp": e, "rec": rec, "res": engine.compute_day(d, rec, st, date.today())})
    return render_template("day.html", d=d, iso=iso, rows=rows, st=st, mk=mkey(d.year, d.month),
                           prev=(d - timedelta(days=1)).isoformat(), next=(d + timedelta(days=1)).isoformat())


@app.route("/quick", methods=["POST"])
def quick_leave():
    """Apply one status (PL, LOP, OD, WFH, ...) to a date range for one or more people."""
    y, m = current_month()
    ids = request.form.getlist("emp_ids")
    code = request.form.get("code", "")
    try:
        d1 = datetime.strptime(request.form["from"], "%Y-%m-%d").date()
        d2 = datetime.strptime(request.form["to"], "%Y-%m-%d").date()
    except (KeyError, ValueError):
        flash("Enter a valid from / to date.", "err")
        return redirect(url_for("dashboard", m=mkey(y, m)))
    if not ids or code not in engine.ENTRY_CODES or d2 < d1 or (d2 - d1).days > 62:
        flash("Pick at least one employee, a valid status and a sensible date range.", "err")
        return redirect(url_for("dashboard", m=mkey(y, m)))
    st = db.get_settings()
    remarks = request.form.get("remarks", "").strip()
    done = skipped = 0
    with db.conn() as c:
        for eid in ids:
            d = d1
            while d <= d2:
                iso = d.isoformat()
                old = c.execute("SELECT * FROM daily WHERE emp_id=? AND date=?", (int(eid), iso)).fetchone()
                old_code = engine.normalize_code(old["code"])[0] if old else ""
                cal_off = (not old_code) and (iso in st["holidays"] or d.weekday() in st["weekly_off_days"])
                if old_code in ("WO", "H") or cal_off:
                    skipped += 1
                else:
                    rec = {"code": code, "in_time": old["in_time"] if old else "", "out_time": old["out_time"] if old else "",
                           "perm_hours": old["perm_hours"] if old else 0, "remarks": remarks or (old["remarks"] if old else "")}
                    db.save_daily(c, int(eid), iso, rec, "manual")
                    done += 1
                d += timedelta(days=1)
    flash(f"Marked {done} day(s) as {code}" + (f"; skipped {skipped} week-off/holiday day(s)." if skipped else "."), "ok")
    return redirect(url_for("dashboard", m=mkey(d1.year, d1.month)))


# -------------------------------------------------------------- review
@app.route("/review")
def review():
    y, m = current_month()
    results, st = compute(y, m)
    items = []
    for r in results:
        for dr in r["days"]:
            for lvl, key, msg in dr["flags"]:
                if lvl != "info":
                    items.append({"emp": r["emp"], "date": dr["date"], "level": lvl, "key": key, "msg": msg,
                                  "action": report.ACTION.get(key, "")})
    order = {"error": 0, "warn": 1}
    items.sort(key=lambda x: (order[x["level"]], x["emp"]["id"], x["date"]))
    py, pm = shift_month(y, m, -1)
    ny, nm = shift_month(y, m, 1)
    return render_template("review.html", items=items, y=y, m=m, mk=mkey(y, m), prev=mkey(py, pm), next=mkey(ny, nm))


# ------------------------------------------------------------ employees
@app.route("/employees", methods=["GET", "POST"])
def employees():
    if request.method == "POST":
        f = request.form
        if f.get("action") == "add":
            if not f.get("name", "").strip():
                flash("Name is required.", "err")
            else:
                try:
                    db.add_employee(f["name"], f.get("designation", ""),
                                    [a.strip() for a in f.get("aliases", "").split(",") if a.strip()], f.get("emp_type", "staff"))
                    flash("Employee added.", "ok")
                except Exception:
                    flash("That name already exists.", "err")
        else:
            try:
                db.update_employee(int(f["id"]), f["name"], f.get("designation", ""),
                                   [a.strip() for a in f.get("aliases", "").split(",") if a.strip()],
                                   f.get("emp_type", "staff"), bool(f.get("active")), int(f.get("sort_order") or 0))
                flash("Employee updated.", "ok")
            except Exception as e:
                flash(f"Could not update: {e}", "err")
        return redirect(url_for("employees"))
    return render_template("employees.html", emps=db.list_employees(active_only=False))


# ------------------------------------------------------------- settings
@app.route("/settings", methods=["GET", "POST"])
def settings():
    if request.method == "POST":
        f = request.form
        try:
            new = {
                "company_name": f["company_name"].strip(),
                "shift_start": engine.norm_time(f["shift_start"]),
                "shift_end": engine.norm_time(f["shift_end"]),
                "late_grace_until": engine.norm_time(f["late_grace_until"]),
                "paid_hours": engine.norm_time(f["paid_hours"]),
                "monthly_permission_hours": float(f["monthly_permission_hours"]),
                "excess_permission_rule": f["excess_permission_rule"],
                "absent_default": f["absent_default"],
                "early_exit_as_permission": bool(f.get("early_exit_as_permission")),
                "short_hours_flag": float(f["short_hours_flag"]),
                "weekly_off_days": [int(x) for x in f.getlist("weekly_off_days")],
            }
            if new["excess_permission_rule"] not in ("half_day_each", "proportional", "none") or new["absent_default"] not in ("LOP", "PL"):
                raise ValueError("invalid option")
            hol = {}
            for line in f.get("holidays", "").splitlines():
                if line.strip():
                    d, _, nm = line.partition(",")
                    hol[datetime.strptime(d.strip(), "%Y-%m-%d").date().isoformat()] = nm.strip()
            new["holidays"] = hol
        except (KeyError, ValueError) as e:
            flash(f"Settings not saved - please check the values ({e}).", "err")
            return redirect(url_for("settings"))
        db.save_settings(new)
        flash("Settings saved. Reports now use the new rules.", "ok")
        return redirect(url_for("settings"))
    st = db.get_settings()
    hol_text = "\n".join(f"{d}, {n}" for d, n in sorted(st["holidays"].items()))
    return render_template("settings.html", st=st, hol_text=hol_text,
                           imports=_recent_imports())


def _recent_imports():
    with db.conn() as c:
        return [dict(r) for r in c.execute("SELECT * FROM imports ORDER BY id DESC LIMIT 10")]


@app.route("/backup", methods=["POST"])
def backup():
    p = db.backup()
    flash(f"Backup saved: {p.name} (in the data/backups folder).", "ok")
    return redirect(request.referrer or url_for("dashboard"))


@app.route("/export/all.csv")
def export_all():
    buf = io.StringIO()
    rows = db.all_daily_rows()
    w = csv.DictWriter(buf, fieldnames=["name", "date", "code", "in_time", "out_time", "perm_hours", "remarks", "source", "updated_at"])
    w.writeheader()
    w.writerows(rows)
    return Response(buf.getvalue(), mimetype="text/csv",
                    headers={"Content-Disposition": "attachment; filename=attendance_all_data.csv"})


def create_app(data_dir=None):
    if data_dir:
        db.configure(data_dir)
    db.init_db()
    return app


if __name__ == "__main__":
    create_app()
    threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:5000")).start()
    app.run(host="127.0.0.1", port=5000, debug=False)
