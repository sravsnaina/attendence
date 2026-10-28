"""Run:  python -m pytest tests -q      (or)   python tests/test_all.py"""
import io
import os
import shutil
import sys
import tempfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
TMP = tempfile.mkdtemp(prefix="att_test_")
os.environ["ATT_DATA_DIR"] = TMP

import engine  # noqa: E402
import importers  # noqa: E402

PDF = os.environ.get("ATT_TEST_PDF", "/mnt/user-data/uploads/Monthly_Attendance__xlsx_-_Google_Sheets.pdf")
ST = engine.merge_settings({})


def day(rec, d=date(2026, 10, 14), today=None, st=ST):
    return engine.compute_day(d, rec, st, today)


# ------------------------------------------------------------ rules
def test_late_rule():
    assert day({"code": "P", "in_time": "09:10", "out_time": "17:30"})["perm_hours"] == 0   # 09:10 still on time
    assert day({"code": "P", "in_time": "09:11", "out_time": "17:30"})["perm_hours"] == 1   # after 09:10 = 1 hour
    assert day({"code": "P", "in_time": "10:00", "out_time": "18:00"})["perm_hours"] == 1
    assert day({"code": "P", "in_time": "10:25", "out_time": "18:00"})["perm_hours"] == 2
    assert day({"code": "P", "in_time": "08:30", "out_time": "17:30"})["perm_hours"] == 0


def test_statuses_and_weights():
    assert day({"code": "A"})["status"] == "LOP" and day({"code": "A"})["lop_w"] == 1
    assert day({"code": "p", "in_time": "9:00", "out_time": "17:30"})["status"] == "P"
    r = day({"code": "P(Gate pass)", "in_time": "09:06", "out_time": "13:30"})
    assert r["status"] == "P" and "Gate pass" in r["note"]
    assert any(f[1] == "EARLY_EXIT" for f in r["flags"]) and any(f[1] == "SHORT_HOURS" for f in r["flags"])
    for code, w in (("PL", (0, 1, 0)), ("LOP", (0, 0, 1)), ("OD", (1, 0, 0)), ("WFH", (1, 0, 0)),
                    ("HPL", (.5, .5, 0)), ("HLOP", (.5, 0, .5)), ("WO", (0, 0, 0)), ("H", (0, 0, 0))):
        r = day({"code": code})
        assert (r["present_w"], r["leave_w"], r["lop_w"]) == w, code
    assert day({"code": "H"})["status"] == "H"


def test_flags():
    r = day({"code": "P", "in_time": "08:48", "out_time": ""})
    assert any(f[1] == "MISSING_OUT" for f in r["flags"])
    assert any(f[1] == "NO_PUNCH" for f in day({"code": "P"})["flags"])
    assert any(f[1] == "UNMARKED" for f in day({})["flags"])
    assert any(f[1] == "UNKNOWN_CODE" for f in day({"code": "ZZ"})["flags"])
    assert any(f[1] == "BAD_TIME" for f in day({"code": "P", "in_time": "25:99", "out_time": "17:30"})["flags"])
    assert any(f[1] == "BAD_TIMES" for f in day({"code": "P", "in_time": "18:00", "out_time": "17:30"})["flags"])
    assert any(f[1] == "ABSENT_LOP" for f in day({"code": "A"})["flags"])
    assert any(f[1] == "PUNCH_ON_NON_WORKING" for f in day({"code": "WO", "in_time": "09:00", "out_time": "12:00"})["flags"])


def test_pending_and_inference():
    r = day({}, d=date(2026, 10, 28), today=date(2026, 10, 28))
    assert r["status"] == "PENDING" and not any(f[0] == "error" for f in r["flags"])
    assert day({}, d=date(2026, 10, 27), today=date(2026, 10, 28))["status"] == "UNMARKED"
    assert day({}, d=date(2026, 10, 4))["status"] == "WO"          # Sunday, no code
    assert day({"in_time": "09:00", "out_time": "17:30"})["status"] == "P"   # code inferred from punches


def test_overtime():
    r = day({"code": "P", "in_time": "08:43", "out_time": "17:59"})
    assert r["worked_min"] == 9 * 60 + 16 and r["ot_min"] == 46     # same as the sheet's 09:16 / 00:46
    assert day({"code": "P", "in_time": "09:00", "out_time": "17:30"})["ot_min"] == 0


def _month(recs, st=ST, today=None):
    emp = {"id": 1, "name": "T", "designation": ""}
    return engine.compute_employee(emp, 2026, 10, {r["date"]: r for r in recs}, st, today)


def _late(d):
    return {"date": f"2026-10-{d:02d}", "code": "P", "in_time": "09:30", "out_time": "18:00"}


def test_permission_quota_half_day_rule():
    res = _month([_late(5), _late(6), _late(7), _late(8)])      # 4 x 1 hour, quota 3
    s = res["summary"]
    assert s["perm_used"] == 4 and s["perm_excess"] == 1 and s["perm_left"] == 0
    assert s["lop"] == 0.5 and s["present"] == 3.5            # 4th instance = 0.5 day LOP


def test_permission_quota_other_rules():
    st = engine.merge_settings({"excess_permission_rule": "none"})
    assert _month([_late(d) for d in (5, 6, 7, 8)], st)["summary"]["lop"] == 0
    st = engine.merge_settings({"excess_permission_rule": "proportional"})
    assert _month([_late(d) for d in (5, 6, 7, 8)], st)["summary"]["lop"] == round(1 / 8.5, 2)
    ok = _month([_late(d) for d in (5, 6, 7)])["summary"]
    assert ok["perm_used"] == 3 and ok["lop"] == 0 and ok["perm_left"] == 0


def test_summary_balance():
    recs = [{"date": "2026-10-02", "code": "H"}, {"date": "2026-10-04", "code": "WO"},
            {"date": "2026-10-05", "code": "P", "in_time": "09:00", "out_time": "17:30"},
            {"date": "2026-10-06", "code": "PL"}, {"date": "2026-10-07", "code": "A"},
            {"date": "2026-10-08", "code": "HPL"}, {"date": "2026-10-09", "code": "OD"}]
    s = _month(recs)["summary"]
    assert (s["weekoff"], s["holidays"]) == (4, 1) and s["working"] == 26   # 4 Sundays inferred as week off
    assert s["present"] == 2.5 and s["leave"] == 1.5 and s["lop"] == 1
    assert s["unresolved"] == 26 - 2.5 - 1.5 - 1 and s["balanced"]


# --------------------------------------------------------- PDF import
def test_pdf_parse():
    blocks = importers.parse_pdf(PDF)
    assert len(blocks) == 11
    by = {b["name"]: {r["date"]: r for r in b["rows"]} for b in blocks}
    r = by["PREM SUNDAR"]["2026-10-21"]
    assert (r["code"], r["in_time"], r["out_time"]) == ("P", "10:46", "17:38")
    assert by["SURYA"]["2026-10-07"]["code"] == "P(Gate pass)"
    assert by["SURYA"]["2026-10-05"]["in_time"] == "09:05"        # '9:05' normalised
    assert by["BHARATH"]["2026-10-07"]["out_time"] == "17:36"
    assert by["THANIGAI MALAI A"]["2026-10-08"]["out_time"] == "18:55"
    assert by["SRAVANI"]["2026-10-01"]["code"] == "p"
    assert by["BHARATH"]["2026-10-08"]["code"] == "A"
    assert [len(b["rows"]) for b in blocks] == [30, 30, 30, 31, 31, 30, 31, 31, 31, 31, 31]


def test_match_employee():
    import db
    db.configure(TMP)
    db.init_db()
    emps = db.list_employees()
    assert importers.match_employee("SRAVANI", emps)[0]["name"] == "Sarvani Lakshmi M"
    assert importers.match_employee("BHARATH", emps)[0]["name"] == "Bharath K"
    assert importers.match_employee("BHARATHI", emps)[0]["name"] == "Bharathi K"
    assert importers.match_employee("Somebody Else", emps)[1] == "none"


def test_ocr_text_parser():
    b = importers.blocks_from_ocr_text("15/10/2026 08:44 17:55\nrubbish\n16/10/2026 09:00 17:30\n17/10/2026 09:12")
    rows = b[0]["rows"]
    assert [r["date"] for r in rows] == ["2026-10-15", "2026-10-16", "2026-10-17"]
    assert rows[0]["in_time"] == "08:44" and rows[0]["out_time"] == "17:55"
    assert rows[2]["in_time"] == "09:12" and rows[2]["out_time"] == ""


# -------------------------------------------------------- whole app
def test_full_flow():
    import app as appmod
    import db
    import openpyxl

    class FixedDate(date):
        @classmethod
        def today(cls):
            return date(2026, 10, 28)
    appmod.date = FixedDate
    flask_app = appmod.create_app(TMP)
    flask_app.config["TESTING"] = True
    c = flask_app.test_client()

    for url in ("/", "/import", "/review?m=2026-10", "/employees", "/settings", "/template.csv", "/export/all.csv"):
        assert c.get(url).status_code == 200, url

    # import PDF -> preview -> confirm
    r = c.post("/import", data={"file": (open(PDF, "rb"), "attendance.pdf")}, content_type="multipart/form-data")
    assert r.status_code == 302
    token = r.headers["Location"].rsplit("/", 1)[-1]
    pv = c.get(f"/import/{token}")
    assert pv.status_code == 200 and b"SRAVANI" in pv.data
    emps = db.list_employees()
    form = {}
    names = [b["name"] for b in importers.parse_pdf(PDF)]
    for i, n in enumerate(names):
        form[f"map_{i}"] = str(importers.match_employee(n, emps)[0]["id"])
    r = c.post(f"/import/{token}/confirm", data=form)
    assert r.status_code == 302
    with db.conn() as cx:
        n1 = cx.execute("SELECT COUNT(*) FROM daily").fetchone()[0]
    assert n1 > 100

    # dashboard numbers for the actual October data
    html = c.get("/?m=2026-10").data.decode()
    assert "Prem Sundar S" in html and "need your decision" in html

    st = db.get_settings()
    res = engine.compute_month(2026, 10, db.list_employees(), db.get_records(2026, 10), st, today=date(2026, 10, 28))
    S = {r["emp"]["name"]: r["summary"] for r in res}
    assert S["Prem Sundar S"]["present"] == 19 and S["Prem Sundar S"]["perm_used"] == 2
    assert S["Gopal P"]["present"] == 19 and S["Gopal P"]["lop"] == 0
    assert S["Bharath K"]["present"] == 5 and S["Bharath K"]["lop"] == 14
    assert S["Surya S"]["present"] == 6 and S["Surya S"]["lop"] == 2
    assert S["Dharanidharan"]["lop"] == 2
    for s in S.values():
        assert s["days"] == 31 and s["working"] == 25 and s["weekoff"] == 4 and s["holidays"] == 2 and s["balanced"]

    # re-import is idempotent
    r = c.post("/import", data={"file": (open(PDF, "rb"), "attendance.pdf")}, content_type="multipart/form-data")
    token = r.headers["Location"].rsplit("/", 1)[-1]
    c.post(f"/import/{token}/confirm", data=form)
    with db.conn() as cx:
        assert cx.execute("SELECT COUNT(*) FROM daily").fetchone()[0] == n1

    # manual edit: fill the missing OUT for Prem Sundar on 24th, mark Surya's blank days as PL
    prem = next(e for e in db.list_employees() if e["name"] == "Prem Sundar S")
    data = {"m": "2026-10"}
    recs = db.get_records(2026, 10, prem["id"])
    for i in range(31):
        iso = f"2026-10-{i + 1:02d}"
        rec = recs.get((prem["id"], iso), {})
        data[f"code_{i}"], data[f"in_{i}"], data[f"out_{i}"] = rec.get("code", ""), rec.get("in_time", ""), rec.get("out_time", "")
        data[f"perm_{i}"], data[f"rem_{i}"] = "", rec.get("remarks", "")
    data["out_23"] = "18:00"
    assert c.post(f"/edit/{prem['id']}", data=data).status_code == 302
    r = engine.compute_employee(prem, 2026, 10, {k[1]: v for k, v in db.get_records(2026, 10, prem["id"]).items()}, st, date(2026, 10, 28))
    d24 = r["days"][23]
    assert d24["out"] == "18:00" and not any(f[1] == "MISSING_OUT" for f in d24["flags"])
    assert d24["source"] == "manual"
    bad = dict(data, in_23="99:99")
    c.post(f"/edit/{prem['id']}", data=bad)                       # invalid time must not be saved
    assert db.get_records(2026, 10, prem["id"])[(prem["id"], "2026-10-24")]["in_time"] == "08:48"

    # PDF re-import must NOT overwrite the manual edit
    r = c.post("/import", data={"file": (open(PDF, "rb"), "attendance.pdf")}, content_type="multipart/form-data")
    token = r.headers["Location"].rsplit("/", 1)[-1]
    c.post(f"/import/{token}/confirm", data=form)
    assert db.get_records(2026, 10, prem["id"])[(prem["id"], "2026-10-24")]["out_time"] == "18:00"

    # quick leave: Surya paid leave 12..17 Oct (6 days) - week-off must be skipped when in range
    surya = next(e for e in db.list_employees() if e["name"] == "Surya S")
    c.post("/quick", data={"emp_ids": [str(surya["id"])], "code": "PL", "from": "2026-10-10", "to": "2026-10-13", "remarks": "approved"})
    rec = db.get_records(2026, 10, surya["id"])
    assert rec[(surya["id"], "2026-10-12")]["code"] == "PL" and rec[(surya["id"], "2026-10-11")]["code"] == "WO"
    assert rec[(surya["id"], "2026-10-10")]["code"] == "PL"       # overrides the P that was there (flagged as punch on leave)

    # one-day-for-everyone screen
    assert c.get("/day/2026-10-26").status_code == 200 and c.get("/day/bad").status_code == 404
    ids = [e["id"] for e in db.list_employees()]
    day_form = {}
    for k in ids:
        day_form.update({f"code_{k}": "P", f"in_{k}": "9:25", f"out_{k}": "17:45", f"perm_{k}": "", f"rem_{k}": ""})
    day_form[f"code_{ids[1]}"] = "WFH"; day_form[f"in_{ids[1]}"] = ""; day_form[f"out_{ids[1]}"] = ""
    day_form[f"in_{ids[2]}"] = "xx"                               # invalid -> that row not saved, others saved
    assert c.post("/day/2026-10-26", data=day_form).status_code == 302
    r26 = db.get_records(2026, 10)
    assert r26[(ids[0], "2026-10-26")]["in_time"] == "09:25" and r26[(ids[1], "2026-10-26")]["code"] == "WFH"
    assert (ids[2], "2026-10-26") not in r26
    day = engine.compute_day(date(2026, 10, 26), r26[(ids[0], "2026-10-26")], st, date(2026, 10, 28))
    assert day["perm_hours"] == 1                                 # 09:25 is after 09:10

    # report
    r = c.get("/report.xlsx?m=2026-10")
    assert r.status_code == 200
    wb = openpyxl.load_workbook(io.BytesIO(r.data))
    assert wb.sheetnames[0] == "Attendance Summary"
    ws = wb["Attendance Summary"]
    assert ws["A2"].value == "Attendance for October Month-26"
    assert [ws.cell(row=3, column=i).value.strip() for i in range(1, 11)] == [
        "S.No", "No.Of Employees", "Designation", "No.Of Days", "No.Of Working", "No.Of Weekoff",
        "Granted Holidays", "No.Of present", "Permitted Leave", "Loss Of Pay"]
    assert ws["B4"].value == "Dharanidharan" and ws["D4"].value == 31 and ws["E4"].value == 25
    assert wb.sheetnames == ["Attendance Summary", "Daily Register", "Day-wise Detail", "Permissions & OT", "Review Items", "Rules Used"]

    # other pages / actions
    assert c.get(f"/edit/{prem['id']}?m=2026-10").status_code == 200
    assert c.post("/finalize?m=2026-10").status_code == 302
    assert c.post("/backup").status_code == 302
    assert any(Path(TMP, "backups").glob("*.db"))
    assert c.post("/settings", data={
        "company_name": "X Ltd", "shift_start": "09:00", "shift_end": "17:30", "late_grace_until": "09:15",
        "paid_hours": "08:30", "monthly_permission_hours": "3", "excess_permission_rule": "none",
        "absent_default": "LOP", "short_hours_flag": "4.5", "weekly_off_days": ["6"],
        "holidays": "2026-10-02, Gandhi Jayanti"}).status_code == 302
    assert db.get_settings()["late_grace_until"] == "09:15"
    assert db.get_settings()["holidays"] == {"2026-10-02": "Gandhi Jayanti"}
    assert c.get("/settings").status_code == 200 and c.get("/").status_code == 200

    # CSV manual timesheet import with a new employee
    csv_text = "Employee,Date,Code,IN,OUT,Remarks\nGopal P,2026-10-26,P,08:55,17:40,\nGopal P,2026-10-27,WFH,,,home\nNew Person,2026-10-26,OD,,,client\nGopalan,2026-10-28,OD,,,x\n"
    r = c.post("/import", data={"file": (io.BytesIO(csv_text.encode()), "sheet.csv")}, content_type="multipart/form-data")
    token = r.headers["Location"].rsplit("/", 1)[-1]
    assert b"New Person" in c.get(f"/import/{token}").data
    gopal = next(e for e in db.list_employees() if e["name"] == "Gopal P")
    c.post(f"/import/{token}/confirm", data={"map_0": str(gopal["id"]), "map_1": "new", "map_2": str(gopal["id"])})
    assert "Gopalan" in db.get_employee(gopal["id"])["aliases"]      # alias saved while other rows were being written
    assert db.get_records(2026, 10, gopal["id"])[(gopal["id"], "2026-10-27")]["code"] == "WFH"
    assert any(e["name"] == "New Person" for e in db.list_employees())

    # bad uploads fail politely
    r = c.post("/import", data={"file": (io.BytesIO(b"hello"), "x.txt")}, content_type="multipart/form-data", follow_redirects=True)
    assert b"Unsupported" in r.data
    r = c.post("/import", data={"file": (io.BytesIO(b"not a pdf"), "x.pdf")}, content_type="multipart/form-data", follow_redirects=True)
    assert r.status_code == 200 and b"Could not read" in r.data or b"No attendance" in r.data


if __name__ == "__main__":
    import traceback
    failed = 0
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            try:
                fn()
                print("PASS", name)
            except Exception:
                failed += 1
                print("FAIL", name)
                traceback.print_exc()
    shutil.rmtree(TMP, ignore_errors=True)
    sys.exit(1 if failed else 0)
