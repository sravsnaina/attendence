"""Attendance rules engine.

Pure Python (no database, no web): give it the employees, the day records and the
settings, and it returns the day-by-day result and the monthly summary.

Everything is recomputed from the IN / OUT punches and the attendance code.
The Actual / Overtime columns of the source sheet are NOT trusted (the sheet
shows wrong values such as 22:22 or 23:25 overtime when a punch is missing).
"""
from __future__ import annotations

import calendar
import math
import re
from datetime import date

DEFAULT_SETTINGS = {
    "company_name": "Enstein-Robots And Automation Private Limited",
    "shift_start": "09:00",
    "shift_end": "17:30",
    # Arrival AFTER this time uses permission hours (09:10 itself is still on time).
    "late_grace_until": "09:10",
    "monthly_permission_hours": 3,
    # What happens when the monthly permission quota is exceeded:
    #   "half_day_each" -> every late/permission instance beyond the quota = 0.5 day LOP
    #   "proportional"  -> hours beyond the quota / paid hours per day = LOP days
    #   "none"          -> only reported, no deduction
    "excess_permission_rule": "half_day_each",
    "paid_hours": "08:30",           # normal working hours per day (overtime starts after this)
    "weekly_off_days": [6],          # Monday=0 ... Sunday=6
    "holidays": {},                  # {"2026-10-02": "Gandhi Jayanti"} used only when a day has no code
    "absent_default": "LOP",         # sheet code "A" = absent without intimation -> LOP (or "PL")
    "early_exit_as_permission": False,
    "short_hours_flag": 4.5,         # flag days worked shorter than this many hours
}

# Final status codes used everywhere
STATUS_LABEL = {
    "P": "Present",
    "OD": "Official On Duty",
    "WFH": "Work From Home",
    "HPL": "Half day present + half day paid leave",
    "HLOP": "Half day present + half day LOP",
    "PL": "Paid / permitted leave",
    "LOP": "Loss of pay (absent without intimation)",
    "WO": "Week off",
    "H": "Granted holiday",
    "UNMARKED": "Not marked - needs your decision",
    "PENDING": "Day not completed yet (today / future)",
}
ENTRY_CODES = ["P", "A", "WO", "H", "PL", "LOP", "OD", "WFH", "HPL", "HLOP"]

CODE_MAP = {
    "P": "P", "PRESENT": "P",
    "A": "A", "AB": "A", "ABS": "A", "ABSENT": "A",
    "WO": "WO", "OFF": "WO", "WEEKOFF": "WO",
    "H": "H", "HOL": "H", "HOLIDAY": "H",
    "PL": "PL", "CL": "PL", "SL": "PL", "EL": "PL", "L": "PL", "LEAVE": "PL",
    "OD": "OD", "WFH": "WFH",
    "LOP": "LOP",
    "HPL": "HPL", "HLOP": "HLOP",
}

# (present, paid leave, loss of pay) weight of one working day
WEIGHTS = {
    "P": (1, 0, 0), "OD": (1, 0, 0), "WFH": (1, 0, 0),
    "HPL": (0.5, 0.5, 0), "HLOP": (0.5, 0, 0.5),
    "PL": (0, 1, 0), "LOP": (0, 0, 1),
}


# ----------------------------------------------------------------- helpers
def to_min(value):
    """'9:05' / '09:05' / '9.05' -> minutes since midnight. '' / None -> None."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    m = re.fullmatch(r"(\d{1,2})[:.](\d{2})(?::\d{2})?", s)
    if not m:
        raise ValueError(f"Bad time '{s}'")
    h, mi = int(m.group(1)), int(m.group(2))
    if h > 23 or mi > 59:
        raise ValueError(f"Bad time '{s}'")
    return h * 60 + mi


def fmt_min(m):
    return "" if m is None else f"{m // 60:02d}:{m % 60:02d}"


def norm_time(value):
    """Normalise to 'HH:MM' ('' when empty). Raises ValueError on garbage."""
    return fmt_min(to_min(value))


def normalize_code(raw):
    """Sheet code -> (code, remark). 'P(Gate pass)' -> ('P', 'Gate pass')."""
    raw = (raw or "").strip()
    if not raw:
        return "", ""
    key, remark = raw, ""
    m = re.fullmatch(r"([A-Za-z]+)\s*\((.*)\)", raw)
    if m:
        key, remark = m.group(1), m.group(2).strip()
    key = re.sub(r"[\s.\-_]+", "", key).upper()
    if key in CODE_MAP:
        return CODE_MAP[key], remark
    return "?", raw


def _r(x, n=3):
    return round(x + 0.0, n)


# ----------------------------------------------------------------- one day
def compute_day(d: date, rec: dict | None, st: dict, today: date | None = None) -> dict:
    rec = rec or {}
    iso = d.isoformat()
    flags = []  # (level, key, message)  level: error | warn | info

    start = to_min(st["shift_start"])
    end = to_min(st["shift_end"])
    grace = to_min(st["late_grace_until"])
    paid_min = to_min(st["paid_hours"])

    raw_code = (rec.get("code") or "").strip()
    code, code_remark = normalize_code(raw_code)

    def _t(key):
        try:
            return to_min(rec.get(key))
        except ValueError:
            flags.append(("error", "BAD_TIME", f"Invalid {key.replace('_', ' ')} '{rec.get(key)}'"))
            return None

    in_m, out_m = _t("in_time"), _t("out_time")
    has_punch = in_m is not None or out_m is not None

    note = code_remark
    absent_raw = False
    if not code:
        if has_punch:
            code = "P"
            flags.append(("info", "CODE_INFERRED", "No code entered; treated as Present from the punches"))
        elif iso in st.get("holidays", {}):
            code = "H"
            note = st["holidays"][iso]
        elif d.weekday() in st.get("weekly_off_days", []):
            code = "WO"
        else:
            code = "UNMARKED"
    elif code == "?":
        flags.append(("error", "UNKNOWN_CODE", f"Unknown attendance code '{raw_code}'"))
        code = "UNMARKED"
    elif code == "A":
        absent_raw = True
        code = st.get("absent_default", "LOP")
        note = (note + " " if note else "") + "Absent without intimation"

    if code == "UNMARKED" and not any(f[1] == "UNKNOWN_CODE" for f in flags):
        if today is not None and d >= today:
            code = "PENDING"
            flags.append(("info", "PENDING", "Day not completed yet - nothing counted"))
        else:
            flags.append(("error", "UNMARKED", "No code and no punches - mark Present / Leave / LOP / OD / WFH"))
    if absent_raw:
        flags.append(("warn", "ABSENT_LOP", f"Absent - counted as {code} (no intimation). If leave was approved, change the code to PL"))

    if rec.get("source") == "ocr":
        flags.append(("warn", "OCR_VERIFY", "Read from a photo (OCR) - verify the times, then save the row"))

    # ---- punches
    worked = None
    if in_m is not None and out_m is not None:
        if out_m <= in_m:
            flags.append(("error", "BAD_TIMES", "OUT is not after IN"))
        else:
            worked = out_m - in_m

    if code in ("P", "HPL", "HLOP"):
        if code == "P" and not has_punch:
            flags.append(("error", "NO_PUNCH", "Marked Present but no punches - enter times or use OD / WFH"))
        if in_m is not None and out_m is None:
            flags.append(("error", "MISSING_OUT", "IN punch present but OUT punch missing"))
        if out_m is not None and in_m is None:
            flags.append(("error", "MISSING_IN", "OUT punch present but IN punch missing"))
    if code in ("WO", "H", "PL", "LOP") and has_punch:
        flags.append(("warn", "PUNCH_ON_NON_WORKING", f"Punches exist but day is {code} - comp-off / overtime / wrong code?"))

    if code == "P":
        if out_m is not None and out_m < end:
            flags.append(("warn", "EARLY_EXIT", f"Left at {fmt_min(out_m)} (before {st['shift_end']})"))
        if worked is not None and worked < st["short_hours_flag"] * 60:
            flags.append(("warn", "SHORT_HOURS", f"Only {fmt_min(worked)} hours worked - half day?"))

    # ---- permission hours (late arrival after grace = 1 hour, +1 hour per extra hour late)
    perm = 0.0
    late_min = 0
    if code == "P":
        if in_m is not None and in_m > grace:
            late_min = in_m - start
            perm += math.ceil(late_min / 60)
        if st.get("early_exit_as_permission") and out_m is not None and out_m < end:
            perm += math.ceil((end - out_m) / 60)
    try:
        perm += float(rec.get("perm_hours") or 0)
    except (TypeError, ValueError):
        flags.append(("error", "BAD_PERM", "Permission hours must be a number"))

    # ---- overtime
    ot = 0
    if code == "P" and worked is not None and worked > paid_min:
        ot = worked - paid_min

    pw, lw, xw = WEIGHTS.get(code, (0, 0, 0))
    return {
        "date": iso, "day": d.day, "weekday": d.strftime("%a"),
        "raw_code": raw_code, "status": code, "note": note,
        "in": fmt_min(in_m), "out": fmt_min(out_m),
        "worked_min": worked, "ot_min": ot, "late_min": late_min,
        "perm_hours": perm, "perm_running": None, "perm_over": False, "deduction": 0.0,
        "present_w": float(pw), "leave_w": float(lw), "lop_w": float(xw),
        "flags": flags, "source": rec.get("source", ""),
        "remarks": rec.get("remarks", "") or "",
    }


# ------------------------------------------------------------ one employee
def compute_employee(emp: dict, year: int, month: int, recs: dict, st: dict, today: date | None = None) -> dict:
    """recs: {iso_date: record dict} for this employee."""
    ndays = calendar.monthrange(year, month)[1]
    days = [compute_day(date(year, month, n), recs.get(date(year, month, n).isoformat()), st, today)
            for n in range(1, ndays + 1)]

    quota = float(st["monthly_permission_hours"])
    paid_day_hours = to_min(st["paid_hours"]) / 60
    rule = st.get("excess_permission_rule", "half_day_each")
    used = 0.0
    for dr in days:
        h = dr["perm_hours"]
        if h <= 0:
            continue
        used += h
        dr["perm_running"] = used
        if used > quota + 1e-9:
            over_h = min(h, used - quota)
            if rule == "half_day_each":
                ded = 0.5
            elif rule == "proportional":
                ded = round(over_h / paid_day_hours, 2)
            else:
                ded = 0.0
            ded = min(ded, dr["present_w"])
            dr["present_w"] -= ded
            dr["lop_w"] += ded
            dr["deduction"] = ded
            dr["perm_over"] = True
            dr["flags"].append(("warn", "PERM_OVER_QUOTA",
                                f"Permission beyond the {quota:g}-hour monthly quota"
                                + (f" - {ded:g} day LOP applied" if ded else " - no deduction (rule: none)")))

    count = lambda s: sum(1 for x in days if x["status"] == s)
    wo, hol = count("WO"), count("H")
    present = _r(sum(x["present_w"] for x in days))
    leave = _r(sum(x["leave_w"] for x in days))
    lop = _r(sum(x["lop_w"] for x in days))
    unresolved = count("UNMARKED")
    pending = count("PENDING")
    working = ndays - wo - hol
    summary = {
        "days": ndays, "working": working, "weekoff": wo, "holidays": hol,
        "present": present, "leave": leave, "lop": lop, "unresolved": unresolved, "pending": pending,
        "od": count("OD"), "wfh": count("WFH"),
        "perm_used": _r(used), "perm_left": _r(max(0.0, quota - used)),
        "perm_excess": _r(max(0.0, used - quota)),
        "ot_min": sum(x["ot_min"] for x in days),
        "payable": _r(ndays - lop),
        "balanced": abs(present + leave + lop + unresolved + pending - working) < 1e-6,
        "errors": sum(1 for x in days for f in x["flags"] if f[0] == "error"),
        "warnings": sum(1 for x in days for f in x["flags"] if f[0] == "warn"),
    }
    return {"emp": emp, "days": days, "summary": summary}


def compute_month(year: int, month: int, employees: list, records: dict, st: dict, today: date | None = None) -> list:
    """records: {(emp_id, iso_date): record}. Returns a list of employee results."""
    out = []
    for emp in employees:
        recs = {iso: r for (eid, iso), r in records.items() if eid == emp["id"]}
        out.append(compute_employee(emp, year, month, recs, st, today))
    return out


def merge_settings(saved: dict | None) -> dict:
    st = dict(DEFAULT_SETTINGS)
    st["holidays"] = dict(DEFAULT_SETTINGS["holidays"])
    if saved:
        st.update({k: v for k, v in saved.items() if k in DEFAULT_SETTINGS})
    return st
