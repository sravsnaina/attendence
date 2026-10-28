"""Importers: fingerprint-sheet PDF, flat Excel/CSV timesheet, photo (OCR, beta).

Every importer returns a list of blocks:
    {"name": str, "designation": str, "source_page": int|None,
     "rows": [{"date": "YYYY-MM-DD", "code": "P", "in_time": "08:43", "out_time": "17:59",
               "remarks": "", "perm_hours": 0}]}
Only the attendance code and the IN / OUT punches are read. Paid / Actual / Overtime
columns are ignored because the engine recomputes them.
"""
from __future__ import annotations

import csv
import difflib
import io
import re
from datetime import date, datetime

from engine import norm_time

DATE_RE = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})$")
TIME_RE = re.compile(r"^\d{1,2}:\d{2}$")
HEAD_COLS = ["Date", "Attendance", "IN*", "IN", "OUT*", "OUT", "Paid", "Actual", "Overtime"]


class ImportProblem(Exception):
    pass


def _decide_day_first(date_words):
    """The sheet prints m/d/yyyy. If any first number is > 12 it must be d/m/yyyy."""
    firsts = [int(DATE_RE.match(w).group(1)) for w in date_words]
    return any(f > 12 for f in firsts)


def _iso(text, day_first):
    m = DATE_RE.match(text)
    a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    d, mo = (a, b) if day_first else (b, a)
    return date(y, mo, d).isoformat()


# ------------------------------------------------------------------ PDF
def parse_pdf(src) -> list[dict]:
    import pdfplumber

    blocks, warnings = [], []
    with pdfplumber.open(src) as pdf:
        for pno, page in enumerate(pdf.pages, start=1):
            words = page.extract_words(x_tolerance=2, y_tolerance=2)
            heads = sorted((w for w in words if w["text"].lower().startswith("attend")),
                           key=lambda w: w["top"])
            for hi, h in enumerate(heads):
                top_limit = heads[hi + 1]["top"] - 95 if hi + 1 < len(heads) else page.height
                blk = _parse_block(words, h, top_limit, pno)
                if blk:
                    blocks.append(blk)
    if not blocks:
        raise ImportProblem("No attendance tables found in this PDF. Is it the fingerprint sheet exported from Google Sheets?")
    return blocks


def _parse_block(words, head_word, bottom, pno):
    htop = head_word["top"]
    hline = [w for w in words if abs(w["top"] - htop) < 4]
    centers = {}
    for w in hline:
        if w["text"] in HEAD_COLS and w["text"] not in centers:
            centers[w["text"]] = (w["x0"] + w["x1"]) / 2
    if not all(c in centers for c in ("Date", "Attendance", "IN", "OUT")):
        return None
    cols = {k: centers[k] for k in HEAD_COLS if k in centers}
    right_limit = max(cols.values()) + 60

    # ---- title: the 2 text lines above the "IN Time" line
    skip = {"IN", "Time", "OUT", "Working", "Hours"}
    title_words = [w for w in words if htop - 100 < w["top"] < htop - 40 and w["text"] not in skip]
    lines = {}
    for w in title_words:
        lines.setdefault(round(w["top"] / 8), []).append(w)
    ordered = [" ".join(x["text"] for x in sorted(v, key=lambda t: t["x0"])) for _, v in sorted(lines.items())]
    name = ordered[0] if ordered else ""
    designation = ordered[1] if len(ordered) > 1 else ""

    # ---- rows anchored on the date words
    date_words = [w for w in words if DATE_RE.match(w["text"]) and htop + 6 < w["top"] < bottom
                  and abs(((w["x0"] + w["x1"]) / 2) - cols["Date"]) < 60]
    date_words.sort(key=lambda w: w["top"])
    if not date_words:
        return None
    day_first = _decide_day_first([w["text"] for w in date_words])

    rows = []
    for dw in date_words:
        cells = {k: [] for k in cols}
        for w in words:
            if abs(w["top"] - dw["top"]) > 7 or w is dw:
                continue
            xc = (w["x0"] + w["x1"]) / 2
            if xc > right_limit or xc < cols["Date"] + 45:
                continue
            nearest = min(cols, key=lambda k: abs(cols[k] - xc))
            cells[nearest].append(w)
        txt = {k: " ".join(x["text"] for x in sorted(v, key=lambda t: t["x0"])) for k, v in cells.items()}
        code = txt.get("Attendance", "").strip()
        tin, tout = txt.get("IN", "").strip(), txt.get("OUT", "").strip()
        problems = []
        for label, t in (("IN", tin), ("OUT", tout)):
            if t and not TIME_RE.match(t):
                problems.append(f"{label} '{t}' is not a time")
        try:
            tin_n, tout_n = norm_time(tin), norm_time(tout)
        except ValueError:
            tin_n, tout_n = tin, tout
        rows.append({"date": _iso(dw["text"], day_first), "code": code,
                     "in_time": tin_n, "out_time": tout_n, "remarks": "", "perm_hours": 0,
                     "problems": problems})
    return {"name": name.strip(), "designation": designation.strip(), "source_page": pno, "rows": rows}


# ------------------------------------------------------- flat Excel / CSV
_ALIASES = {
    "employee": {"employee", "name", "employee name", "emp", "no.of employees"},
    "date": {"date"},
    "code": {"code", "attendance", "status", "attendance code"},
    "in_time": {"in", "in time", "intime", "in_time"},
    "out_time": {"out", "out time", "outtime", "out_time"},
    "remarks": {"remarks", "remark", "note", "notes"},
    "perm_hours": {"permission", "permission hours", "perm hours", "perm_hours"},
}


def _col_map(header):
    m = {}
    for i, h in enumerate(header):
        key = str(h or "").strip().lower()
        for field, names in _ALIASES.items():
            if key in names and field not in m:
                m[field] = i
    for need in ("employee", "date"):
        if need not in m:
            raise ImportProblem(f"Column '{need.title()}' not found. Use the template: Employee, Date, Code, IN, OUT, Remarks, Permission")
    return m


def _to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ImportProblem(f"Cannot read date '{s}' (use YYYY-MM-DD)")


def _to_time(v):
    if v is None or v == "":
        return ""
    if hasattr(v, "hour") and hasattr(v, "minute"):
        return f"{v.hour:02d}:{v.minute:02d}"
    if isinstance(v, float) and 0 <= v < 1:  # excel time fraction
        mins = round(v * 24 * 60)
        return f"{mins // 60:02d}:{mins % 60:02d}"
    return norm_time(str(v))


def _rows_to_blocks(header, body):
    cm = _col_map(header)
    grouped = {}
    for line_no, r in enumerate(body, start=2):
        if not any(str(c or "").strip() for c in r):
            continue
        get = lambda f: r[cm[f]] if f in cm and cm[f] < len(r) else ""
        name = str(get("employee") or "").strip()
        d = _to_date(get("date"))
        if not name or d is None:
            raise ImportProblem(f"Row {line_no}: employee and date are required")
        try:
            tin, tout = _to_time(get("in_time")), _to_time(get("out_time"))
        except ValueError as e:
            raise ImportProblem(f"Row {line_no}: {e}")
        try:
            perm = float(get("perm_hours") or 0)
        except ValueError:
            raise ImportProblem(f"Row {line_no}: permission hours must be a number")
        grouped.setdefault(name, []).append({
            "date": d.isoformat(), "code": str(get("code") or "").strip(),
            "in_time": tin, "out_time": tout,
            "remarks": str(get("remarks") or "").strip(), "perm_hours": perm, "problems": []})
    if not grouped:
        raise ImportProblem("No data rows found in the file")
    return [{"name": n, "designation": "", "source_page": None, "rows": rows} for n, rows in grouped.items()]


def parse_csv(src) -> list[dict]:
    text = src.read().decode("utf-8-sig") if hasattr(src, "read") else open(src, encoding="utf-8-sig").read()
    rows = list(csv.reader(io.StringIO(text)))
    if len(rows) < 2:
        raise ImportProblem("The CSV has no data rows")
    return _rows_to_blocks(rows[0], rows[1:])


def parse_xlsx(src) -> list[dict]:
    import openpyxl
    wb = openpyxl.load_workbook(src, data_only=True)
    ws = wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    if len(rows) < 2:
        raise ImportProblem("The Excel sheet has no data rows")
    return _rows_to_blocks(rows[0], rows[1:])


def template_csv() -> str:
    return ("Employee,Date,Code,IN,OUT,Remarks,Permission\n"
            "Prem Sundar S,2026-10-26,P,08:55,17:40,,\n"
            "Prem Sundar S,2026-10-27,P,09:25,17:45,,\n"
            "Gopal P,2026-10-26,OD,,,Client visit,\n"
            "Gopal P,2026-10-27,WFH,,,,\n"
            "Surya S,2026-10-26,PL,,,Approved leave,\n")


# ----------------------------------------------------------- photo (OCR)
def parse_image(path) -> list[dict]:
    """Best-effort OCR. Every row is saved as source='ocr' and flagged for verification."""
    try:
        import pytesseract
        from PIL import Image, ImageOps
    except ImportError:
        raise ImportProblem("Photo import needs 'pytesseract' + the Tesseract program installed. Use the PDF or the daily entry grid instead.")
    try:
        img = ImageOps.autocontrast(ImageOps.grayscale(Image.open(path)))
        text = pytesseract.image_to_string(img, config="--psm 6")
    except Exception as e:  # tesseract binary missing etc.
        raise ImportProblem(f"OCR failed: {e}")
    return blocks_from_ocr_text(text)


def blocks_from_ocr_text(text: str) -> list[dict]:
    date_re = re.compile(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})")
    time_re = re.compile(r"\b([01]?\d|2[0-3])[:.]([0-5]\d)\b")
    rows = []
    for line in text.splitlines():
        m = date_re.search(line)
        if not m:
            continue
        a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            d = date(y, b, a) if a > 12 else date(y, a, b) if b > 12 else date(y, b, a)
        except ValueError:
            continue
        times = [f"{int(h):02d}:{mi}" for h, mi in time_re.findall(line[m.end():])]
        rows.append({"date": d.isoformat(), "code": "", "in_time": times[0] if times else "",
                     "out_time": times[-1] if len(times) > 1 else "", "remarks": "",
                     "perm_hours": 0, "problems": []})
    if not rows:
        raise ImportProblem("No dates/times could be read from the photo. Try a straighter, brighter photo or use the PDF.")
    return [{"name": "", "designation": "", "source_page": None, "rows": rows}]


# ---------------------------------------------------- employee matching
def _key(s):
    return re.sub(r"[^a-z0-9 ]", "", (s or "").lower()).strip()


def match_employee(name, employees):
    """Returns (employee_or_None, how) where how in exact | fuzzy | none."""
    k = _key(name)
    if not k:
        return None, "none"
    for e in employees:
        names = [e["name"]] + list(e.get("aliases", []))
        if k in {_key(n) for n in names}:
            return e, "exact"
    best, score = None, 0.0
    for e in employees:
        for n in [e["name"]] + list(e.get("aliases", [])):
            r = difflib.SequenceMatcher(None, k, _key(n)).ratio()
            if r > score:
                best, score = e, r
    if best is not None and score >= 0.8:
        return best, "fuzzy"
    return None, "none"
