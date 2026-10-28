# Attendance Manager

Import the fingerprint-sheet PDF, fix the few days that need a decision, download the monthly Excel for the CEO.
Everything you import or type is saved in `data/attendance.db` for future months.

## Start (Windows)
1. Install Python 3.10+ from python.org (tick **Add Python to PATH**).
2. Double-click **run.bat** (first time it installs the libraries, then opens the browser at http://127.0.0.1:5000).

Mac/Linux: `./run.sh`

## Every month (about 5 minutes)
1. **Import PDF / file** -> choose the fingerprint PDF -> check the preview -> *Save to database*.
2. **Monthly report** -> the *Status* column shows what is left (red = needs your decision).
3. Fix them:
   - **Enter a day for everyone** - type the latest IN/OUT from the machine for all employees at once.
   - **Open days** (per employee) - correct any code or time.
   - **Mark leave / OD / WFH for a date range** - paid leave, on duty, work from home.
4. **Download Excel for CEO** (first sheet = your standard format). **Save month snapshot** keeps a copy.

## What it calculates
| Item | Rule |
|---|---|
| Working days | Days in month - Week off - Granted holidays |
| Present | P + OD + WFH (half days 0.5) |
| Late | IN after 09:10 = 1 hour permission (10:01 -> 2 hours, and so on) |
| Permission quota | 3 hours per month (editable) |
| Over the quota | 0.5 day LOP per extra instance (editable in Settings: proportional / none) |
| Sheet code `A` | Loss of pay (change to `PL` when leave was intimated/approved) |
| Payable days | Days in month - Loss of pay |
| Overtime | (OUT - IN) beyond 08:30 |

Codes: `P` present, `A` absent, `WO` week off, `H` holiday, `PL` paid leave, `LOP` loss of pay, `OD` on duty, `WFH` work from home, `HPL` half day + half paid leave, `HLOP` half day + half LOP. `P(Gate pass)` is read as `P` with a note.

The app **recalculates everything from IN/OUT** - the sheet's own Actual/Overtime columns are ignored
(they show wrong values like 22:22 when a punch is missing).

## Nothing is guessed
Days it cannot decide are listed in **Review items** and are not counted until you fix them:
no code and no punches, missing OUT/IN, unknown code, absent (LOP vs PL), left early, very short day,
photo-read (OCR) rows. Days today/in the future are shown as pending.

## Files you can import
- **PDF** exported from your Google Sheet (exact).
- **Excel / CSV**: columns `Employee, Date, Code, IN, OUT, Remarks, Permission` (template on the Import page).
- **Photo** (beta, needs Tesseract OCR installed): rows are marked *verify* until you open the employee and save.

## Your data
`data/attendance.db` (all records, history of every change, settings) - `data/uploads` (your original PDFs) -
`data/reports` (every Excel produced) - `data/backups` (automatic before each import/save).
Copy the `data` folder to move to another computer. *Settings -> Export all saved data (CSV)* gives a plain CSV.

## Tests
`python tests/test_all.py` (needs the October PDF; set `ATT_TEST_PDF` to its path).
