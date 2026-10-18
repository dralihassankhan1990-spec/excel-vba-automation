# TEST-REPORT — Weekly Report Generator

**Date:** 2026-10-16 · **Runner:** `python3 weekly_report.py --date 2026-10-16`
**Environment:** Ubuntu, Python 3.12, openpyxl 3.1.2, LibreOffice 24.2
**Result:** ✅ PASS — Excel + PDF generated, KPIs hand-verified, idempotency proven.

## 1. End-to-end run (real output)

```
Wrote Task-Tracker.xlsx with 30 tasks.
Loaded 30 tasks from .../Task-Tracker.xlsx.
Wrote .../output/Weekly-Report-2026-10-16.xlsx
KPIs: total=30 completed=6 overdue=8 due_this_week=7 completion=20%
Wrote .../output/Weekly-Report-2026-10-16.pdf
DONE — .../output/Weekly-Report-2026-10-16.xlsx
```

Output files: `output/Weekly-Report-2026-10-16.xlsx` (9,254 bytes),
`output/Weekly-Report-2026-10-16.pdf` (37,748 bytes).

## 2. Output contents verified (read back from the generated .xlsx)

- Sheets: `Summary`, `Overdue`, `Due This Week`, `Workload by Owner`.
- Summary: Total 30, Completed 6, In Progress 12, Blocked 3, Not Started 9,
  Overdue 8, Due this week 7, Completion 20%.
- Overdue sheet: 8 rows, red fill, sorted oldest first
  (T-016, 2026-10-05, 11 days overdue … T-009, 2026-10-15, 1 day overdue).
- Due This Week sheet: 7 rows, amber fill.
- Workload by Owner: 6 owners.

## 3. KPIs hand-verified against the input

Counted directly from `Task-Tracker.xlsx` (reference date 2026-10-16):
- Completed = 6 (T-001, T-006, T-012, T-013, T-023, T-028). ✔
- Overdue = 8 (due before 2026-10-16 and not Completed: T-002, T-004, T-009,
  T-011, T-016, T-021, T-025, T-030). ✔
- Due this week = 7 (due 2026-10-16…2026-10-22 and not Completed: T-003,
  T-007, T-014, T-018, T-019, T-024, T-029). ✔
- Completion = 6/30 = 20%. ✔
- A `--date 2026-10-20` run correctly shifts the boundary (overdue 12, due
  this week 5) — date logic is driven by the reference date, not hard-coded.

## 4. Idempotency

Ran `python3 weekly_report.py --date 2026-10-16` twice in a row. The
`output/` folder still contained exactly 2 files (same dated `.xlsx` + `.pdf`,
overwritten, no duplicates).

## 5. Error handling (real exit codes)

| Case | Behaviour | Exit code |
|---|---|---|
| Missing input file | `FileNotFoundError` with a clear message (suggests `build_tracker.py`) | 1 |
| Wrong header row | `ValueError: Unexpected header row: ['Wrong', 'Headers']...` | 1 |
| Bad `--date` (`16-10-2026`) | `[ERROR] --date must be YYYY-MM-DD` | 2 |
| Normal run | Excel + PDF written, `DONE` logged | 0 |
| `--no-pdf` | Excel only, no soffice call | 0 |

## 6. Static checks

- `task-scheduler.xml` parses as well-formed XML (Python `xml.etree`).
- `run_weekly_report.sh` is executable; `.bat` uses `%~dp0` so it works from
  any working directory.
- PDF pages inspected visually: Summary KPI table and Overdue table render
  cleanly with correct values and colour fills.
