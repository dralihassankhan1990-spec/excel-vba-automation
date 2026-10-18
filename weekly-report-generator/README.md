# Weekly Report Generator — Auto-Trigger Status Pack (Python + openpyxl)

**by Muhammad Bilal** · 100% local · No network calls · No API keys · No cloud

## The problem

Every Monday morning the PM team hand-builds the same status pack: count
tasks, find overdue items, list what is due this week, tally workload per
owner — then format it in Excel and export a PDF to share. It takes 30–45
minutes, it is error-prone, and it never happens on time.

## The solution

One script — `weekly_report.py` — reads `Task-Tracker.xlsx` and produces,
in seconds:

| Output | Contents |
|---|---|
| `output/Weekly-Report-YYYY-MM-DD.xlsx` | **Summary** (KPI table), **Overdue** (red), **Due This Week** (amber), **Workload by Owner** |
| `output/Weekly-Report-YYYY-MM-DD.pdf` | Same report as PDF (via LibreOffice), ready to share |

Then a scheduler runs it automatically every **Monday at 8:00 AM**, so the
report is waiting in `output/` before the stand-up.

## Files

| File | Purpose |
|---|---|
| `weekly_report.py` | The generator (openpyxl + LibreOffice PDF conversion) |
| `Task-Tracker.xlsx` | Sample input — 30 fictional PM tasks (regenerate with `build_tracker.py`) |
| `build_tracker.py` | Regenerates the sample tracker (reproducible demo data) |
| `run_weekly_report.sh` | Linux/mac launcher (manual double-click / cron) |
| `run_weekly_report.bat` | Windows launcher (manual double-click / Task Scheduler) |
| `task-scheduler.xml` | Importable Windows Task Scheduler task — weekly, Monday 8:00 AM |
| `requirements.txt` | Pinned dependencies (`openpyxl==3.1.2`) |
| `TEST-REPORT.md` | Real run output proving end-to-end execution |

## Quick start

```bash
pip install -r requirements.txt
python weekly_report.py
# → output/Weekly-Report-YYYY-MM-DD.xlsx + .pdf
```

Manual run on Windows: double-click `run_weekly_report.bat`.
Manual run on Linux/mac: `./run_weekly_report.sh`.

### Options

```bash
python weekly_report.py --input path/to/My-Tracker.xlsx   # use your own tracker
python weekly_report.py --date 2026-10-19                # report as of a specific date
python weekly_report.py --no-pdf                         # Excel only (skip PDF)
python weekly_report.py --outdir reports/                 # different output folder
```

## Input format (`Task-Tracker.xlsx`)

One sheet named `Tasks` with this exact header row:

| Task ID | Title | Project | Owner | Status | Priority | Due Date |
|---|---|---|---|---|---|---|
| T-001 | … | … | … | Completed / In Progress / Not Started / Blocked | Critical / High / Medium / Low | real Excel date |

The script validates the header and fails with a clear error if it does not
match. Dates must be real Excel dates (not text) for overdue math to work.

## Auto-trigger setup

### Windows — Task Scheduler (recommended)

1. Copy this folder to e.g. `C:\Automation\weekly-report-generator\`.
2. Open `task-scheduler.xml` and check the two paths (they must match your
   folder): the `/c "...run_weekly_report.bat"` argument and
   `<WorkingDirectory>`. Edit if you placed the folder elsewhere.
3. Open **Task Scheduler** → **Action → Import Task…** → select
   `task-scheduler.xml` → **OK**.
4. Right-click the task → **Run** to test it once. Check `output/` for the
   report and `weekly-report.log` for the run log.
5. The task now runs every **Monday at 8:00 AM**, retries up to 3 times on
   failure (5-minute intervals), and also runs if the PC was off at 8:00
   (`StartWhenAvailable`).

Alternative (command line, admin PowerShell):

```powershell
schtasks /create /tn "Weekly Report Generator" /xml "C:\Automation\weekly-report-generator\task-scheduler.xml"
```

### Linux / macOS — cron

```bash
crontab -e
# add this line (adjust the path to where you put the folder):
0 8 * * 1 cd /path/to/weekly-report-generator && ./run_weekly_report.sh
```

`0 8 * * 1` = every Monday at 08:00. Verify with `crontab -l`.

## Idempotency

The output filename is derived from the **report date**
(`Weekly-Report-YYYY-MM-DD.xlsx`). Re-running for the same date overwrites
the same two files — it never creates duplicates. Each week's report is a
separate file (different date → different filename), so history is kept.

## How it works

```
Task-Tracker.xlsx
        │  weekly_report.py (openpyxl)
        ▼
load + validate → compute KPIs (overdue / due-this-week / per-owner / per-status)
        ▼
write 4-sheet .xlsx (formatted: header fills, frozen panes, fit-to-page)
        ▼
soffice --convert-to pdf  (LibreOffice headless; skipped with a warning if absent)
```

## Error handling & logging

- Missing input file, wrong sheet, or wrong header → clear error, exit code 1.
- Bad `--date` → clear error, exit code 2.
- LibreOffice missing → warning, Excel still delivered, exit code 0.
- Every run appends to `weekly-report.log` (console + file).

## Limitations (honest)

- The report is only as good as the tracker: missing dates or typo'd statuses
  produce a hard error rather than a silent wrong report (by design).
- PDF conversion needs LibreOffice installed (`soffice` in PATH); on a PC
  without it, only the `.xlsx` is produced.
- The Windows scheduler XML ships with a placeholder folder
  (`C:\Automation\weekly-report-generator`) — edit the two paths to match
  your machine before importing.
- Single-user file automation: if the tracker is open in Excel with unsaved
  changes, the script reads the last saved version.
