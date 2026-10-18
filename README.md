# Excel Automation for PM Teams

**by Muhammad Bilal** · 100% local · No external APIs · No API keys · No network calls

Two automations that remove the weekly spreadsheet grind for a project
management team. Both use **fictional demo data** — swap in your own tracker
and they work the same way.

## The automations

### 1. Task Dashboard — One-Click Excel Dashboard (VBA)
📁 [`task-dashboard-vba/`](task-dashboard-vba/)

A VBA macro (`BuildDashboard`) turns a flat `Task Export` sheet into a
formatted `Dashboard` sheet: KPI tiles (total / completed / in-progress /
overdue / completion %), tasks by owner, tasks by status, an overdue list in
red, and a due-in-7-days list in amber. One click, under a second.

- **Deliverables:** `Task-Dashboard.xlsm` (40 sample tasks + embedded macro),
  `TaskDashboard.bas` (importable VBA source), `build_workbook.py`,
  `build_vba_project.py`, `run_macro.py`, before/after screenshots,
  `README.md`, `TEST-REPORT.md`, demo video.
- **Proof:** the macro was executed for real via the LibreOffice UNO bridge;
  the embedded VBA project was validated with `oletools.olevba` as
  byte-identical to the source; Dashboard KPIs (40 / 12 / 14 / 9 / 30%)
  hand-verified.
- **Run in Excel:** open the `.xlsm`, Enable Content, Alt+F8 → `BuildDashboard`.
- **Run headless:** `python3 run_macro.py`
- **Video:** `task-dashboard-vba/task-dashboard-vba-demo.mp4`

### 2. Weekly Report Generator — Auto-Trigger Status Pack (Python + openpyxl)
📁 [`weekly-report-generator/`](weekly-report-generator/)

A Python script reads `Task-Tracker.xlsx` and produces a dated weekly report
(`Weekly-Report-YYYY-MM-DD.xlsx` + `.pdf`): Summary KPIs, Overdue, Due This
Week, Workload by Owner. A scheduler runs it every **Monday at 8:00 AM**, so
the report is waiting before the stand-up. Idempotent — re-runs overwrite the
same dated files, never duplicates.

- **Deliverables:** `weekly_report.py`, `Task-Tracker.xlsx` (30 sample tasks),
  `build_tracker.py`, `run_weekly_report.bat` / `.sh`, `task-scheduler.xml`
  (importable Windows Task Scheduler task), `requirements.txt`
  (`openpyxl==3.1.2`), `README.md`, `TEST-REPORT.md`, demo video.
- **Proof:** executed end-to-end (Excel + PDF produced); KPIs (30 / 6 / 8 / 7 /
  20%) hand-verified against the input; idempotency proven by double-run;
  error paths tested (missing file → exit 1, bad date → exit 2).
- **Run manually:** `python weekly_report.py` (or double-click the `.bat` / run the `.sh`).
- **Auto-trigger:** import `task-scheduler.xml` into Windows Task Scheduler,
  or add the cron line from the README (`0 8 * * 1`).
- **Video:** `weekly-report-generator/weekly-report-generator-demo.mp4`

## How the two fit together

| | Task Dashboard (VBA) | Weekly Report Generator (Python) |
|---|---|---|
| Trigger | On demand (one click in Excel) | Scheduled (Monday 8 AM) or on demand |
| Input | `Task Export` sheet | `Task-Tracker.xlsx` |
| Output | `Dashboard` sheet in the same workbook | Dated `.xlsx` + `.pdf` in `output/` |
| Best for | Live, in-meeting status checks | The recurring Monday status pack |

Use the Dashboard when you need the numbers *now*, in the meeting. Use the
Weekly Report Generator when the same pack must land in everyone's inbox
folder every Monday without anyone remembering to make it.

## Honest limitations

- The VBA macro is built for **Excel** as its host. LibreOffice can *run* it
  (proven via UNO) but its VBA import did not load the embedded project in
  this environment, and LibreOffice cannot author VBA projects — see the
  Task Dashboard `TEST-REPORT.md`.
- Both automations trust the input format: the `Task Export` / `Tasks` header
  row and column order must match the README. Bad input fails loudly with a
  clear error rather than silently producing wrong numbers.
- PDF conversion in the Weekly Report Generator needs LibreOffice installed;
  without it, only the `.xlsx` is produced.
- The Windows scheduler XML ships with a placeholder folder path
  (`C:\Automation\weekly-report-generator`) — edit the two paths to match
  your machine before importing (see the Weekly Report README).

## Repo layout

```
03-Excel-VBA-Automation/
├── README.md                     ← this file
├── build_videos.py               ← video assembly (TTS + stills + ffmpeg)
├── render_cards.py               ← title / explainer / terminal cards
├── video_assets/                 ← rendered cards
├── task-dashboard-vba/
│   ├── Task-Dashboard.xlsm       ← deliverable workbook (macro + Dashboard)
│   ├── TaskDashboard.bas         ← VBA source
│   ├── build_workbook.py / build_vba_project.py / run_macro.py
│   ├── README.md / TEST-REPORT.md
│   ├── screenshots/              ← before / after
│   └── task-dashboard-vba-demo.mp4
└── weekly-report-generator/
    ├── weekly_report.py          ← the generator
    ├── Task-Tracker.xlsx         ← sample input (30 tasks)
    ├── build_tracker.py
    ├── run_weekly_report.bat / .sh
    ├── task-scheduler.xml        ← Monday 8 AM task
    ├── requirements.txt          ← openpyxl==3.1.2
    ├── README.md / TEST-REPORT.md
    ├── output/                   ← generated reports
    └── weekly-report-generator-demo.mp4
```
