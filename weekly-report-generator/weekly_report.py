"""Weekly Report Generator — build the PM team's Monday status pack automatically.

Reads Task-Tracker.xlsx and produces:
  output/Weekly-Report-YYYY-MM-DD.xlsx  (Summary, Overdue, Due This Week, Workload by Owner)
  output/Weekly-Report-YYYY-MM-DD.pdf   (same report as PDF, via LibreOffice)

Idempotent: re-running for the same date overwrites the same files — never
creates duplicates. 100% local: no network calls, no API keys, no cloud.
Error handling + logging throughout. by Muhammad Bilal.

Usage:
    python weekly_report.py [--input Task-Tracker.xlsx] [--outdir output]
                            [--date 2026-10-19] [--no-pdf]
"""
import argparse
import datetime as dt
import logging
import os
import subprocess
import sys
from collections import defaultdict

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

LOG = logging.getLogger("weekly_report")

NAVY = "0E4C81"
GREY = "5A6E82"
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
AMBER_FILL = PatternFill("solid", fgColor="FFEB9C")
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=16, color=NAVY)
SUB_FONT = Font(italic=True, size=10, color="666666")
THIN = Border(*[Side(style="thin", color="B0B0B0")] * 4)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP = Alignment(vertical="center", wrap_text=True)

STATUSES = ("Not Started", "In Progress", "Completed", "Blocked")


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )


def load_tasks(path: str):
    """Read the Tasks sheet -> list of dicts. Raises a clear error if the
    file or the expected columns are missing."""
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Input file not found: {path!r}. Run build_tracker.py to create the sample, "
            "or point --input at your own tracker.")
    wb = openpyxl.load_workbook(path, data_only=True)
    if "Tasks" not in wb.sheetnames:
        raise ValueError(f"{path!r} has no 'Tasks' sheet (found: {wb.sheetnames}).")
    ws = wb["Tasks"]
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    expected = ["Task ID", "Title", "Project", "Owner", "Status", "Priority", "Due Date"]
    if header != expected:
        raise ValueError(f"Unexpected header row: {header!r}. Expected: {expected!r}")
    tasks = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        due = row[6]
        if isinstance(due, dt.datetime):
            due = due.date()
        tasks.append({
            "id": row[0], "title": row[1], "project": row[2], "owner": row[3],
            "status": row[4], "priority": row[5], "due": due,
        })
    LOG.info("Loaded %d tasks from %s.", len(tasks), path)
    return tasks


def compute(tasks, ref: dt.date):
    """KPIs + slices. All date math is relative to the reference date."""
    overdue, due_week = [], []
    by_owner = defaultdict(lambda: {"total": 0, "done": 0, "overdue": 0})
    by_status = {s: 0 for s in STATUSES}
    for t in tasks:
        by_status[t["status"]] = by_status.get(t["status"], 0) + 1
        o = by_owner[t["owner"]]
        o["total"] += 1
        if t["status"] == "Completed":
            o["done"] += 1
        if t["status"] != "Completed" and isinstance(t["due"], dt.date):
            if t["due"] < ref:
                overdue.append({**t, "days": (ref - t["due"]).days})
                o["overdue"] += 1
            elif t["due"] <= ref + dt.timedelta(days=6):
                due_week.append({**t, "days": (t["due"] - ref).days})
    overdue.sort(key=lambda t: t["due"])
    due_week.sort(key=lambda t: t["due"])
    total = len(tasks)
    done = by_status.get("Completed", 0)
    kpis = {
        "Report date": ref.isoformat(),
        "Total tasks": total,
        "Completed": done,
        "In Progress": by_status.get("In Progress", 0),
        "Blocked": by_status.get("Blocked", 0),
        "Not Started": by_status.get("Not Started", 0),
        "Overdue": len(overdue),
        "Due this week": len(due_week),
        "Completion %": (done / total if total else 0),
    }
    return kpis, overdue, due_week, by_owner, by_status


def _sheet_header(ws, title: str, subtitle: str, ncols: int) -> int:
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    ws.cell(row=1, column=1, value=title).font = TITLE_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(row=2, column=1, value=subtitle)
    c.font = SUB_FONT
    c.alignment = WRAP
    ws.row_dimensions[2].height = 45
    return 4


def _table_header(ws, row: int, cols, widths) -> None:
    for i, (h, w) in enumerate(zip(cols, widths), start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.alignment = CENTER
        c.border = THIN
        ws.column_dimensions[get_column_letter(i)].width = w


def _row(ws, row: int, values, fill=None, fmts=None) -> None:
    for i, v in enumerate(values, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.border = THIN
        c.alignment = WRAP
        if fill is not None:
            c.fill = fill
        if fmts and fmts.get(i):
            c.number_format = fmts[i]


def build_report(tasks, ref: dt.date, out_xlsx: str) -> dict:
    kpis, overdue, due_week, by_owner, by_status = compute(tasks, ref)
    wb = openpyxl.Workbook()

    # --- Summary ---
    ws = wb.active
    ws.title = "Summary"
    r = _sheet_header(ws, "WEEKLY STATUS REPORT",
                      f"Week of {ref.isoformat()} · Source: Task-Tracker.xlsx · "
                      "Auto-generated — re-run weekly_report.py to refresh.", 2)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 18
    _table_header(ws, r, ["Metric", "Value"], [22, 18])
    r += 1
    for k, v in kpis.items():
        _row(ws, r, [k, v], fmts={2: "0%" if k == "Completion %" else None})
        if k == "Completion %":
            ws.cell(row=r, column=2).number_format = "0%"
        ws.cell(row=r, column=1).font = Font(bold=True)
        r += 1
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    c = ws.cell(row=r, column=1,
                value="Blocked tasks need a decision this week; see the Overdue sheet.")
    c.font = SUB_FONT
    c.alignment = WRAP
    ws.row_dimensions[r].height = 30

    # --- Overdue ---
    ws = wb.create_sheet("Overdue")
    r = _sheet_header(ws, "OVERDUE TASKS",
                      f"{len(overdue)} task(s) past due as of {ref.isoformat()}. Sorted by due date (oldest first).", 7)
    cols = ["Task ID", "Title", "Project", "Owner", "Due Date", "Days Overdue", "Priority"]
    _table_header(ws, r, cols, [12, 42, 26, 16, 13, 14, 10])
    r += 1
    for t in overdue:
        _row(ws, r, [t["id"], t["title"], t["project"], t["owner"], t["due"], t["days"], t["priority"]],
             fill=RED_FILL, fmts={5: "yyyy-mm-dd"})
        r += 1
    if not overdue:
        ws.cell(row=r, column=1, value="No overdue tasks — nice work.").font = SUB_FONT

    # --- Due This Week ---
    ws = wb.create_sheet("Due This Week")
    r = _sheet_header(ws, "DUE THIS WEEK",
                      f"{len(due_week)} task(s) due {ref.isoformat()} – {(ref + dt.timedelta(days=6)).isoformat()}.", 7)
    _table_header(ws, r, ["Task ID", "Title", "Project", "Owner", "Due Date", "Days Left", "Priority"],
                  [12, 42, 26, 16, 13, 14, 10])
    r += 1
    for t in due_week:
        _row(ws, r, [t["id"], t["title"], t["project"], t["owner"], t["due"], t["days"], t["priority"]],
             fill=AMBER_FILL, fmts={5: "yyyy-mm-dd"})
        r += 1
    if not due_week:
        ws.cell(row=r, column=1, value="Nothing due this week.").font = SUB_FONT

    # --- Workload by Owner ---
    ws = wb.create_sheet("Workload by Owner")
    r = _sheet_header(ws, "WORKLOAD BY OWNER",
                      "Per-owner load and completion. Use it to rebalance before the week starts.", 5)
    _table_header(ws, r, ["Owner", "Total", "Completed", "Overdue", "Completion %"],
                  [18, 10, 12, 10, 14])
    r += 1
    for owner in sorted(by_owner):
        o = by_owner[owner]
        pct = o["done"] / o["total"] if o["total"] else 0
        _row(ws, r, [owner, o["total"], o["done"], o["overdue"], pct], fmts={5: "0%"})
        r += 1

    for sheet in wb.worksheets:
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0

    wb.save(out_xlsx)
    LOG.info("Wrote %s", out_xlsx)
    return kpis


def convert_to_pdf(xlsx_path: str, outdir: str) -> str | None:
    """Convert via LibreOffice headless. Returns the PDF path or None if
    LibreOffice is unavailable (Excel file is still delivered)."""
    try:
        r = subprocess.run(
            ["soffice", "--headless", "--nologo", "--norestore",
             "--convert-to", "pdf", "--outdir", outdir, xlsx_path],
            capture_output=True, text=True, timeout=180)
    except FileNotFoundError:
        LOG.warning("LibreOffice (soffice) not found — skipping PDF conversion.")
        return None
    except subprocess.TimeoutExpired:
        LOG.warning("PDF conversion timed out — skipping.")
        return None
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(xlsx_path))[0] + ".pdf")
    if r.returncode == 0 and os.path.exists(pdf):
        LOG.info("Wrote %s", pdf)
        return pdf
    LOG.warning("PDF conversion failed (soffice exit %s): %s", r.returncode, r.stderr.strip()[-300:])
    return None


def main() -> int:
    setup_logging()
    ap = argparse.ArgumentParser(description="Generate the weekly PM status report.")
    ap.add_argument("--input", default="Task-Tracker.xlsx", help="Input task tracker (.xlsx)")
    ap.add_argument("--outdir", default="output", help="Output directory")
    ap.add_argument("--date", default=None,
                    help="Reference date YYYY-MM-DD (default: today). Same date = same output file.")
    ap.add_argument("--no-pdf", action="store_true", help="Skip PDF conversion")
    args = ap.parse_args()

    try:
        ref = dt.date.fromisoformat(args.date) if args.date else dt.date.today()
    except ValueError:
        LOG.error("--date must be YYYY-MM-DD, got %r.", args.date)
        return 2

    try:
        here = os.path.dirname(os.path.abspath(__file__))
        in_path = args.input if os.path.isabs(args.input) else os.path.join(here, args.input)
        outdir = args.outdir if os.path.isabs(args.outdir) else os.path.join(here, args.outdir)
        os.makedirs(outdir, exist_ok=True)
        log_path = os.path.join(here, "weekly-report.log")
        fh = logging.FileHandler(log_path, encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        LOG.addHandler(fh)

        tasks = load_tasks(in_path)
        fname = f"Weekly-Report-{ref.isoformat()}.xlsx"
        out_xlsx = os.path.join(outdir, fname)
        kpis = build_report(tasks, ref, out_xlsx)
        LOG.info("KPIs: total=%s completed=%s overdue=%s due_this_week=%s completion=%.0f%%",
                 kpis["Total tasks"], kpis["Completed"], kpis["Overdue"],
                 kpis["Due this week"], kpis["Completion %"] * 100)
        if not args.no_pdf:
            convert_to_pdf(out_xlsx, outdir)
        LOG.info("DONE — %s", out_xlsx)
        return 0
    except Exception as exc:  # noqa: BLE001 - top-level guard: log and fail loudly
        LOG.exception("FAILED: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
