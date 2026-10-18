"""Render title / explainer / terminal / result cards (1280x720) for the two demo videos."""
import os
import textwrap

from PIL import Image, ImageDraw, ImageFont

W, H = 1280, 720
NAVY = (14, 76, 129)
DARK = (16, 28, 40)
WHITE = (255, 255, 255)
LIGHT = (200, 220, 235)
GREEN = (120, 220, 140)
AMBER = (255, 200, 90)
GREY = (140, 155, 170)

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "video_assets")
os.makedirs(OUT, exist_ok=True)


def font(size, bold=False):
    for p in ("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold
              else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
              "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"):
        try:
            return ImageFont.truetype(p, size)
        except OSError:
            continue
    return ImageFont.load_default()


def mono(size):
    try:
        return ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf", size)
    except OSError:
        return ImageFont.load_default()


def base(bg=NAVY):
    img = Image.new("RGB", (W, H), bg)
    d = ImageDraw.Draw(img)
    d.rectangle([0, H - 8, W, H], fill=(255, 193, 7))
    return img, d


def centered(d, y, text, f, fill=WHITE):
    w = d.textlength(text, font=f)
    d.text(((W - w) / 2, y), text, font=f, fill=fill)


def wrap_card(name, title, lines, footer, bg=NAVY, title_size=54):
    img, d = base(bg)
    centered(d, 110, title, font(title_size, True))
    y = 240
    for line in lines:
        centered(d, y, line, font(30), LIGHT)
        y += 52
    centered(d, H - 110, footer, font(22), GREY)
    img.save(os.path.join(OUT, name))
    print("wrote", name)


def terminal_card(name, title, log_lines, footer):
    img, d = base(DARK)
    d.text((70, 60), title, font=font(34, True), fill=WHITE)
    d.rectangle([60, 130, W - 60, H - 130], outline=(60, 80, 100), width=2)
    y = 160
    for line in log_lines:
        d.text((90, y), line, font=mono(24), fill=GREEN if line.startswith("[INFO]") else WHITE)
        y += 42
    d.text((70, H - 90), footer, font=font(22), fill=GREY)
    img.save(os.path.join(OUT, name))
    print("wrote", name)


# ---- Video 1 cards ----
wrap_card("v1_title.png", "TASK DASHBOARD", [
    "One click turns a flat Excel task export",
    "into a live project dashboard.",
], "VBA macro  ·  by Muhammad Bilal  ·  100% local — no network, no API keys")

wrap_card("v1_how.png", "HOW IT WORKS", [
    "Task Export  →  BuildDashboard macro  →  Dashboard sheet",
    "KPI tiles  ·  Tasks by Owner  ·  Tasks by Status",
    "Overdue list in red  ·  Due in 7 days in amber",
    "Re-run any time — the Dashboard rebuilds from current data.",
], "Excel is the intended host · macro also runs via LibreOffice UNO")

terminal_card("v1_terminal.png", "$ python3 run_macro.py", [
    "[INFO] Macro-built workbook saved as _tmp_after_macro.xlsx",
    "[INFO] vbaProject.bin built (10240 bytes)",
    "[INFO] olevba: embedded macro is BYTE-IDENTICAL to TaskDashboard.bas",
    "[INFO] Dashboard KPIs — total=40 completed=12 in_progress=14",
    "[INFO]   overdue=9 completion=30.0%",
    "[INFO] Dashboard verification PASSED",
    "[INFO] DONE — Task-Dashboard.xlsm ready",
], "Real run output — not a mockup")

wrap_card("v1_result.png", "RESULT", [
    "40 tasks  →  1 click  →  live dashboard",
    "9 overdue flagged red  ·  7 due soon flagged amber  ·  30% complete",
    "No manual counting. No stale numbers.",
], "Task-Dashboard.xlsm  ·  TaskDashboard.bas  ·  by Muhammad Bilal")

# ---- Video 2 cards ----
wrap_card("v2_title.png", "WEEKLY REPORT GENERATOR", [
    "The Monday status pack, built automatically.",
], "Python + openpyxl  ·  by Muhammad Bilal  ·  100% local — no network, no API keys")

wrap_card("v2_how.png", "HOW IT WORKS", [
    "Task-Tracker.xlsx  →  weekly_report.py  →  Weekly-Report-YYYY-MM-DD.xlsx + .pdf",
    "Summary  ·  Overdue  ·  Due This Week  ·  Workload by Owner",
    "Auto-trigger: Windows Task Scheduler / cron — every Monday, 8:00 AM",
    "Idempotent: same date → same files, overwritten — never duplicates.",
], "PDF via LibreOffice headless · logs to weekly-report.log")

terminal_card("v2_terminal.png", "$ python3 weekly_report.py", [
    "[INFO] Loaded 30 tasks from Task-Tracker.xlsx",
    "[INFO] Wrote output/Weekly-Report-2026-10-16.xlsx",
    "[INFO] KPIs: total=30 completed=6 overdue=8",
    "[INFO]   due_this_week=7 completion=20%",
    "[INFO] Wrote output/Weekly-Report-2026-10-16.pdf",
    "[INFO] DONE",
], "Real run output — not a mockup")

wrap_card("v2_result.png", "RESULT", [
    "30 tasks  →  1 command  →  Excel + PDF",
    "8 overdue  ·  7 due this week  ·  20% complete",
    "Every Monday at 8:00 AM, automatically.",
], "run_weekly_report.bat / .sh  ·  task-scheduler.xml  ·  by Muhammad Bilal")
