@echo off
REM Weekly Report Generator — Windows launcher (by Muhammad Bilal)
REM Double-click to generate this week's report, or let Task Scheduler run it.
REM Requires: Python 3 + openpyxl (pip install -r requirements.txt).
REM LibreOffice (soffice) in PATH is needed for the PDF; without it, only .xlsx is produced.
cd /d "%~dp0"
python weekly_report.py >> weekly-report.log 2>&1
if errorlevel 1 (
  echo [ERROR] weekly_report.py failed. See weekly-report.log for details.
  exit /b 1
)
echo Weekly report generated. See the output\ folder.
