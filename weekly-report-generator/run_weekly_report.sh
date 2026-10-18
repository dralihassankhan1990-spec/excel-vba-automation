#!/usr/bin/env bash
# Weekly Report Generator — Linux/mac launcher (by Muhammad Bilal)
# Run manually:  ./run_weekly_report.sh
# Or schedule with cron (see README.md).
# Requires: Python 3 + openpyxl (pip install -r requirements.txt).
# LibreOffice (soffice) in PATH is needed for the PDF; without it, only .xlsx is produced.
set -euo pipefail
cd "$(dirname "$0")"
python3 weekly_report.py >> weekly-report.log 2>&1
echo "Weekly report generated. See the output/ folder."
