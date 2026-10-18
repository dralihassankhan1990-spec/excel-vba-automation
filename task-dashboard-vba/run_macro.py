"""Full build + proof pipeline for Task-Dashboard.xlsm.

1. build_workbook.py  -> Task-Dashboard-base.xlsx (Task Export sheet, 40 tasks)
2. LibreOffice UNO: open the base workbook, inject the VBA source
   (TaskDashboard.bas) as a Basic module, execute BuildDashboard, and verify
   the Dashboard KPIs in memory. This PROVES the macro runs and is correct.
3. Save the macro-built workbook from LibreOffice as .xlsx, then inject the
   hand-built vbaProject.bin (build_vba_project.py, MS-OVBA + MS-CFB) into it
   -> final Task-Dashboard.xlsm containing BOTH the embedded macro AND the
   Dashboard sheet the macro itself produced.
4. Verification: oletools.olevba extracts the embedded macro from the final
   .xlsm and it is compared byte-for-byte with TaskDashboard.bas; LibreOffice
   re-opens the final file and confirms both sheets are intact.

Why the two-step dance: LibreOffice can RUN VBA macros but cannot AUTHOR
Excel VBA projects (it never exports Basic macros as vbaProject.bin, and its
VBA import did not load hand-built projects in this environment). The macro
logic is therefore proven by execution (step 2) and the embedded project is
proven byte-identical to that proven source (step 4). Excel is the intended
host for the delivered .xlsm; see TEST-REPORT.md for the honest limits.

Usage: python3 run_macro.py [--keep-soffice]
Requires: LibreOffice + python3-uno + oletools. 100% local. by Muhammad Bilal.
"""
import argparse
import logging
import os
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue

HERE = os.path.dirname(os.path.abspath(__file__))
BASE_XLSX = os.path.join(HERE, "Task-Dashboard-base.xlsx")
TMP_XLSX = os.path.join(HERE, "_tmp_after_macro.xlsx")
OUT_XLSM = os.path.join(HERE, "Task-Dashboard.xlsm")
BAS_FILE = os.path.join(HERE, "TaskDashboard.bas")
PORT = 2002

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("run_macro")


def pv(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def start_soffice():
    proc = subprocess.Popen(
        ["soffice", "--headless", "--nologo", "--norestore",
         f"--accept=socket,host=localhost,port={PORT};urp;StarOffice.ServiceManager",
         "--nofirststartwizard"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    log.info("soffice started (pid %s), waiting for UNO socket...", proc.pid)
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx)
    for _ in range(60):
        try:
            ctx = resolver.resolve(
                f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext")
            log.info("UNO bridge connected.")
            return proc, ctx
        except Exception:
            time.sleep(1)
    raise RuntimeError("Could not connect to LibreOffice UNO bridge")


def verify_dashboard(doc):
    sheets = doc.getSheets()
    assert sheets.hasByName("Dashboard"), "Dashboard sheet was not created!"
    dash = sheets.getByName("Dashboard")
    title = dash.getCellByPosition(0, 0).getString()
    assert title == "TASK DASHBOARD", f"Unexpected title: {title!r}"
    total = int(dash.getCellByPosition(0, 4).getValue())
    completed = int(dash.getCellByPosition(2, 4).getValue())
    inprog = int(dash.getCellByPosition(4, 4).getValue())
    overdue = int(dash.getCellByPosition(6, 4).getValue())
    pct = dash.getCellByPosition(8, 4).getValue()
    log.info("Dashboard KPIs — total=%d completed=%d in_progress=%d overdue=%d completion=%.1f%%",
             total, completed, inprog, overdue, pct * 100)
    assert (total, completed, inprog, overdue) == (40, 12, 14, 9), "KPI mismatch"
    assert abs(pct - 0.30) < 0.001, f"expected 30% completion, got {pct}"
    found = any("OVERDUE TASKS" in dash.getCellByPosition(0, r).getString()
                for r in range(0, 60))
    assert found, "OVERDUE TASKS section not found on Dashboard"
    log.info("Dashboard verification PASSED (KPIs + overdue section).")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--keep-soffice", action="store_true")
    args = ap.parse_args()

    # 1. base workbook
    log.info("Step 1: building base workbook...")
    r = subprocess.run([sys.executable, os.path.join(HERE, "build_workbook.py")],
                       cwd=HERE, capture_output=True, text=True)
    sys.stdout.write(r.stdout)
    if r.returncode != 0:
        log.error("build_workbook.py failed:\n%s", r.stderr)
        return 1

    # 2. run the macro via UNO (injected Basic module = proven execution path)
    with open(BAS_FILE, encoding="utf-8") as f:
        vba = f.read()
    vba_lines = [ln for ln in vba.splitlines() if not ln.startswith("Attribute ")]
    module_code = "Option VBASupport 1\n" + "\n".join(vba_lines)
    log.info("Step 2: running BuildDashboard via UNO (VBA source: %d lines)...", len(vba_lines))

    proc, ctx = start_soffice()
    try:
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(BASE_XLSX), "_blank", 0, ())
        libs = doc.BasicLibraries
        if not libs.hasByName("Standard"):
            libs.createLibrary("Standard")
        std = libs.getByName("Standard")
        if std.hasByName("TaskDashboard"):
            std.removeByName("TaskDashboard")
        std.insertByName("TaskDashboard", module_code)
        sp = smgr.createInstanceWithContext(
            "com.sun.star.script.provider.ScriptProviderForBasic", ctx)
        sp.initialize((doc,))
        script = sp.getScript(
            "vnd.sun.star.script:Standard.TaskDashboard.BuildDashboard"
            "?language=Basic&location=document")
        script.invoke((), (), ())
        log.info("BuildDashboard returned without error.")
        verify_dashboard(doc)
        if os.path.exists(TMP_XLSX):
            os.remove(TMP_XLSX)
        doc.storeAsURL(uno.systemPathToFileUrl(TMP_XLSX),
                       (pv("FilterName", "Calc MS Excel 2007 XML"), pv("Overwrite", True)))
        log.info("Macro-built workbook saved as %s.", TMP_XLSX)
        doc.close(True)
    finally:
        if not args.keep_soffice:
            proc.terminate()
            try:
                proc.wait(timeout=10)
            except subprocess.TimeoutExpired:
                proc.kill()
            log.info("soffice stopped.")

    # 3. inject the hand-built VBA project -> final .xlsm
    log.info("Step 3: building vbaProject.bin + injecting into macro-built workbook...")
    sys.path.insert(0, HERE)
    from build_vba_project import build_vba_project_bin, inject_into_xlsm
    vba_bin = build_vba_project_bin(BAS_FILE)
    log.info("vbaProject.bin built (%d bytes).", len(vba_bin))
    if os.path.exists(OUT_XLSM):
        os.remove(OUT_XLSM)
    inject_into_xlsm(TMP_XLSX, vba_bin, OUT_XLSM)
    log.info("Final %s (Dashboard + embedded macro).", OUT_XLSM)

    # 4. verify: olevba byte-identical extraction + LO re-open confirms sheets
    log.info("Step 4: verifying delivered .xlsm...")
    from oletools.olevba import VBA_Parser
    p = VBA_Parser(OUT_XLSM)
    assert p.detect_vba_macros(), "olevba: no macros detected in delivered .xlsm!"
    mods = list(p.extract_macros())
    std_mods = [m for m in mods if m[2] == "TaskDashboard.bas"]
    assert len(std_mods) == 1, f"expected TaskDashboard module, got {[m[2] for m in mods]}"
    expected = "\r\n".join(vba_lines) + "\r\n"
    assert std_mods[0][3] == expected, "extracted macro differs from TaskDashboard.bas!"
    log.info("olevba: embedded macro is BYTE-IDENTICAL to TaskDashboard.bas.")

    proc, ctx = start_soffice()
    try:
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(OUT_XLSM), "_blank", 0, ())
        sheets = doc.getSheets()
        assert sheets.hasByName("Dashboard") and sheets.hasByName("Task Export"), \
            "sheets missing from delivered .xlsm"
        verify_dashboard(doc)
        doc.close(True)
        log.info("Delivered .xlsm re-opened: both sheets intact, Dashboard KPIs hold.")
    finally:
        if not args.keep_soffice:
            proc.terminate()
            try:
                proc.wait(timeout=10)
            except subprocess.TimeoutExpired:
                proc.kill()

    if os.path.exists(TMP_XLSX):
        os.remove(TMP_XLSX)
    log.info("DONE — %s ready.", OUT_XLSM)
    return 0


if __name__ == "__main__":
    sys.exit(main())
