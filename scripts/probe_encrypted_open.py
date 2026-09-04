"""Probe: what does Workbooks.Open do on an encrypted file with a wrong
supplied password? PID-precise, self-cleaning, Excel must be closed."""
import os
import sys
import tempfile
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from xlsx_mcp.com.instances import list_excel_pids, pid_alive, taskkill  # noqa: E402

base = list_excel_pids()
if base:
    print("SKIP: excel running", base)
    sys.exit(0)

scratch = Path(tempfile.mkdtemp(prefix="ks4xl_probe_"))
p = scratch / "enc.xlsx"

import openpyxl  # noqa: E402
wb = openpyxl.Workbook()
wb.active["A1"] = "x"
wb.save(p)
wb.close()

import win32com.client as win32  # noqa: E402
app = win32.DispatchEx("Excel.Application")
app.Visible = False
app.DisplayAlerts = False
pid = (list_excel_pids() - base).pop()
print("worker pid", pid)

# encrypt
wbx = app.Workbooks.Open(str(p))
wbx.Password = "hunter2"
wbx.Save()
wbx.Close(SaveChanges=False)
print("encrypted")

results = {}


def try_open(label, **kw):
    """Main-thread open with a watchdog that taskkills the OWNED pid on a
    hang (a modal dialog blocks the COM call unboundedly)."""
    t0 = time.time()
    done = threading.Event()
    killed = {"yes": False}

    def watchdog():
        if not done.wait(10):
            killed["yes"] = True
            taskkill(pid)

    th = threading.Thread(target=watchdog, daemon=True)
    th.start()
    try:
        b = app.Workbooks.Open(str(p), **kw)
        results[label] = ("OPENED", None)
        b.Close(SaveChanges=False)
    except Exception as exc:  # noqa: BLE001
        tag = "HANG->killed" if killed["yes"] else "ERROR"
        results[label] = (tag, repr(exc)[:200])
    done.set()
    print(label, results[label], f"{time.time()-t0:.1f}s")
    return not killed["yes"]


def try_open_pos(label, *args):
    t0 = time.time()
    done = threading.Event()
    killed = {"yes": False}

    def watchdog():
        if not done.wait(10):
            killed["yes"] = True
            taskkill(pid)

    th = threading.Thread(target=watchdog, daemon=True)
    th.start()
    try:
        b = app.Workbooks.Open(str(p), *args)
        results[label] = ("OPENED", None)
        b.Close(SaveChanges=False)
    except Exception as exc:  # noqa: BLE001
        tag = "HANG->killed" if killed["yes"] else "ERROR"
        results[label] = (tag, repr(exc)[:200])
    done.set()
    print(label, results[label], f"{time.time()-t0:.1f}s")
    return not killed["yes"]


# Positional: Open(FileName, UpdateLinks, ReadOnly, Format, Password)
ok = try_open_pos("correct-pw-positional", 0, False, None, "hunter2")
if ok:
    ok = try_open_pos("wrong-pw-positional", 0, False, None, "nope")
if ok:
    ok = try_open("correct-password-named", Password="hunter2")

# cleanup by owned pid
try:
    app.Quit()
except Exception:
    pass
time.sleep(1)
if pid_alive(pid):
    taskkill(pid)
for _ in range(30):
    if not pid_alive(pid):
        break
    time.sleep(1)
time.sleep(3)
leftover = list_excel_pids() - base
for extra in list(leftover):
    print("post-kill transient EXCEL:", extra, "alive:", pid_alive(extra))
    taskkill(extra)  # spawned by this probe's kill; ours to reclaim
print("final pids:", sorted(list_excel_pids() - base))
