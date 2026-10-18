"""
daily_master.py  -  ONE script that runs ALL your TB skills every day.

Run:   python daily_master.py              (all jobs)
       python daily_master.py dtc tpt      (only the jobs you name)

Jobs (in this order):
  notification  Notification registers (diagnosing + enrolment base)   -> Excel
  taluka        Beneficiary bank status + PMTBMBA ABHA seeding         -> Excel + small PPT
  tpt           TPT report                                             -> Excel
  ptfu          PTFU due-vs-done analysis                              -> Excel
  dtc           Differentiated care (DTC) report                       -> Excel
  review        District review deck + Monthly review deck             -> PowerPoint (needs Node.js)

Everything is saved in   <DATA_FOLDER>\\OUTPUT\\<today>\\
A summary of what worked / failed is written to  daily_master_log.txt  (next to this script).
A failed job never stops the other jobs.
"""
import calendar
import os
import re
import shutil
import subprocess
import sys
import time
import traceback
from datetime import date, datetime, timedelta

# ----------------------------------------------------------------------
# 1. SETTINGS  (the only part you normally edit)
# ----------------------------------------------------------------------
DATA_FOLDER = os.environ.get("TB_DATA_FOLDER") or r"C:\Users\ADMIN\OneDrive\Desktop\FINAL M LANGUAGE PREPARED"

# Exact file names (case does not matter) that each job reads from DATA_FOLDER
FILES = {
    "notification": "NOTIFICATION.xlsx",
    "tpt": "tb preventve therapy master.xlsx",
    "ptfu": "PTFU CUMULATVE.xlsx",
    "dtc": "DTC CUMULATIVE.xlsx",
}

# Which jobs run when you do not name any. Remove a name to switch a job off.
DEFAULT_JOBS = ["notification", "taluka", "tpt", "ptfu", "dtc", "review"]

MAKE_MONTHLY_DECK = False   # True = also make the short 13-slide Monthly Review deck (different format)
PTFU_MONTHS_BACK = 4        # PTFU report covers the current month + the 3 before it
WARN_IF_FILE_OLDER_THAN_DAYS = 3

# ----------------------------------------------------------------------
# 2. HELPERS
# ----------------------------------------------------------------------
HERE = os.path.dirname(os.path.abspath(__file__))
TODAY = date.today()
OUT_DIR = os.path.join(DATA_FOLDER, "OUTPUT", f"{TODAY:%Y-%m-%d}")
WORK_DIR = os.path.join(DATA_FOLDER, "OUTPUT", "_work")
STAGE_DIR = os.path.join(WORK_DIR, "staging")
LOG = []

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def log(msg=""):
    line = f"{datetime.now():%H:%M:%S}  {msg}"
    print(line, flush=True)
    LOG.append(line)


def child_env(extra=None):
    env = dict(os.environ)
    env["PYTHONUTF8"] = "1"                 # stops Windows 'charmap' errors with Unicode text
    env["PYTHONIOENCODING"] = "utf-8"
    env["TB_DATA_FOLDER"] = DATA_FOLDER
    if extra:
        env.update(extra)
    return env


def run(cmd, timeout=1800, extra_env=None, shell=False):
    """Run a command, return (ok, combined_output)."""
    p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace",
                       timeout=timeout, env=child_env(extra_env), shell=shell)
    out = (p.stdout or "") + (p.stderr or "")
    return p.returncode == 0, out.strip()


def tail(text, n=12):
    lines = [l for l in text.splitlines() if l.strip()]
    return "\n".join("      " + l for l in lines[-n:])


def find_input(name):
    """Case-insensitive lookup of a file in DATA_FOLDER (top level)."""
    want = name.lower()
    try:
        for f in os.listdir(DATA_FOLDER):
            if f.lower() == want:
                return os.path.join(DATA_FOLDER, f)
    except FileNotFoundError:
        pass
    return None


def script(name):
    p = os.path.join(HERE, name)
    if not os.path.isfile(p):
        raise FileNotFoundError(f"script file missing next to daily_master.py: {name}")
    return p


def month_tok(d):
    return d.strftime("%b-%y")


def prev_month_first(d):
    return (d.replace(day=1) - timedelta(days=1)).replace(day=1)


def sanitize(name):
    """Claude-style upload names (spaces/commas -> underscore) that the review skill expects."""
    return re.sub(r"[^A-Za-z0-9._-]", "_", name)


def build_staging():
    """Copy every .xlsx in DATA_FOLDER to STAGE_DIR with underscore names (only if changed)."""
    os.makedirs(STAGE_DIR, exist_ok=True)
    n = 0
    for f in os.listdir(DATA_FOLDER):
        src = os.path.join(DATA_FOLDER, f)
        if not (os.path.isfile(src) and f.lower().endswith(".xlsx")) or f.startswith("~$"):
            continue
        dst = os.path.join(STAGE_DIR, sanitize(f))
        if (not os.path.exists(dst) or os.path.getsize(dst) != os.path.getsize(src)
                or os.path.getmtime(dst) < os.path.getmtime(src)):
            shutil.copy2(src, dst)
        n += 1
    return n


# ----------------------------------------------------------------------
# 3. JOBS  (each returns a short result text, or raises SkipJob / an error)
# ----------------------------------------------------------------------
class SkipJob(Exception):
    pass


def need(key):
    p = find_input(FILES[key])
    if not p:
        raise SkipJob(f'input file not found in the data folder: "{FILES[key]}"')
    return p


def job_notification():
    src = need("notification")
    out = os.path.join(OUT_DIR, "Notification_Registers")
    os.makedirs(out, exist_ok=True)
    ok, txt = run([sys.executable, script("build_notification_registers.py"), "--input", src, "--outdir", out])
    if not ok:
        raise RuntimeError(txt)
    # the review skill reads fresh copies of these two registers
    os.makedirs(STAGE_DIR, exist_ok=True)
    for f in os.listdir(out):
        if f.startswith("Notification_Register_") and f.endswith(".xlsx"):
            shutil.copy2(os.path.join(out, f), os.path.join(STAGE_DIR, f))
    return f"saved in {out}"


def job_taluka():
    ok, txt = run([sys.executable, script("daily_tb_report.py")])
    if not ok:
        raise RuntimeError(txt)
    if "ERROR" in txt or "SKIPPED" in txt:
        log(tail(txt))
        return "finished with warnings (see lines above)"
    return f"saved in {OUT_DIR}"


def job_tpt():
    src = need("tpt")
    last = calendar.monthrange(TODAY.year, TODAY.month)[1]
    ok, txt = run([sys.executable, script("build_tpt_report.py"), "--input", src,
                   "--start", f"{TODAY.year}-01-01", "--end", f"{TODAY.year}-{TODAY.month:02d}-{last:02d}",
                   "--outdir", OUT_DIR])
    if not ok:
        raise RuntimeError(txt)
    return f"saved in {OUT_DIR}"


def job_ptfu():
    src = need("ptfu")
    months, d = [], TODAY.replace(day=1)
    for _ in range(PTFU_MONTHS_BACK):
        months.append(month_tok(d))
        d = prev_month_first(d)
    months = months[::-1]
    out = os.path.join(OUT_DIR, f"PTFU_Analysis_{months[0].replace('-', '')}_to_{months[-1].replace('-', '')}.xlsx")
    cmd = [sys.executable, script("generate_ptfu_report.py"), "--source", src,
           "--months", ",".join(months), "--output", out]
    ok, txt = run(cmd)
    if not ok:                                   # the skill notes say: retry once
        log("      first attempt failed, retrying once ...")
        ok, txt = run(cmd)
    if not ok:
        raise RuntimeError(txt)
    return f"months {months[0]} to {months[-1]} -> {os.path.basename(out)}"


def job_dtc():
    src = need("dtc")
    if TODAY.month == 1:
        y, last_m = TODAY.year - 1, 12
    else:
        y, last_m = TODAY.year, TODAY.month - 1
    months = ",".join(date(y, m, 1).strftime("%b-%y") for m in range(1, last_m + 1))
    out = os.path.join(OUT_DIR, f"DTC_Report_{y}_Jan_to_{date(y, last_m, 1):%b}.xlsx")
    ok, txt = run([sys.executable, script("build_dtc_report.py"), "--source", src,
                   "--months", months, "--output", out])
    if not ok:
        raise RuntimeError(txt)
    return f"Jan-{y} to {date(y, last_m, 1):%b}-{y} -> {os.path.basename(out)}"


def node_setup():
    """Return NODE_PATH for pptxgenjs, or raise SkipJob with install instructions."""
    node = shutil.which("node")
    if not node:
        raise SkipJob("Node.js is not installed (needed only for the PowerPoint decks). "
                      "Install the LTS version from https://nodejs.org then run:  npm install -g pptxgenjs")
    ok, root = run("npm root -g", timeout=60, shell=True)
    root = root.splitlines()[-1].strip() if ok and root else ""
    if not root or not os.path.isdir(os.path.join(root, "pptxgenjs")):
        raise SkipJob("pptxgenjs is not installed. Open Command Prompt and run:  npm install -g pptxgenjs")
    return node, root


REVIEW_REQUIRED = [
    "BENIFICIARY REGISTER.xlsx",
    "CO MORBIDITY REGISTER FOR PENIDNG HIV AND DIABETS .xlsx",
    "CURRENT NOTIFICATION REGISTER FOR TREATMENT NOT INITED AND OUTCOME PENIDNGR.xlsx",
    "DTC CUMULATIVE.xlsx",
    "NOTIFICATION.xlsx",
    "NOTIFICATION TARGET TU WISE MAHISAGAR DISTRICT YEAR 2026.xlsx",
    "PATIENT LAB REGI IP AND CP PENIDNG LIST.xlsx",
    "PMTBMBA REGISTER FOR ABHASEEDING AND CONSENT FOR KIT.xlsx",
    "PRESUMPTIVE REGISTER FOR XRAY , NAAT, MICROSCOPY.xlsx",
    "PTFU CUMULATVE.xlsx",
    "cbnaat target.xlsx",
    "monthly notification target.xlsx",
    "tb preventve therapy master.xlsx",
]


def job_review():
    node, node_root = node_setup()
    n = build_staging()
    log(f"      staged {n} Excel files with underscore names")
    missing = [f for f in REVIEW_REQUIRED if not os.path.exists(os.path.join(STAGE_DIR, sanitize(f)))]
    if missing:
        raise SkipJob("these files are missing from the data folder: " + "; ".join(missing))
    # the notification job (if it ran today) already placed fresh Notification_Register_* files in staging
    if not any(f.startswith("Notification_Register_") for f in os.listdir(STAGE_DIR)):
        raise SkipJob("Notification_Register_* files missing - run the 'notification' job first")

    os.makedirs(WORK_DIR, exist_ok=True)
    r_json = os.path.join(WORK_DIR, "R.json")
    m_json = os.path.join(WORK_DIR, "monthly.json")
    env = {"NODE_PATH": node_root}
    results = []

    log("      extracting district indicators (this takes several minutes) ...")
    ok, txt = run([sys.executable, script("extract_data.py"), STAGE_DIR, r_json, TODAY.isoformat()], timeout=3600)
    if not ok:
        raise RuntimeError("extract_data.py failed:\n" + tail(txt, 25))

    district = os.path.join(OUT_DIR, f"Mahisagar_District_TB_Performance_Review_{TODAY.year}_{TODAY:%d-%m-%Y}.pptx")
    ok, txt = run([node, script("build_district_deck.js"), r_json, district], extra_env=env)
    if not ok:
        raise RuntimeError("district deck failed:\n" + tail(txt, 25))
    results.append(os.path.basename(district))

    if not MAKE_MONTHLY_DECK:
        return "made " + results[0] + "  (monthly deck is switched off)"

    review_month = prev_month_first(TODAY)
    log(f"      monthly review for {review_month:%B %Y} ...")
    ok, txt = run([sys.executable, script("extract_monthly.py"), STAGE_DIR, m_json,
                   "--month", review_month.strftime("%Y-%m")], timeout=3600)
    if not ok:
        raise RuntimeError("extract_monthly.py failed:\n" + tail(txt, 25))
    monthly = os.path.join(OUT_DIR, f"Monthly_Review_{review_month:%b%Y}.pptx")
    ok, txt = run([node, script("build_monthly_deck.js"), m_json, r_json, monthly], extra_env=env)
    if not ok:
        raise RuntimeError("monthly deck failed:\n" + tail(txt, 25))
    results.append(os.path.basename(monthly))
    return "made " + " and ".join(results)


JOBS = {
    "notification": job_notification,
    "taluka": job_taluka,
    "tpt": job_tpt,
    "ptfu": job_ptfu,
    "dtc": job_dtc,
    "review": job_review,
}


# ----------------------------------------------------------------------
# 4. FRESHNESS CHECK + MAIN
# ----------------------------------------------------------------------
def freshness_report():
    log("Checking how fresh your data files are:")
    stale = 0
    for key, name in FILES.items():
        p = find_input(name)
        if not p:
            log(f"   MISSING   {name}")
            continue
        age = (datetime.now() - datetime.fromtimestamp(os.path.getmtime(p))).days
        flag = "OLD " if age >= WARN_IF_FILE_OLDER_THAN_DAYS else "ok  "
        stale += flag == "OLD "
        log(f"   {flag}      {name}  (last updated {age} day(s) ago)")
    if stale:
        log(f"   NOTE: {stale} file(s) were not updated recently - the reports use whatever is in the folder.")


def main(argv):
    wanted = [a.lower() for a in argv] or DEFAULT_JOBS
    unknown = [w for w in wanted if w not in JOBS]
    if unknown:
        print(f"Unknown job(s): {unknown}. Choose from: {list(JOBS)}")
        return 2

    log(f"Data folder: {DATA_FOLDER}")
    if not os.path.isdir(DATA_FOLDER):
        log("ERROR: data folder not found. Fix DATA_FOLDER at the top of daily_master.py")
        return 1
    os.makedirs(OUT_DIR, exist_ok=True)
    freshness_report()

    summary = []
    for name in [j for j in JOBS if j in wanted]:      # fixed order, whatever order you typed
        log("")
        log(f"=== {name.upper()} ===")
        t0 = time.time()
        try:
            msg = JOBS[name]()
            status = "OK"
        except SkipJob as e:
            status, msg = "SKIPPED", str(e)
        except subprocess.TimeoutExpired:
            status, msg = "FAILED", "took too long and was stopped"
        except Exception as e:
            status, msg = "FAILED", str(e)
            if not isinstance(e, RuntimeError):
                log(traceback.format_exc())
        took = time.time() - t0
        if status == "FAILED":
            log(f"   FAILED after {took:.0f}s. Last lines of the error:")
            log(tail(msg, 10))
        else:
            log(f"   {status}: {msg}   ({took:.0f}s)")
        summary.append((name, status, msg, took))

    log("")
    log("================ SUMMARY ================")
    for name, status, msg, took in summary:
        lines = [l.strip() for l in msg.splitlines() if l.strip()]
        shown = (lines[-1] if status == "FAILED" else lines[0]) if lines else ""
        log(f"{status:8s} {name:13s} {shown}")
    log(f"Output folder: {OUT_DIR}")
    return 0 if all(s[1] != "FAILED" for s in summary) else 1


if __name__ == "__main__":
    code = 1
    try:
        code = main(sys.argv[1:])
    except Exception:
        log("UNEXPECTED ERROR:\n" + traceback.format_exc())
    try:
        with open(os.path.join(HERE, "daily_master_log.txt"), "a", encoding="utf-8") as f:
            f.write("\n".join(LOG) + "\n" + "-" * 70 + "\n")
        os.makedirs(OUT_DIR, exist_ok=True)
        with open(os.path.join(OUT_DIR, "_run_summary.txt"), "w", encoding="utf-8") as f:
            f.write("\n".join(LOG))
    except Exception:
        pass
    sys.exit(code)
