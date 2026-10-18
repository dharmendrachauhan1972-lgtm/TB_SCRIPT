"""
daily_tb_report.py  -  run this once a day (Task Scheduler at 10:00 AM).

What it does
1. Looks in your data folder and picks the NEWEST file of each type
   (Beneficiary register, PMTBMBA register, HF register).
2. Runs your skill's pivot engine (hf_link_pivot.py) for both reports
   -> two Excel files (taluka-wise tables + line lists).
3. Builds one PowerPoint summarising the same numbers.
4. Saves everything in  <data folder>\\OUTPUT\\<today's date>\\
5. Writes a log file so you can see what happened.

Keep this file and hf_link_pivot.py in the SAME folder.
"""
import glob
import os
import subprocess
import sys
import traceback
from collections import Counter, defaultdict
from datetime import datetime

import openpyxl
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

# ----------------------------------------------------------------------
# 1. SETTINGS  (the only part you normally edit)
# ----------------------------------------------------------------------
DATA_FOLDER = os.environ.get("TB_DATA_FOLDER") or r"C:\Users\ADMIN\OneDrive\Desktop\FINAL M LANGUAGE PREPARED"

# Business rule from your skill: Kadana has no private facilities,
# so its private-sector patients are counted under Lunawada.
NONFUNCTIONAL_TALUKA = "KADANA"
REASSIGN_TO = "LUNAWADA"
REASSIGN_SECTOR = "Private"

# Words that must appear in the file name (case does not matter).
# Excel temp files (~$...) and the OUTPUT folder are ignored automatically.
FILE_KEYWORDS = {
    "beneficiary": ["benif"],        # matches BENIFICIARY_... or BENIFICIARY ...
    "pmtbmba": ["pmtbmba"],
    "hf": ["hf", "register"],        # HF_REGISTER ...
}

REPORTS = [
    {
        "key": "beneficiary",
        "register_sheet": "BENIFICIARY REGISTER",
        "taluka_col": "CurrentFacilityTBU",
        "facility_col": "CurrentFacilityPHI",
        "type_col": "TypeOfPatient",
        "public_value": "IndiaTbPublic",
        "status_col": "Beneficiary_Status",
        "status_order": ["EMPTY", "ENTERED", "SENT", "NOT_VALIDATED",
                         "VALIDATED", "REJECTED", "USER_REJECTED"],
        "highlight_status": "VALIDATED",
        "title": "TALUKA-WISE BENEFICIARY BANK DETAIL STATUS - PUBLIC & PRIVATE HEALTH FACILITIES",
        "ppt_title": "Beneficiary Bank Detail Status (Taluka-wise)",
        "sheet_name": "TALUKA WISE BANK STATUS",
        "out_name": "Beneficiary_Bank_Status_Taluka_Wise",
        "linelist": ("ForegoBenefits=Yes:FOREGO BENEFITS - YES LIST:"
                     "EpisodeId,PatientName,Age,Gender,PrimaryPhone,Stage,TypeOfCase,"
                     "Beneficiary_Status,NoOfBenefitEligible,TotalBenefitAmountEligible,"
                     "NoOfBenefitPaid_Nikshay,TotalAmountPaid_Nikshay"),
    },
    {
        "key": "pmtbmba",
        "register_sheet": "PMTBMBA REGISTER",
        "taluka_col": "Spectrum_Current_TBU",
        "facility_col": "Spectrum_Current_PHI",
        "type_col": "Current_Spectrum_HF_Type",
        "public_value": "Public",
        "status_col": "ABHA generated/seeded",
        "status_order": ["Yes", "No"],
        "highlight_status": "Yes",
        "title": "TALUKA-WISE ABHA ID SEEDING STATUS - PUBLIC & PRIVATE HEALTH FACILITIES (PMTBMBA)",
        "ppt_title": "ABHA ID Seeding Status (Taluka-wise)",
        "sheet_name": "ABHA SEEDING STATUS",
        "out_name": "PMTBMBA_ABHA_Seeding_Status_Taluka_Wise",
        "linelist": ("ABHA generated/seeded=No:ABHA SEEDING - NOT DONE LIST:"
                     "Episode_ID,Patient_Name,Age,Gender,PrimaryPhone,"
                     "Spectrum_Diagnosis_Date,PatientConsent,Providedsupport"),
    },
]

# ----------------------------------------------------------------------
# 2. HELPERS
# ----------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PIVOT_SCRIPT = os.path.join(SCRIPT_DIR, "hf_link_pivot.py")
LOG_LINES = []


def log(msg):
    line = f"{datetime.now():%H:%M:%S}  {msg}"
    print(line)
    LOG_LINES.append(line)


def all_xlsx():
    """Every .xlsx in the data folder AND its subfolders (OUTPUT folder is skipped).
    Returns (full_path, lowercase relative path) pairs."""
    found = []
    for root, dirs, files in os.walk(DATA_FOLDER):
        dirs[:] = [d for d in dirs if d.upper() != "OUTPUT"]
        for f in files:
            if f.lower().endswith(".xlsx") and not f.startswith("~$"):
                full = os.path.join(root, f)
                rel = os.path.relpath(full, DATA_FOLDER).lower()
                found.append((full, rel))
    return found


def find_latest(keywords):
    """Newest .xlsx whose path (folder + file name) contains ALL keywords."""
    hits = [p for p, rel in all_xlsx() if all(k in rel for k in keywords)]
    return max(hits, key=os.path.getmtime) if hits else None


def find_hf_register():
    """HF register = path has 'hf' and 'register' but is not one of the patient registers.
    Works if the file is in the main folder or inside a folder called 'HF REGISTER'."""
    hits = []
    for p, rel in all_xlsx():
        if "pmtbmba" in rel or "benif" in rel or "benef" in rel:
            continue
        if all(k in rel for k in FILE_KEYWORDS["hf"]):
            hits.append(p)
    return max(hits, key=os.path.getmtime) if hits else None


def first_sheet_name(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    return wb.sheetnames[0]


def eff_taluka(taluka, ptype):
    t = str(taluka).strip().upper()
    if ptype == REASSIGN_SECTOR and t == NONFUNCTIONAL_TALUKA:
        return REASSIGN_TO
    return t


def load_rows(path, sheet):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    rows = ws.iter_rows(values_only=True)
    header = list(next(rows))
    idx = {h: i for i, h in enumerate(header)}
    data = [r for r in rows if any(v is not None for v in r)]
    return idx, data


def count_by_taluka(cfg, register_path):
    """Python-side counts (same rules as the Excel formulas) for the PPT."""
    idx, rows = load_rows(register_path, cfg["register_sheet"])
    counts = defaultdict(Counter)
    for r in rows:
        is_public = str(r[idx[cfg["type_col"]]]).strip().upper() == cfg["public_value"].upper()
        ptype = "Public" if is_public else "Private"
        tal = eff_taluka(r[idx[cfg["taluka_col"]]], ptype)
        status = str(r[idx[cfg["status_col"]]]).strip()
        counts[tal][status] += 1
    return counts, len(rows)


# ----------------------------------------------------------------------
# 3. POWERPOINT
# ----------------------------------------------------------------------
NAVY = RGBColor(0x1F, 0x4E, 0x78)
GREY = RGBColor(0x59, 0x59, 0x59)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PALE = RGBColor(0xD9, 0xE1, 0xF2)


def add_title(slide, text, sub=None):
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(0.8))
    p = tb.text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = NAVY
    if sub:
        tb2 = slide.shapes.add_textbox(Inches(0.5), Inches(1.0), Inches(12.3), Inches(0.5))
        p2 = tb2.text_frame.paragraphs[0]
        p2.text = sub
        p2.font.size = Pt(14)
        p2.font.color.rgb = GREY


def set_cell(cell, text, bold=False, fill=None, color=None, size=13, center=True):
    cell.text = str(text)
    para = cell.text_frame.paragraphs[0]
    para.font.size = Pt(size)
    para.font.bold = bold
    if center:
        para.alignment = 2  # center
    if color is not None:
        para.font.color.rgb = color
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill


def build_ppt(results, out_path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Title slide
    s = prs.slides.add_slide(blank)
    add_title(s, "Mahisagar District - TB Programme", f"Daily status report - {datetime.now():%d %B %Y}")
    tb = s.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(12), Inches(2))
    tf = tb.text_frame
    for i, res in enumerate(results):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"{res['cfg']['ppt_title']}: {res['total']} records"
        p.font.size = Pt(20)
        p.font.color.rgb = GREY

    # One slide per report
    for res in results:
        cfg, counts, total = res["cfg"], res["counts"], res["total"]
        statuses = cfg["status_order"]
        hi = cfg["highlight_status"]
        talukas = sorted(counts.keys())

        s = prs.slides.add_slide(blank)
        grand = Counter()
        for t in talukas:
            grand.update(counts[t])
        hi_pct = (grand[hi] / total * 100) if total else 0
        add_title(s, cfg["ppt_title"],
                  f"Total {total} records  |  {hi}: {grand[hi]} ({hi_pct:.1f}%)")

        ncols = 1 + len(statuses) + 2          # Taluka + statuses + Total + highlight %
        nrows = len(talukas) + 2               # header + talukas + grand total
        tbl = s.shapes.add_table(nrows, ncols, Inches(0.5), Inches(1.7),
                                 Inches(12.3), Inches(0.5 * nrows)).table
        heads = ["Taluka"] + statuses + ["Total", f"{hi} %"]
        first_w = 2.0
        other_w = (12.3 - first_w) / (ncols - 1)
        tbl.columns[0].width = Inches(first_w)
        for c in range(1, ncols):
            tbl.columns[c].width = Inches(other_w)
        head_size = 12 if ncols <= 6 else 10
        for c, h in enumerate(heads):
            set_cell(tbl.cell(0, c), h.replace("_", " "), bold=True, fill=NAVY,
                     color=WHITE, size=head_size)

        def fill_row(r, label, ctr, bold=False, fill=None):
            tot = sum(ctr.values())
            vals = [label] + [ctr[st] for st in statuses] + [tot,
                    f"{(ctr[hi] / tot * 100):.1f}%" if tot else "0.0%"]
            for c, v in enumerate(vals):
                set_cell(tbl.cell(r, c), v, bold=bold, fill=fill, center=(c != 0))

        for r, t in enumerate(talukas, start=1):
            fill_row(r, t, counts[t])
        fill_row(nrows - 1, "GRAND TOTAL", grand, bold=True, fill=PALE)

    prs.save(out_path)


# ----------------------------------------------------------------------
# 4. MAIN
# ----------------------------------------------------------------------
def main():
    log(f"Data folder: {DATA_FOLDER}")
    if not os.path.isdir(DATA_FOLDER):
        log("ERROR: data folder not found. Check DATA_FOLDER at the top of the script.")
        return 1

    day_folder = os.path.join(DATA_FOLDER, "OUTPUT", f"{datetime.now():%Y-%m-%d}")
    os.makedirs(day_folder, exist_ok=True)

    hf_path = find_hf_register()
    if not hf_path:
        log("ERROR: HF register file not found (name should contain 'HF' and 'REGISTER').")
        return 1
    log(f"HF register : {os.path.relpath(hf_path, DATA_FOLDER)}")

    results = []
    for cfg in REPORTS:
        reg_path = find_latest(FILE_KEYWORDS[cfg["key"]])
        if not reg_path:
            log(f"SKIPPED {cfg['key']}: no matching .xlsx file found in the folder.")
            continue
        log(f"{cfg['key']:12s}: {os.path.relpath(reg_path, DATA_FOLDER)}")

        out_xlsx = os.path.join(day_folder, f"{cfg['out_name']}_{datetime.now():%d-%m-%Y}.xlsx")
        cmd = [
            sys.executable, PIVOT_SCRIPT,
            "--register", reg_path, "--register-sheet", cfg["register_sheet"],
            "--hf-register", hf_path, "--hf-sheet", first_sheet_name(hf_path),
            "--taluka-col", cfg["taluka_col"], "--facility-col", cfg["facility_col"],
            "--type-col", cfg["type_col"], "--public-value", cfg["public_value"],
            "--status-col", cfg["status_col"],
            "--status-order", ",".join(cfg["status_order"]),
            "--nonfunctional-taluka", NONFUNCTIONAL_TALUKA,
            "--reassign-to", REASSIGN_TO, "--reassign-sector", REASSIGN_SECTOR,
            "--title", cfg["title"], "--sheet-name", cfg["sheet_name"],
            "--linelist", cfg["linelist"],
            "--output", out_xlsx,
        ]
        proc = subprocess.run(cmd, capture_output=True, text=True)
        if proc.returncode != 0:
            log(f"ERROR building Excel for {cfg['key']}:\n{proc.stderr}")
            continue
        log(f"  Excel saved: {out_xlsx}")

        counts, total = count_by_taluka(cfg, reg_path)
        results.append({"cfg": cfg, "counts": counts, "total": total})

    if not results:
        log("Nothing was built.")
        return 1

    ppt_path = os.path.join(day_folder, f"TB_Daily_Report_{datetime.now():%d-%m-%Y}.pptx")
    build_ppt(results, ppt_path)
    log(f"PPT saved: {ppt_path}")
    return 0


if __name__ == "__main__":
    code = 1
    try:
        code = main()
    except Exception:
        log("UNEXPECTED ERROR:\n" + traceback.format_exc())
    # write a log next to the script so you can check what happened
    try:
        with open(os.path.join(SCRIPT_DIR, "daily_run_log.txt"), "a", encoding="utf-8") as f:
            f.write("\n".join(LOG_LINES) + "\n" + "-" * 60 + "\n")
    except Exception:
        pass
    sys.exit(code)
