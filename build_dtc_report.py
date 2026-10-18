#!/usr/bin/env python3
"""
Differentiated Care (DTC) monthly analysis - Mahisagar District.
Built from DIFFERENTIAL_CARE_FINAL_SKILL.md (rules, windows, colours, layout).

Usage:
  python build_dtc_report.py --source "DTC CUMULATIVE.xlsx" \
      --months "Jan-26,Feb-26,Mar-26" --output DTC_Report.xlsx [--order chrono|recent]

Output sheets: one per month (Baseline..6th FU, Low/Med/High risk, Not Done, TX Not Initiated),
then "DTC Not Done" and "TX Not Initiated" patient lists.
All % cells are live Excel formulas (they calculate when the file is opened in Excel).
"""
import argparse
import calendar
from datetime import date, datetime, timedelta

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

TALUKAS = ['Santrampur', 'Lunawada', 'Balasinor', 'Virpur', 'Kadana', 'Khanpur']
TALUKA_COL = 'Spectrum_Current_TBU'
PHC_COL = 'HF REGISTER.Name of PHC(HWC)'

FU_KEYS = ['baseline', '1st', '2nd', '3rd', '4th', '5th', '6th']
FU_LABELS = ['Baseline', '1st Follow-Up', '2nd Follow-Up', '3rd Follow-Up',
             '4th Follow-Up', '5th Follow-Up', '6th Follow-Up']
FU_DTC = {'baseline': 'baseline', '1st': 'followUp1stMonth', '2nd': 'followUp2ndMonth',
          '3rd': 'followUp3rdMonth', '4th': 'followUp4thMonth', '5th': 'followUp5thMonth',
          '6th': 'followUp6thMonth'}

# August-2026 anchor windows (section 4 of the skill)
ANCHOR_START = date(2026, 8, 1)
AUG_DELTAS = [
    ('baseline', 'notification', date(2026, 8, 1), date(2026, 8, 30)),
    ('1st', 'treatment', date(2026, 6, 6), date(2026, 7, 4)),
    ('2nd', 'treatment', date(2026, 5, 9), date(2026, 6, 6)),
    ('3rd', 'treatment', date(2026, 4, 11), date(2026, 5, 9)),
    ('4th', 'treatment', date(2026, 3, 14), date(2026, 4, 11)),
    ('5th', 'treatment', date(2026, 2, 14), date(2026, 3, 14)),
    ('6th', 'treatment', date(2026, 1, 17), date(2026, 2, 14)),
]
GOOD_OUTCOMES = ['CURED', 'TREATMENT_COMPLETE']


def get_windows(m_start, m_end):
    delta = (m_start - ANCHOR_START).days
    wins = []
    for fk, dc, f, t in AUG_DELTAS:
        if fk == 'baseline':
            wins.append((fk, dc, m_start, m_end))
        else:
            wins.append((fk, dc, f + timedelta(days=delta), t + timedelta(days=delta)))
    return wins


def norm_taluka(x):
    s = str(x).strip().title()
    return {'Lunavada': 'Lunawada'}.get(s, s)


def parse_months(text):
    out = []
    for tok in text.split(','):
        tok = tok.strip()
        if not tok:
            continue
        d = datetime.strptime(tok, '%b-%y')
        last = calendar.monthrange(d.year, d.month)[1]
        out.append((d.strftime('%B %Y'), d.strftime('%b %Y'), date(d.year, d.month, 1),
                    date(d.year, d.month, last)))
    return out


# ----------------------------------------------------------------------
# DATA
# ----------------------------------------------------------------------
def load(src):
    df = pd.read_excel(src, sheet_name='NOTIFICATION VS DTC')
    for c in ['Spectrum_Diagnosis_Date', 'Treatment_Initiation_Date', 'Date_of_Treatment_Outcome']:
        df[c] = pd.to_datetime(df[c], errors='coerce')

    cols = set(df.columns)
    cfg = {
        'assess': 'DTC.Assessment Type' if 'DTC.Assessment Type' in cols else 'Assessment Type',
        'low': 'DTC.Low Risk' if 'DTC.Low Risk' in cols else 'Low Risk',
        'med': 'DTC.Medium Risk' if 'DTC.Medium Risk' in cols else 'Medium Risk',
        'high': 'DTC.High Risk' if 'DTC.High Risk' in cols else 'High Risk',
        'phi': 'Spectrum_Enrolment_PHI' if 'Spectrum_Enrolment_PHI' in cols else 'Spectrum_Current_HF',
    }
    for k in ('assess', 'low', 'med', 'high', 'phi'):
        if cfg[k] not in cols:
            raise SystemExit(f'Column for "{k}" not found in NOTIFICATION VS DTC sheet (looked for {cfg[k]}).')

    hf = pd.read_excel(src, sheet_name='HF REGISTER')
    phi_phc = dict(zip(hf['Health Facility'].dropna(), hf['Name of PHC(HWC)']))
    phc_src = df[PHC_COL] if PHC_COL in cols else pd.Series([None] * len(df), index=df.index)
    df['_phc_name'] = phc_src.fillna(df[cfg['phi']].map(phi_phc))
    df['_phi_name'] = df[cfg['phi']]
    df['_tal'] = df[TALUKA_COL].map(norm_taluka)
    df[cfg['assess']] = df[cfg['assess']].astype('string').str.strip()
    uniq = df.drop_duplicates(subset='Episode_ID', keep='first').copy()
    return df, uniq, cfg


def eligible(dfu, dcol, f, t):
    f_ts = pd.Timestamp(f)
    t_ts = pd.Timestamp(t) + pd.Timedelta(days=1)   # inclusive of the last day
    if dcol == 'notification':
        d = dfu['Spectrum_Diagnosis_Date']
        return dfu[d.notna() & (d >= f_ts) & (d < t_ts)]
    tx = dfu['Treatment_Initiation_Date']
    out = dfu['Treatment_Outcome']
    od = dfu['Date_of_Treatment_Outcome']
    ok = out.isna() | out.isin(GOOD_OUTCOMES) | (od.notna() & (od >= f_ts))
    return dfu[tx.notna() & (tx >= f_ts) & (tx < t_ts) & ok]


def fmt(d):
    return '' if pd.isna(d) else pd.Timestamp(d).strftime('%d-%b-%Y')


def compute(df, uniq, cfg, months):
    """Return per-month results plus the two patient lists."""
    by_t_all = {t: df[df['_tal'] == t] for t in TALUKAS}
    by_t_uniq = {t: uniq[uniq['_tal'] == t] for t in TALUKAS}
    results, nd_rows, tx_rows = {}, [], []

    for month_name, short, m_start, m_end in months:
        res = {}
        for fk, dcol, f, t in get_windows(m_start, m_end):
            dtc_type = FU_DTC[fk]
            for tal in TALUKAS:
                u = by_t_uniq[tal]
                el = eligible(u, dcol, f, t)
                el_ids = set(el['Episode_ID'])
                a = by_t_all[tal]
                done = a[(a[cfg['assess']] == dtc_type) & a['Episode_ID'].isin(el_ids)] \
                    .drop_duplicates('Episode_ID')
                hh = done[cfg['high']].notna()
                hm = done[cfg['med']].notna()
                hl = done[cfg['low']].notna()
                row = {'elig': len(el_ids), 'done': len(done), 'low': int((hl & ~hm & ~hh).sum()),
                       'med': int((hm & ~hh).sum()), 'high': int(hh.sum()), 'tx': None}

                nd_ids = el_ids - set(done['Episode_ID'])
                for _, pr in u[u['Episode_ID'].isin(nd_ids)].iterrows():
                    nd_rows.append([month_name, tal, pr['_phc_name'], pr['_phi_name'],
                                    FU_LABELS[FU_KEYS.index(fk)], pr['Episode_ID'],
                                    pr.get('Patient_Name'), fmt(pr['Spectrum_Diagnosis_Date']),
                                    fmt(pr['Treatment_Initiation_Date']), pr.get('Treatment_Outcome'),
                                    f.strftime('%d-%b-%Y'), t.strftime('%d-%b-%Y')])

                if fk == 'baseline':
                    d = u['Spectrum_Diagnosis_Date']
                    m = d.notna() & (d >= pd.Timestamp(f)) & (d < pd.Timestamp(t) + pd.Timedelta(days=1)) \
                        & u['Treatment_Initiation_Date'].isna()
                    txp = u[m]
                    row['tx'] = len(txp)
                    for _, pr in txp.iterrows():
                        tx_rows.append([month_name, tal, pr['_phc_name'], pr['_phi_name'],
                                        pr['Episode_ID'], pr.get('Patient_Name'),
                                        fmt(pr['Spectrum_Diagnosis_Date']), pr.get('Treatment_Outcome')])
                res[(tal, fk)] = row
        results[month_name] = res
    return results, nd_rows, tx_rows


# ----------------------------------------------------------------------
# EXCEL
# ----------------------------------------------------------------------
def fill(hexcol):
    return PatternFill('solid', start_color=hexcol, end_color=hexcol)


THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADERS = ['Follow-Up Period', 'Eligible for DTC', 'DTC Done (No.)', 'DTC Done (%)',
           'Low Risk No.', 'Low % of DTC', 'Low % of Eligible',
           'Medium Risk No.', 'Med % of DTC', 'Med % of Eligible',
           'High Risk No.', 'High % of DTC', 'High % of Eligible',
           'DTC NOT Done (No.)', 'DTC NOT Done (%)',
           'TX NOT Initiated (No.)', 'TX NOT Initiated (% Elig)']
WIDTHS = [22, 11, 10, 9, 8, 10, 11, 8, 10, 11, 8, 10, 11, 11, 10, 12, 11]
# column -> (fill, font colour)
COLSTYLE = {2: ('F2F2F2', '000000'), 3: ('D6E4F7', '000000'), 4: ('D6E4F7', '000000'),
            5: ('E2EFDA', '375623'), 6: ('E2EFDA', '375623'), 7: ('E2EFDA', '375623'),
            8: ('FFF2CC', '7F6000'), 9: ('FFF2CC', '7F6000'), 10: ('FFF2CC', '7F6000'),
            11: ('FCE4D6', '9C0006'), 12: ('FCE4D6', '9C0006'), 13: ('FCE4D6', '9C0006'),
            14: ('F5CBCC', '9C0006'), 15: ('F5CBCC', '9C0006'),
            16: ('FDE9D1', 'C65911'), 17: ('FDE9D1', 'C65911')}
COUNT_COLS = {'B': 'elig', 'C': 'done', 'E': 'low', 'H': 'med', 'K': 'high'}


def pct_formulas(r):
    return {4: f'=IFERROR(C{r}/B{r},0)', 6: f'=IFERROR(E{r}/C{r},0)', 7: f'=IFERROR(E{r}/B{r},0)',
            9: f'=IFERROR(H{r}/C{r},0)', 10: f'=IFERROR(H{r}/B{r},0)',
            12: f'=IFERROR(K{r}/C{r},0)', 13: f'=IFERROR(K{r}/B{r},0)',
            14: f'=B{r}-C{r}', 15: f'=IFERROR(N{r}/B{r},0)', 17: f'=IFERROR(P{r}/B{r},0)'}


def style_row(ws, r, base_fill=None, base_font='FFFFFF', bold=True):
    for c in range(1, 18):
        cell = ws.cell(r, c)
        cell.border = BORDER
        if base_fill:
            cell.fill = fill(base_fill)
            cell.font = Font(bold=bold, color=base_font, size=10)
        if c in (4, 6, 7, 9, 10, 12, 13, 15, 17):
            cell.number_format = '0.0%'
        if c > 1:
            cell.alignment = Alignment(horizontal='center', vertical='center')


def write_month_sheet(wb, sheet_name, month_name, res):
    ws = wb.create_sheet(sheet_name)
    for i, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[L(i)].width = w

    ws.merge_cells('A1:Q1')
    ws['A1'] = f'DIFFERENTIATED CARE (DTC) - MAHISAGAR DISTRICT - {month_name.upper()}'
    ws['A1'].font = Font(bold=True, size=14, color='FFFFFF')
    ws['A1'].fill = fill('1F4E79')
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 26
    ws.merge_cells('A2:Q2')
    ws['A2'] = 'Eligible = Baseline: notified in month | 1st-6th FU: treatment started in the matching window'
    ws['A2'].font = Font(italic=True, size=10, color='FFFFFF')
    ws['A2'].fill = fill('2E75B6')
    ws['A2'].alignment = Alignment(horizontal='center')

    groups = [(1, 1, 'Follow-Up'), (2, 2, 'Eligible'), (3, 4, 'DTC Done'), (5, 7, 'Low Risk'),
              (8, 10, 'Medium Risk'), (11, 13, 'High Risk'), (14, 15, 'DTC Not Done'),
              (16, 17, 'TX Not Initiated')]
    for c1, c2, label in groups:
        if c1 != c2:
            ws.merge_cells(start_row=3, start_column=c1, end_row=3, end_column=c2)
        cell = ws.cell(3, c1, label)
        for c in range(c1, c2 + 1):
            ws.cell(3, c).fill = fill('4472C4')
            ws.cell(3, c).border = BORDER
        cell.font = Font(bold=True, color='FFFFFF', size=10)
        cell.alignment = Alignment(horizontal='center')
    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(4, c, h)
        cell.font = Font(bold=True, color='FFFFFF', size=9)
        cell.fill = fill('1F4E79')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[4].height = 38
    ws.freeze_panes = 'B5'

    r = 5
    taluka_rows = {}      # (taluka, fk) -> excel row
    subtotal_rows = []
    for tal in TALUKAS:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=17)
        ws.cell(r, 1, tal.upper())
        style_row(ws, r, '4472C4')
        ws.cell(r, 1).alignment = Alignment(horizontal='left')
        r += 1
        first = r
        for i, fk in enumerate(FU_KEYS):
            d = res[(tal, fk)]
            ws.cell(r, 1, FU_LABELS[i])
            ws.cell(r, 2, d['elig'])
            ws.cell(r, 3, d['done'])
            ws.cell(r, 5, d['low'])
            ws.cell(r, 8, d['med'])
            ws.cell(r, 11, d['high'])
            if d['tx'] is not None:
                ws.cell(r, 16, d['tx'])
            for c, f in pct_formulas(r).items():
                if c == 17 and fk != 'baseline':
                    continue          # TX-not-initiated cells stay blank on follow-up rows
                ws.cell(r, c, f)
            style_row(ws, r)
            for c, (bg, fg) in COLSTYLE.items():
                ws.cell(r, c).fill = fill(bg)
                ws.cell(r, c).font = Font(color=fg, size=10)
            if fk == 'baseline':
                ws.cell(r, 1).fill = fill('1F4E79')
                ws.cell(r, 1).font = Font(bold=True, color='FFFFFF', size=10)
                ws.cell(r, 16).fill = fill('F4B942')
                ws.cell(r, 16).font = Font(bold=True, color='FFFFFF', size=10)
            taluka_rows[(tal, fk)] = r
            r += 1
        last = r - 1
        ws.cell(r, 1, f'SUBTOTAL - {tal}')
        for col in ('B', 'C', 'E', 'H', 'K', 'P'):
            ws[f'{col}{r}'] = f'=SUM({col}{first}:{col}{last})'
        for c, f in pct_formulas(r).items():
            ws.cell(r, c, f)
        style_row(ws, r, '4472C4')
        subtotal_rows.append(r)
        r += 2

    # District block
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=17)
    ws.cell(r, 1, 'MAHISAGAR DISTRICT - TOTAL')
    style_row(ws, r, '0D3268')
    ws.cell(r, 1).alignment = Alignment(horizontal='left')
    r += 1
    d_first = r
    for i, fk in enumerate(FU_KEYS):
        ws.cell(r, 1, FU_LABELS[i])
        for col in ('B', 'C', 'E', 'H', 'K', 'P'):
            refs = '+'.join(f'{col}{taluka_rows[(t, fk)]}' for t in TALUKAS)
            if col == 'P' and fk != 'baseline':
                continue
            ws[f'{col}{r}'] = f'={refs}'
        for c, f in pct_formulas(r).items():
            if c == 17 and fk != 'baseline':
                continue
            ws.cell(r, c, f)
        style_row(ws, r)
        for c, (bg, fg) in COLSTYLE.items():
            ws.cell(r, c).fill = fill(bg)
            ws.cell(r, c).font = Font(color=fg, size=10, bold=True)
        ws.cell(r, 1).fill = fill('0D3268')
        ws.cell(r, 1).font = Font(bold=True, color='FFFFFF', size=10)
        r += 1
    d_last = r - 1
    ws.cell(r, 1, 'MAHISAGAR GRAND TOTAL')
    for col in ('B', 'C', 'E', 'H', 'K', 'P'):
        ws[f'{col}{r}'] = f'=SUM({col}{d_first}:{col}{d_last})'
    for c, f in pct_formulas(r).items():
        ws.cell(r, c, f)
    style_row(ws, r, '0D3268')
    ws.cell(r, 14).fill = fill('C00000')
    ws.cell(r, 15).fill = fill('C00000')
    ws.cell(r, 16).fill = fill('C65911')
    ws.cell(r, 17).fill = fill('C65911')


def write_list_sheet(wb, name, headers, rows, widths):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(1, c)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = fill('1F4E79')
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        ws.column_dimensions[L(c)].width = widths[c - 1]
    rows = sorted(rows, key=lambda x: (str(x[1]), str(x[2]), str(x[0])))
    for row in rows:
        ws.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in row])
    ws.auto_filter.ref = f'A1:{L(len(headers))}{max(len(rows) + 1, 2)}'
    ws.freeze_panes = 'A2'


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--source', required=True)
    ap.add_argument('--months', required=True, help='e.g. "Jan-26,Feb-26,Mar-26"')
    ap.add_argument('--output', required=True)
    ap.add_argument('--order', default='chrono', choices=['chrono', 'recent'])
    a = ap.parse_args()

    months = parse_months(a.months)
    if a.order == 'recent':
        months = months[::-1]
    df, uniq, cfg = load(a.source)
    print(f'Loaded {len(df)} rows, {len(uniq)} unique patients')
    results, nd_rows, tx_rows = compute(df, uniq, cfg, months)

    wb = Workbook()
    wb.remove(wb.active)
    for month_name, short, _, _ in months:
        write_month_sheet(wb, short, month_name, results[month_name])
    write_list_sheet(wb, 'DTC Not Done',
                     ['Month', 'Taluka', 'PHC Name', 'PHI Name', 'Follow-Up', 'Episode ID',
                      'Patient Name', 'Diagnosis Date', 'Treat. Init Date', 'Treatment Outcome',
                      'Window From', 'Window To'], nd_rows,
                     [14, 14, 26, 28, 16, 12, 24, 14, 14, 22, 13, 13])
    write_list_sheet(wb, 'TX Not Initiated',
                     ['Month', 'Taluka', 'PHC Name', 'PHI Name', 'Episode ID', 'Patient Name',
                      'Diagnosis Date', 'Treatment Outcome'], tx_rows,
                     [14, 14, 26, 28, 12, 24, 14, 22])
    wb.save(a.output)
    print(f'Saved {a.output}  (DTC not done rows: {len(nd_rows)}, TX not initiated rows: {len(tx_rows)})')


if __name__ == '__main__':
    main()
