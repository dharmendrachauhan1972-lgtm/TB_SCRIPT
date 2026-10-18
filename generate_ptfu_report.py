#!/usr/bin/env python3
"""
Generate a Post-Treatment Follow-up (PTFU) due-vs-performed Excel analysis
for one or more report months, from a "PTFU_CUMULATIVE"-style export that
contains a "NOTIFICATION VS PTFU" sheet (Nikshay/Spectrum RNTCP-style data).

Usage:
    python generate_ptfu_report.py \
        --source /path/to/PTFU_CUMULATVE.xlsx \
        --months Jun-26,Jul-26,Aug-26,Sep-26 \
        --output /path/to/PTFU_Analysis.xlsx

Report month format: "Mon-YY" (e.g. Jun-26 = June 2026).

For each report month, four follow-up cohorts are evaluated:
    6 month  -> Treatment Outcome date = report month - 6 months
    12 month -> Treatment Outcome date = report month - 12 months
    18 month -> Treatment Outcome date = report month - 18 months
    24 month -> Treatment Outcome date = report month - 24 months

Denominator: patients with Treatment_Outcome in {CURED, TREATMENT_COMPLETE}
whose Date_of_Treatment_Outcome falls in the cohort month.

Numerator: denominator patients who have a PTFU record for that exact
follow-up type (PTFU.MonthOfCohort matches 6/12/18/24) with a non-blank
PTFU.Response.

Output workbook (3 sheets):
    1. PTFU Analysis        - due/performed/%/pending table, one block per report month, with live formulas
    2. PTFU Not Done List   - patient-wise list of pending follow-ups (Taluka / PHI / PHC / Episode ID / Patient Name)
    3. PTFU_Data            - deduplicated per-patient raw data + Done flags that the Analysis formulas reference
"""

import argparse
import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

SOURCE_SHEET = 'NOTIFICATION VS PTFU'
TYPE_OFFSETS = [6, 12, 18, 24]
OUTCOME_OK = {'CURED', 'TREATMENT_COMPLETE'}

REQUIRED_COLS = [
    'Episode_ID', 'Patient_Name', 'Spectrum_Current_TBU', 'Spectrum_Current_HF',
    'HF REGISTER.Name of PHC(HWC)', 'Treatment_Outcome', 'Date_of_Treatment_Outcome',
    'PTFU.MonthOfCohort', 'PTFU.Response',
]

FONT = 'Arial'


def parse_month_token(token):
    """'Jun-26' -> (2026, 6)"""
    token = token.strip()
    dt = pd.to_datetime('01-' + token, format='%d-%b-%y', errors='coerce')
    if pd.isna(dt):
        dt = pd.to_datetime(token, errors='coerce')
    if pd.isna(dt):
        raise ValueError(f'Could not parse report month "{token}". Use format like Jun-26.')
    return dt.year, dt.month


def cohort_month(ry, rm, offset):
    total_m = (ry * 12 + (rm - 1)) - offset
    cy, cm = divmod(total_m, 12)
    return cy, cm + 1


def build_report(source_path, report_month_tokens, output_path, sheet_name=SOURCE_SHEET):
    report_months = []
    for tok in report_month_tokens:
        ry, rm = parse_month_token(tok)
        report_months.append((tok.strip(), ry, rm))

    df = pd.read_excel(source_path, sheet_name=sheet_name)
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(
            f'Source sheet "{sheet_name}" is missing expected columns: {missing}. '
            'This skill expects a "NOTIFICATION VS PTFU"-style export (Notification data '
            'left-joined with PTFU visit records and the HF register PHC mapping).'
        )

    df['Date_of_Treatment_Outcome'] = pd.to_datetime(df['Date_of_Treatment_Outcome'], errors='coerce')

    base_cols = ['Episode_ID', 'Patient_Name', 'Spectrum_Current_TBU', 'Spectrum_Current_HF',
                 'HF REGISTER.Name of PHC(HWC)', 'Treatment_Outcome', 'Date_of_Treatment_Outcome']
    base = df[base_cols].drop_duplicates(subset='Episode_ID').copy()

    for t in TYPE_OFFSETS:
        sub = df[(df['PTFU.MonthOfCohort'] == t) & (df['PTFU.Response'].notna())]
        done_eps = set(sub['Episode_ID'].unique())
        base[f'Done_{t}'] = base['Episode_ID'].isin(done_eps)

    base = base.sort_values(
        ['Spectrum_Current_TBU', 'Spectrum_Current_HF', 'Patient_Name']
    ).reset_index(drop=True)
    outcome_ok_mask = base['Treatment_Outcome'].isin(OUTCOME_OK)

    hdr_font = Font(name=FONT, bold=True, color='FFFFFF', size=11)
    hdr_fill = PatternFill('solid', fgColor='1F4E78')
    title_font = Font(name=FONT, bold=True, size=14, color='1F4E78')
    sub_font = Font(name=FONT, bold=True, size=11, color='444444')
    normal_font = Font(name=FONT, size=10)
    bold_font = Font(name=FONT, bold=True, size=10)
    thin = Side(style='thin', color='B7B7B7')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    total_fill = PatternFill('solid', fgColor='D9E1F2')
    note_font = Font(name=FONT, size=9, italic=True, color='777777')
    pend_hdr_fill = PatternFill('solid', fgColor='C00000')
    pend_title_font = Font(name=FONT, bold=True, size=14, color='C00000')

    wb = Workbook()

    # ---------------- PTFU_Data ----------------
    ws_data = wb.active
    ws_data.title = 'PTFU_Data'
    data_cols = ['Episode_ID', 'Patient_Name', 'Taluka (Current)', 'PHI Name (Current HF)',
                 'PHC Name', 'Treatment_Outcome', 'Date_of_Treatment_Outcome',
                 'Done_6Month', 'Done_12Month', 'Done_18Month', 'Done_24Month']
    ws_data.append(data_cols)
    for c in range(1, len(data_cols) + 1):
        cell = ws_data.cell(row=1, column=c)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for _, r in base.iterrows():
        ws_data.append([
            int(r['Episode_ID']), r['Patient_Name'], r['Spectrum_Current_TBU'], r['Spectrum_Current_HF'],
            r['HF REGISTER.Name of PHC(HWC)'] if pd.notna(r['HF REGISTER.Name of PHC(HWC)']) else '',
            r['Treatment_Outcome'] if pd.notna(r['Treatment_Outcome']) else '',
            r['Date_of_Treatment_Outcome'].to_pydatetime() if pd.notna(r['Date_of_Treatment_Outcome']) else None,
            bool(r['Done_6']), bool(r['Done_12']), bool(r['Done_18']), bool(r['Done_24']),
        ])

    last_data_row = len(base) + 1
    for row in ws_data.iter_rows(min_row=2, max_row=last_data_row, min_col=1, max_col=len(data_cols)):
        for cell in row:
            cell.font = normal_font
            cell.border = border
            if cell.column == 7 and cell.value is not None:
                cell.number_format = 'dd-mmm-yyyy'
            if cell.column in (8, 9, 10, 11):
                cell.alignment = Alignment(horizontal='center')

    for i, w in enumerate([12, 26, 18, 26, 20, 20, 18, 12, 12, 12, 12], start=1):
        ws_data.column_dimensions[get_column_letter(i)].width = w
    ws_data.freeze_panes = 'A2'
    tbl = Table(displayName='PTFU_Data_Table', ref=f'A1:K{last_data_row}')
    tbl.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
    ws_data.add_table(tbl)

    # ---------------- PTFU Analysis ----------------
    ws = wb.create_sheet('PTFU Analysis', 0)
    ws['A1'] = 'Post-Treatment Follow-up (PTFU) Analysis'
    ws['A1'].font = title_font
    ws.merge_cells('A1:G1')
    month_label_str = ', '.join(m[0] for m in report_months)
    ws['A2'] = f'District PTFU due vs. performed — Report months: {month_label_str}'
    ws['A2'].font = sub_font
    ws.merge_cells('A2:G2')

    headers = ['Report Month', 'Type of Follow-up Due', 'Treatment Complete Month (Cohort)',
               'Denominator (Due)', 'PTFU Completed (Numerator)', '% Achievement',
               'Not Entered / Blank (Pending)']

    DONE_COL = {6: 'H', 12: 'I', 18: 'J', 24: 'K'}
    COL_DATE, COL_OUT = 'G', 'F'

    cur_row = 4
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=cur_row, column=i, value=h)
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = border
    cur_row += 1

    for rm_label, ry, rm in report_months:
        first_r = cur_row
        for t in TYPE_OFFSETS:
            cy, cm = cohort_month(ry, rm, t)
            ws.cell(row=cur_row, column=1, value=rm_label)
            ws.cell(row=cur_row, column=2, value=f'{t} Month Follow-up')
            ws.cell(row=cur_row, column=3, value=f'=TEXT(DATE({cy},{cm},1),"mmm-yy")')

            start_expr = f'DATE({cy},{cm},1)'
            end_expr = f'EOMONTH(DATE({cy},{cm},1),0)'

            denom_formula = (
                f'=SUMPRODUCT((PTFU_Data!${COL_DATE}$2:${COL_DATE}${last_data_row}>={start_expr})'
                f'*(PTFU_Data!${COL_DATE}$2:${COL_DATE}${last_data_row}<={end_expr})'
                f'*((PTFU_Data!${COL_OUT}$2:${COL_OUT}${last_data_row}="CURED")'
                f'+(PTFU_Data!${COL_OUT}$2:${COL_OUT}${last_data_row}="TREATMENT_COMPLETE")))'
            )
            ws.cell(row=cur_row, column=4, value=denom_formula)

            dcol = DONE_COL[t]
            num_formula = (
                f'=SUMPRODUCT((PTFU_Data!${COL_DATE}$2:${COL_DATE}${last_data_row}>={start_expr})'
                f'*(PTFU_Data!${COL_DATE}$2:${COL_DATE}${last_data_row}<={end_expr})'
                f'*((PTFU_Data!${COL_OUT}$2:${COL_OUT}${last_data_row}="CURED")'
                f'+(PTFU_Data!${COL_OUT}$2:${COL_OUT}${last_data_row}="TREATMENT_COMPLETE"))'
                f'*(PTFU_Data!${dcol}$2:${dcol}${last_data_row}=TRUE))'
            )
            ws.cell(row=cur_row, column=5, value=num_formula)
            ws.cell(row=cur_row, column=6, value=f'=IFERROR(E{cur_row}/D{cur_row},0)')
            ws.cell(row=cur_row, column=6).number_format = '0.0%'
            ws.cell(row=cur_row, column=7, value=f'=D{cur_row}-E{cur_row}')

            for col in range(1, 8):
                c = ws.cell(row=cur_row, column=col)
                c.font = normal_font
                c.border = border
                c.alignment = Alignment(horizontal='center')
            cur_row += 1

        subtotal_row = cur_row
        last_r = cur_row - 1
        ws.cell(row=subtotal_row, column=1, value=f'{rm_label} Total')
        ws.cell(row=subtotal_row, column=2, value='All Follow-up Types')
        ws.cell(row=subtotal_row, column=4, value=f'=SUM(D{first_r}:D{last_r})')
        ws.cell(row=subtotal_row, column=5, value=f'=SUM(E{first_r}:E{last_r})')
        ws.cell(row=subtotal_row, column=6, value=f'=IFERROR(E{subtotal_row}/D{subtotal_row},0)')
        ws.cell(row=subtotal_row, column=6).number_format = '0.0%'
        ws.cell(row=subtotal_row, column=7, value=f'=SUM(G{first_r}:G{last_r})')
        for col in range(1, 8):
            c = ws.cell(row=subtotal_row, column=col)
            c.font = bold_font
            c.fill = total_fill
            c.border = border
            c.alignment = Alignment(horizontal='center')
        cur_row += 2

    ws.cell(row=cur_row, column=1, value='Notes:')
    ws.cell(row=cur_row, column=1).font = bold_font
    cur_row += 1
    notes = [
        'Denominator: patients with Treatment Outcome = Cured or Treatment Completed, whose treatment-outcome '
        '(completion) date falls in the cohort month shown, drawn from the "NOTIFICATION VS PTFU" export.',
        'Numerator (PTFU Completed): denominator patients for whom a Post-Treatment Follow-up form for that '
        'specific follow-up type (6/12/18/24 month) was submitted with a response recorded.',
        'Pending: denominator patients with no PTFU response recorded for that follow-up type (form not entered '
        'or left blank) — see the "PTFU Not Done List" sheet for the patient-wise list.',
        'Raw linked data used for all formulas is on the "PTFU_Data" sheet.',
    ]
    for n in notes:
        ws.cell(row=cur_row, column=1, value='- ' + n)
        ws.cell(row=cur_row, column=1).font = note_font
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=7)
        cur_row += 1

    for col, w in zip('ABCDEFG', [14, 22, 26, 16, 22, 14, 22]):
        ws.column_dimensions[col].width = w
    ws.sheet_view.showGridLines = False

    # ---------------- PTFU Not Done List ----------------
    rows_out = []
    for rm_label, ry, rm in report_months:
        for t in TYPE_OFFSETS:
            cy, cm = cohort_month(ry, rm, t)
            cohort_start = pd.Timestamp(cy, cm, 1)
            cohort_end = cohort_start + pd.offsets.MonthEnd(0)
            mask = (
                outcome_ok_mask
                & (base['Date_of_Treatment_Outcome'] >= cohort_start)
                & (base['Date_of_Treatment_Outcome'] <= cohort_end)
            )
            denom_df = base[mask]
            pending_df = denom_df[~denom_df[f'Done_{t}']]
            for _, r in pending_df.iterrows():
                rows_out.append({
                    'Report Month': rm_label,
                    'Follow-up Type Due': f'{t} Month',
                    'Treatment Complete Month': cohort_start.strftime('%b-%y'),
                    'Taluka (Current)': r['Spectrum_Current_TBU'],
                    'PHI Name (Current HF)': r['Spectrum_Current_HF'],
                    'PHC Name': r['HF REGISTER.Name of PHC(HWC)'] if pd.notna(r['HF REGISTER.Name of PHC(HWC)']) else '',
                    'Episode ID': int(r['Episode_ID']),
                    'Patient Name': r['Patient_Name'],
                    'Treatment Outcome': r['Treatment_Outcome'],
                    'Treatment Outcome Date': r['Date_of_Treatment_Outcome'],
                })

    pending_all = pd.DataFrame(rows_out)
    if not pending_all.empty:
        pending_all = pending_all.sort_values(
            ['Report Month', 'Follow-up Type Due', 'Taluka (Current)', 'PHI Name (Current HF)', 'Patient Name']
        ).reset_index(drop=True)

    ws2 = wb.create_sheet('PTFU Not Done List')
    ws2['A1'] = 'Post-Treatment Follow-up — Patients Pending / Not Entered'
    ws2['A1'].font = pend_title_font
    ws2.merge_cells('A1:J1')
    ws2['A2'] = f'Report months: {month_label_str} — patient-wise, by Taluka / PHI / PHC'
    ws2['A2'].font = sub_font
    ws2.merge_cells('A2:J2')

    header_row = 4
    cols = ['Report Month', 'Follow-up Type Due', 'Treatment Complete Month', 'Taluka (Current)',
            'PHI Name (Current HF)', 'PHC Name', 'Episode ID', 'Patient Name',
            'Treatment Outcome', 'Treatment Outcome Date']
    for i, h in enumerate(cols, start=1):
        c = ws2.cell(row=header_row, column=i, value=h)
        c.font = hdr_font
        c.fill = pend_hdr_fill
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = border

    r = header_row + 1
    for _, row in pending_all.iterrows():
        for i, h in enumerate(cols, start=1):
            val = row[h]
            if h == 'Treatment Outcome Date' and pd.notna(val):
                val = val.to_pydatetime()
            c = ws2.cell(row=r, column=i, value=val)
            c.font = normal_font
            c.border = border
            if h == 'Treatment Outcome Date' and val:
                c.number_format = 'dd-mmm-yyyy'
        r += 1

    last_row = r - 1
    for i, w in enumerate([13, 16, 22, 18, 26, 20, 12, 26, 20, 20], start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = f'A{header_row + 1}'
    if last_row >= header_row + 1:
        tbl2 = Table(displayName='PTFU_Pending_Table', ref=f'A{header_row}:J{last_row}')
        tbl2.tableStyleInfo = TableStyleInfo(name='TableStyleMedium3', showRowStripes=True)
        ws2.add_table(tbl2)
    ws2.sheet_view.showGridLines = False

    wb.save(output_path)
    return {
        'output_path': output_path,
        'report_months': [m[0] for m in report_months],
        'pending_total': int(len(pending_all)),
        'pending_by_month_type': (
            pending_all.groupby(['Report Month', 'Follow-up Type Due']).size().to_dict()
            if not pending_all.empty else {}
        ),
    }


def main():
    ap = argparse.ArgumentParser(description='Generate PTFU due-vs-performed Excel analysis.')
    ap.add_argument('--source', required=True, help='Path to the PTFU cumulative export .xlsx')
    ap.add_argument('--months', required=True,
                     help='Comma-separated report months, e.g. "Jun-26,Jul-26,Aug-26,Sep-26"')
    ap.add_argument('--output', required=True, help='Path to write the output .xlsx')
    ap.add_argument('--sheet', default=SOURCE_SHEET,
                     help=f'Source sheet name (default: "{SOURCE_SHEET}")')
    args = ap.parse_args()

    months = [m.strip() for m in args.months.split(',') if m.strip()]
    result = build_report(args.source, months, args.output, sheet_name=args.sheet)

    print(f"Saved: {result['output_path']}")
    print(f"Report months: {', '.join(result['report_months'])}")
    print(f"Total pending PTFU records: {result['pending_total']}")
    for k, v in sorted(result['pending_by_month_type'].items()):
        print(f"  {k[0]} / {k[1]}: {v} pending")


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'ERROR: {e}', file=sys.stderr)
        sys.exit(1)
