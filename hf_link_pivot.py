#!/usr/bin/env python3
"""
hf_link_pivot.py — Build a Taluka-wise, Public/Private, PHC/Facility-wise status
pivot from any TB-program patient register, linked to an HF (Health Facility)
Register, with live Excel formulas (COUNTIFS + row-percentage) and optional
static line-list sheets.

This is the generalized engine behind the "Beneficiary Bank Detail Status"
and "PMTBMBA ABHA Seeding Status" Taluka-wise workbooks built for the
Mahisagar district TB program. See SKILL.md for the exact invocations used
for those two registers, and for the business rules (e.g. a taluka that is
non-functional for the private sector gets its private-sector records
reassigned to another taluka).

Usage (example — Beneficiary Register, bank detail status):

    python3 hf_link_pivot.py \\
      --register BENIFICIARY_REGISTER.xlsx --register-sheet "BENIFICIARY REGISTER" \\
      --hf-register HF_REGISTER.xlsx --hf-sheet Sheet1 \\
      --taluka-col CurrentFacilityTBU --facility-col CurrentFacilityPHI \\
      --type-col TypeOfPatient --public-value IndiaTbPublic \\
      --status-col Beneficiary_Status \\
      --status-order EMPTY,ENTERED,SENT,NOT_VALIDATED,VALIDATED,REJECTED,USER_REJECTED \\
      --nonfunctional-taluka KADANA --reassign-to LUNAWADA --reassign-sector Private \\
      --title "TALUKA-WISE BENEFICIARY BANK DETAIL STATUS — PUBLIC & PRIVATE HEALTH FACILITIES" \\
      --sheet-name "TALUKA WISE BANK STATUS" \\
      --linelist "ForegoBenefits=Yes:FOREGO BENEFITS - YES LIST:EpisodeId,PatientName,Age,Gender,PrimaryPhone,Stage,TypeOfCase,Beneficiary_Status,NoOfBenefitEligible,TotalBenefitAmountEligible,NoOfBenefitPaid_Nikshay,TotalAmountPaid_Nikshay" \\
      --output Beneficiary_Bank_Status_Taluka_Wise.xlsx

After running, recalculate the formulas with the xlsx skill's recalc.py
(this script only writes formula text; it does not evaluate it):

    python3 /mnt/skills/public/xlsx/scripts/recalc.py <output.xlsx> 150
"""
import argparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict

BOLD = Font(bold=True, name='Arial')
TITLE_FONT = Font(bold=True, size=14, name='Arial')
TALUKA_FONT = Font(bold=True, size=11, color='FFFFFF', name='Arial')
TALUKA_FILL = PatternFill('solid', fgColor='1F4E78')
SUBTOTAL_FONT = Font(bold=True, size=10, name='Arial')
SUBTOTAL_FILL = PatternFill('solid', fgColor='D9E1F2')
HEADER_FONT = Font(bold=True, size=10, color='FFFFFF', name='Arial')
HEADER_FILL = PatternFill('solid', fgColor='305496')
PCT_HEADER_FILL = PatternFill('solid', fgColor='548235')
NORMAL_FONT = Font(size=10, name='Arial')
PCT_FONT = Font(size=10, italic=True, color='595959', name='Arial')
GRAND_FONT = Font(bold=True, size=11, color='FFFFFF', name='Arial')
GRAND_FILL = PatternFill('solid', fgColor='7030A0')
LL_HEADER_FONT = Font(bold=True, size=10, color='FFFFFF', name='Arial')
LL_HEADER_FILL = PatternFill('solid', fgColor='305496')
thin = Side(style='thin', color='B7B7B7')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal='center', vertical='center')


def load_register(path, sheet):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    idx = {c: i for i, c in enumerate(header)}
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    return header, idx, rows


def load_hf_map(hf_path, hf_sheet):
    hf_wb = openpyxl.load_workbook(hf_path, read_only=True, data_only=True)
    hf_ws = hf_wb[hf_sheet]
    # Expected columns: District, TB Unit (Taluka), Health Facility, NIN, Name of PHC(HWC)
    hf_rows = [r for r in hf_ws.iter_rows(min_row=2, values_only=True)
               if isinstance(r[2], str)]
    hf_map = {}
    for r in hf_rows:
        key = (str(r[1]).strip().upper(), str(r[2]).strip().upper())
        hf_map[key] = str(r[4]).strip() if r[4] else 'UNMAPPED'
    return hf_rows, hf_map


def write_hf_map_sheet(wb, hf_rows):
    ws = wb.create_sheet('HF MAP')
    ws.append(['TB Unit (Taluka)', 'Health Facility', 'Name of PHC(HWC)', 'Key'])
    for j in range(1, 5):
        ws.cell(row=1, column=j).font = BOLD
    r_i = 2
    for r in hf_rows:
        taluka = str(r[1]).strip()
        facility = str(r[2]).strip()
        phc = str(r[4]).strip() if r[4] else 'UNMAPPED'
        ws.cell(row=r_i, column=1, value=taluka)
        ws.cell(row=r_i, column=2, value=facility)
        ws.cell(row=r_i, column=3, value=phc)
        ws.cell(row=r_i, column=4, value=f'=UPPER(TRIM(A{r_i}))&"|"&UPPER(TRIM(B{r_i}))')
        r_i += 1
    for col, w in zip('ABCD', [16, 22, 20, 30]):
        ws.column_dimensions[col].width = w
    return ws


def build_raw_sheet(wb, header, idx, rows, args):
    raw = wb.create_sheet('RAW DATA')
    n_cols = len(header)
    for j, h in enumerate(header, start=1):
        raw.cell(row=1, column=j, value=h)
    extra = ['Patient_Type', 'Group_Name', 'Effective_Taluka']
    for k, h in enumerate(extra):
        raw.cell(row=1, column=n_cols + 1 + k, value=h)

    type_col = n_cols + 1
    group_col = n_cols + 2
    eff_col = n_cols + 3
    tal_l = get_column_letter(idx[args.taluka_col] + 1)
    fac_l = get_column_letter(idx[args.facility_col] + 1)
    typ_l = get_column_letter(idx[args.type_col] + 1)
    type_col_l = get_column_letter(type_col)

    for i, r in enumerate(rows, start=2):
        for j, v in enumerate(r, start=1):
            raw.cell(row=i, column=j, value=v)
        raw.cell(row=i, column=type_col,
                  value=f'=IF(UPPER(TRIM({typ_l}{i}))=UPPER("{args.public_value}"),"Public","Private")')
        formula = (
            f'=IF({type_col_l}{i}="Public",'
            f'IFERROR(INDEX(\'HF MAP\'!$C:$C,MATCH(UPPER(TRIM({tal_l}{i}))&"|"&UPPER(TRIM({fac_l}{i})),\'HF MAP\'!$D:$D,0)),"UNMAPPED"),'
            f'TRIM({fac_l}{i}))'
        )
        raw.cell(row=i, column=group_col, value=formula)
        if args.nonfunctional_taluka and args.reassign_to:
            sector = args.reassign_sector or 'Private'
            eff_formula = (
                f'=IF(AND({type_col_l}{i}="{sector}",UPPER(TRIM({tal_l}{i}))=UPPER("{args.nonfunctional_taluka}")),'
                f'"{args.reassign_to.upper()}",{tal_l}{i})'
            )
        else:
            eff_formula = f'={tal_l}{i}'
        raw.cell(row=i, column=eff_col, value=eff_formula)

    for j in range(1, eff_col + 1):
        raw.cell(row=1, column=j).font = BOLD
        raw.column_dimensions[get_column_letter(j)].width = 16
    raw.freeze_panes = 'A2'
    return raw, type_col, group_col, eff_col, len(rows) + 1


def python_side_groups(idx, rows, args, hf_map):
    """Recompute Patient_Type / Group_Name / Effective_Taluka in Python
    (mirrors the Excel formulas) purely to know which rows/PHCs to render
    in the pivot — the actual counts in the workbook come from formulas."""
    def eff_taluka(tal, ptype):
        if (args.nonfunctional_taluka and args.reassign_to
                and ptype == (args.reassign_sector or 'Private')
                and tal.strip().upper() == args.nonfunctional_taluka.upper()):
            return args.reassign_to.upper()
        return tal.strip()

    def group_name(tal, fac, ptype):
        if ptype == 'Public':
            key = (tal.strip().upper(), fac.strip().upper())
            return hf_map.get(key, 'UNMAPPED')
        return fac.strip()

    struct = defaultdict(lambda: {'Public': set(), 'Private': set()})
    talukas = set()
    for r in rows:
        tal = str(r[idx[args.taluka_col]])
        fac = str(r[idx[args.facility_col]])
        ptype = 'Public' if str(r[idx[args.type_col]]).strip().upper() == args.public_value.upper() else 'Private'
        et = eff_taluka(tal, ptype)
        gn = group_name(tal, fac, ptype)
        talukas.add(et)
        struct[et][ptype].add(gn)
    return struct, sorted(talukas), eff_taluka, group_name


def build_pivot_sheet(wb, sheet_name, title_text, status_order, raw_last_row,
                       tal_col_l, type_col_l, group_col_l, status_col_l,
                       struct, talukas_order, note_lines=None):
    ws = wb.create_sheet(sheet_name, 0)
    RAW_TAL = f"'RAW DATA'!${tal_col_l}$2:${tal_col_l}${raw_last_row}"
    RAW_TYPE = f"'RAW DATA'!${type_col_l}$2:${type_col_l}${raw_last_row}"
    RAW_GROUP = f"'RAW DATA'!${group_col_l}$2:${group_col_l}${raw_last_row}"
    RAW_STATUS = f"'RAW DATA'!${status_col_l}$2:${status_col_l}${raw_last_row}"

    n_status = len(status_order)
    ncols = 2 + n_status * 2 + 1
    last_col_letter = get_column_letter(ncols)

    ws.merge_cells('A1:' + last_col_letter + '1')
    ws['A1'] = title_text
    ws['A1'].font = TITLE_FONT
    ws['A1'].alignment = CENTER

    HDR_ROW = 3
    ws.merge_cells(f'A{HDR_ROW}:B{HDR_ROW}')
    ws.cell(row=HDR_ROW, column=1, value='Taluka / Type / PHC-Facility').font = HEADER_FONT
    ws.cell(row=HDR_ROW, column=1).fill = HEADER_FILL
    ws.cell(row=HDR_ROW, column=1).alignment = CENTER

    col_ptr = 3
    count_col, pct_col = {}, {}
    for status in status_order:
        c = ws.cell(row=HDR_ROW, column=col_ptr, value=status)
        c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = CENTER
        count_col[status] = col_ptr
        c2 = ws.cell(row=HDR_ROW, column=col_ptr + 1, value=status + ' %')
        c2.font = HEADER_FONT; c2.fill = PCT_HEADER_FILL; c2.alignment = CENTER
        pct_col[status] = col_ptr + 1
        col_ptr += 2
    gt_col = col_ptr
    c = ws.cell(row=HDR_ROW, column=gt_col, value='Grand Total')
    c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = CENTER
    for col in range(1, ncols + 1):
        ws.cell(row=HDR_ROW, column=col).border = BORDER

    row = HDR_ROW + 1

    def write_row(label, indent, criteria_extra, is_subtotal=False, is_taluka=False, is_grand=False):
        nonlocal row
        if not is_grand:
            ws.merge_cells(f'A{row}:B{row}')
        cell = ws.cell(row=row, column=1, value=('    ' * indent) + label)
        if is_taluka:
            cell.font = TALUKA_FONT
            for col in range(1, ncols + 1):
                ws.cell(row=row, column=col).fill = TALUKA_FILL
        elif is_grand:
            cell.font = GRAND_FONT
            for col in range(1, ncols + 1):
                ws.cell(row=row, column=col).fill = GRAND_FILL
        elif is_subtotal:
            cell.font = SUBTOTAL_FONT
            for col in range(1, ncols + 1):
                ws.cell(row=row, column=col).fill = SUBTOTAL_FILL
        else:
            cell.font = NORMAL_FONT
        cell.alignment = Alignment(horizontal='left', indent=indent)

        for status in status_order:
            parts = [f'{RAW_STATUS},"{status}"']
            for rng, val in criteria_extra:
                parts.append(f'{rng},"{val}"')
            formula = '=COUNTIFS(' + ','.join(parts) + ')'
            cc = count_col[status]
            c = ws.cell(row=row, column=cc, value=formula)
            c.font = GRAND_FONT if is_grand else (SUBTOTAL_FONT if (is_subtotal or is_taluka) else NORMAL_FONT)
            c.alignment = CENTER
            count_ref = f'{get_column_letter(cc)}{row}'
            gt_ref = f'{get_column_letter(gt_col)}{row}'
            pcell = ws.cell(row=row, column=pct_col[status], value=f'=IFERROR({count_ref}/{gt_ref},0)')
            pcell.number_format = '0.0%'
            pcell.font = PCT_FONT if not (is_grand or is_subtotal or is_taluka) else (
                GRAND_FONT if is_grand else SUBTOTAL_FONT)
            pcell.alignment = CENTER

        count_refs = '+'.join(f'{get_column_letter(count_col[s])}{row}' for s in status_order)
        c = ws.cell(row=row, column=gt_col, value=f'={count_refs}')
        c.font = GRAND_FONT if is_grand else (SUBTOTAL_FONT if (is_subtotal or is_taluka) else NORMAL_FONT)
        c.alignment = CENTER
        for col in range(1, ncols + 1):
            ws.cell(row=row, column=col).border = BORDER
        this_row = row
        row += 1
        return this_row

    for tal in talukas_order:
        write_row(tal, 0, [(RAW_TAL, tal)], is_taluka=True)
        write_row('PUBLIC', 1, [(RAW_TAL, tal), (RAW_TYPE, 'Public')], is_subtotal=True)
        for phc in sorted(struct[tal]['Public']):
            write_row(phc, 2, [(RAW_TAL, tal), (RAW_TYPE, 'Public'), (RAW_GROUP, phc)])
        if struct[tal]['Private']:
            write_row('PRIVATE', 1, [(RAW_TAL, tal), (RAW_TYPE, 'Private')], is_subtotal=True)
            for fac in sorted(struct[tal]['Private']):
                write_row(fac, 2, [(RAW_TAL, tal), (RAW_TYPE, 'Private'), (RAW_GROUP, fac)])

    gt_row = write_row('GRAND TOTAL (ALL TALUKAS)', 0, [], is_grand=True)

    ws.column_dimensions['A'].width = 34
    ws.column_dimensions['B'].width = 4
    for status in status_order:
        ws.column_dimensions[get_column_letter(count_col[status])].width = 11
        ws.column_dimensions[get_column_letter(pct_col[status])].width = 10
    ws.column_dimensions[get_column_letter(gt_col)].width = 14
    ws.freeze_panes = f'C{HDR_ROW+1}'
    ws.sheet_view.showGridLines = False

    r = row + 1
    if note_lines:
        for line in note_lines:
            ws.cell(row=r, column=1, value=line)
            ws.cell(row=r, column=1).font = Font(italic=True, size=9, color='808080', name='Arial')
            ws.merge_cells(f'A{r}:{last_col_letter}{r}')
            r += 1

    ws.print_area = f'A1:{last_col_letter}{r}'
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws, gt_row


def build_linelist_sheet(wb, sheet_name, rows, idx, filter_col, filter_val, extra_cols,
                          eff_taluka_fn, group_name_fn, args):
    ws = wb.create_sheet(sheet_name)
    headers = ['Taluka', 'Type', 'PHC Name (Public) / Facility (Private)', 'Current Facility (Raw)'] + extra_cols
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=j, value=h)
        c.font = LL_HEADER_FONT; c.fill = LL_HEADER_FILL
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    r_i = 2
    for r in rows:
        if str(r[idx[filter_col]]) != filter_val:
            continue
        tal_raw = str(r[idx[args.taluka_col]])
        fac = str(r[idx[args.facility_col]])
        ptype = 'Public' if str(r[idx[args.type_col]]).strip().upper() == args.public_value.upper() else 'Private'
        et = eff_taluka_fn(tal_raw, ptype)
        gn = group_name_fn(tal_raw, fac, ptype)
        row_vals = [et, ptype, gn, fac] + [r[idx[col]] for col in extra_cols]
        for j, v in enumerate(row_vals, start=1):
            ws.cell(row=r_i, column=j, value=v)
        r_i += 1

    last_row = r_i - 1
    if last_row >= 2:
        ws.auto_filter.ref = f'A1:{get_column_letter(len(headers))}{last_row}'
    ws.freeze_panes = 'A2'
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 9
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 26
    for j in range(5, len(headers) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 16
    for row_cells in ws.iter_rows(min_row=2, max_row=max(last_row, 2), max_col=len(headers)):
        for c in row_cells:
            c.font = Font(size=10, name='Arial')
    ws.cell(row=last_row + 2, column=1, value=f'Total {filter_col}={filter_val} records: {max(last_row - 1, 0)}')
    ws.cell(row=last_row + 2, column=1).font = Font(bold=True, italic=True, size=9, name='Arial')
    return ws, last_row - 1


def parse_linelist_spec(spec):
    # format: "FilterColumn=Value:SheetName:Col1,Col2,Col3"
    cond, sheet_name, cols = spec.split(':', 2)
    col, val = cond.split('=', 1)
    return col.strip(), val.strip(), sheet_name.strip(), [c.strip() for c in cols.split(',') if c.strip()]


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('--register', required=True)
    p.add_argument('--register-sheet', required=True)
    p.add_argument('--hf-register', required=True)
    p.add_argument('--hf-sheet', default='Sheet1')
    p.add_argument('--taluka-col', required=True)
    p.add_argument('--facility-col', required=True)
    p.add_argument('--type-col', required=True, help='Column that distinguishes Public vs Private')
    p.add_argument('--public-value', required=True, help='Value in --type-col meaning "Public"')
    p.add_argument('--status-col', required=True)
    p.add_argument('--status-order', required=True, help='Comma-separated status categories, in column order')
    p.add_argument('--nonfunctional-taluka', default=None,
                   help='Taluka name that is non-functional for one sector (e.g. KADANA)')
    p.add_argument('--reassign-to', default=None, help='Taluka to reassign those records to (e.g. LUNAWADA)')
    p.add_argument('--reassign-sector', default='Private', help='Sector the reassignment rule applies to')
    p.add_argument('--title', required=True)
    p.add_argument('--sheet-name', required=True)
    p.add_argument('--output', required=True)
    p.add_argument('--linelist', action='append', default=[],
                   help='Repeatable. Format: "FilterColumn=Value:SheetName:Col1,Col2,..."')
    p.add_argument('--note', action='append', default=[], help='Extra footnote line(s) on the pivot sheet')
    args = p.parse_args()

    status_order = [s.strip() for s in args.status_order.split(',') if s.strip()]
    header, idx, rows = load_register(args.register, args.register_sheet)
    hf_rows, hf_map = load_hf_map(args.hf_register, args.hf_sheet)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    raw, type_col, group_col, eff_col, raw_last_row = build_raw_sheet(wb, header, idx, rows, args)
    write_hf_map_sheet(wb, hf_rows)

    struct, talukas_order, eff_taluka_fn, group_name_fn = python_side_groups(idx, rows, args, hf_map)

    tal_l = get_column_letter(eff_col)
    type_l = get_column_letter(type_col)
    group_l = get_column_letter(group_col)
    status_l = get_column_letter(idx[args.status_col] + 1)

    notes = list(args.note)
    if args.nonfunctional_taluka and args.reassign_to:
        notes.append(
            f'Note: Taluka {args.nonfunctional_taluka.upper()} is non-functional for the '
            f'{(args.reassign_sector or "Private")} sector, so matching records are regrouped here '
            f'under {args.reassign_to.upper()} (see Effective_Taluka in RAW DATA).'
        )
    notes.append(
        'Note: "UNMAPPED" under PUBLIC means the facility name in the register did not exactly '
        'match any Health Facility name in the HF Register for that Taluka — verify spelling/'
        'numbering in one of the two source files.'
    )
    notes.append(
        'Note: The "%" column next to each status shows that status\'s count as a percentage of '
        'the row\'s own Grand Total.'
    )

    build_pivot_sheet(wb, args.sheet_name, args.title, status_order, raw_last_row,
                       tal_l, type_l, group_l, status_l, struct, talukas_order, note_lines=notes)

    for spec in args.linelist:
        filter_col, filter_val, sheet_name, extra_cols = parse_linelist_spec(spec)
        build_linelist_sheet(wb, sheet_name, rows, idx, filter_col, filter_val, extra_cols,
                              eff_taluka_fn, group_name_fn, args)

    wb.save(args.output)
    print(f'Saved {args.output} — sheets: {wb.sheetnames}')
    print('Now run recalc, e.g.:')
    print(f'  python3 /mnt/skills/public/xlsx/scripts/recalc.py {args.output} 150')


if __name__ == '__main__':
    main()
