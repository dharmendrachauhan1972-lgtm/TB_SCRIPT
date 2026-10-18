#!/usr/bin/env python3
"""Build the two Mahisagar TB Notification Registers (month-wise) from a NIKSHAY/Spectrum
notification workbook.

  1. Diagnosing-facility base  (sheet "DIAGNOSIS FACILITY WISE")
  2. Enrolment-facility base   (sheet "ENROLMENT FACILITY WISE", already joined to HF REGISTER)

Both are written as separate .xlsx files with live COUNTIFS formulas on a Data sheet.
Run recalc.py (xlsx skill) afterwards so the formulas get cached values.

Usage:
  python build_notification_registers.py --input NOTIFICATION.xlsx --outdir outputs
"""
import argparse, os
import pandas as pd, numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L

th=Side(style='thin',color='999999');bd=Border(left=th,right=th,top=th,bottom=th)
DK=PatternFill('solid',fgColor='1F4E78');LB=PatternFill('solid',fgColor='BDD7EE');GR=PatternFill('solid',fgColor='D9D9D9');YL=PatternFill('solid',fgColor='FFF2CC')
W=Font(bold=True,color='FFFFFF');B=Font(bold=True)
def q(s): return str(s).replace('"','""')

def build(df,title,out,notes):
    df=df.copy()
    df['dx']=pd.to_datetime(df['Spectrum_Diagnosis_Date'])
    df['month']=df.dx.dt.strftime('%b-%y')
    mp=pd.period_range(df.dx.min(),df.dx.max(),freq='M')
    months=[p.strftime('%b-%y') for p in mp]
    years=sorted(df.dx.dt.year.unique())
    wb=Workbook();ws=wb.active;ws.title='Notification Register'
    wd=wb.create_sheet('Data')
    cols=['Episode_ID','Patient_Name','Spectrum_Diagnosis_Date','Month','Type','Taluka','PHI / Facility']
    for j,c in enumerate(cols,1):
        x=wd.cell(1,j,c);x.font=W;x.fill=DK
    for i,r in enumerate(df.itertuples(),2):
        wd.cell(i,1,r.Episode_ID);wd.cell(i,2,r.Patient_Name);wd.cell(i,3,r.dx.date());wd.cell(i,4,r.month);wd.cell(i,5,r.type_);wd.cell(i,6,r.tal);wd.cell(i,7,r.phi)
    n=len(df)+1
    wd.auto_filter.ref=f'A1:G{n}';wd.freeze_panes='A2'
    for j,w in enumerate([14,28,14,9,10,24,42],1): wd.column_dimensions[L(j)].width=w
    R=lambda col:f'Data!${col}$2:${col}${n}'
    nm=len(months);ycol0=2+nm;gcol=ycol0+len(years)
    ws.cell(1,1,title).font=Font(bold=True,size=14)
    ws.cell(2,1,'Count of notified TB patients by month of diagnosis (Spectrum_Diagnosis_Date). Counts are live COUNTIFS formulas on the Data sheet.').font=Font(italic=True)
    hr=4
    hd=['Row Labels']+months+[f'{y} Total' for y in years]+['Grand Total']
    for j,h in enumerate(hd,1):
        x=ws.cell(hr,j,h);x.font=W;x.fill=DK;x.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True);x.border=bd
    r=hr+1;type_rows=[]
    def fill_totals(row,kind,rows_or_range):
        for j in range(2,2+nm):
            c=L(j)
            ws.cell(row,j,'='+('+'.join(f'{c}{x}' for x in rows_or_range) if kind=='rows' else f'SUM({c}{rows_or_range[0]}:{c}{rows_or_range[1]})'))
    def year_grand(row):
        for k,y in enumerate(years):
            idx=[i for i,m in enumerate(months) if m.endswith(str(y)[2:])]
            ws.cell(row,ycol0+k,f'=SUM({L(2+idx[0])}{row}:{L(2+idx[-1])}{row})')
        ws.cell(row,gcol,f'=SUM({L(2)}{row}:{L(1+nm)}{row})')
    def style(row,fill,font=None,indent=0):
        for j in range(1,gcol+1):
            x=ws.cell(row,j);x.border=bd
            if fill: x.fill=fill
            if font: x.font=font
            if j>1: x.alignment=Alignment(horizontal='center')
        ws.cell(row,1).alignment=Alignment(indent=indent)
    for t in ['Public','Private']:
        dt=df[df.type_==t]
        if dt.empty: continue
        trow=r;ws.cell(r,1,t);r+=1;tal_rows=[]
        tals=sorted(dt.tal.unique(),key=lambda s:(s.startswith('('),s.lower()))
        for tl in tals:
            dtl=dt[dt.tal==tl];arow=r;ws.cell(r,1,tl);r+=1;p0=r
            for p in sorted(dtl.phi.unique(),key=str.lower):
                ws.cell(r,1,p)
                for j,m in enumerate(months):
                    ws.cell(r,2+j,f'=COUNTIFS({R("E")},"{q(t)}",{R("F")},"{q(tl)}",{R("G")},"{q(p)}",{R("D")},"{m}")')
                year_grand(r);style(r,None,None,3);r+=1
            fill_totals(arow,'range',(p0,r-1));year_grand(arow);style(arow,LB,B,1);tal_rows.append(arow)
        fill_totals(trow,'rows',tal_rows);year_grand(trow);style(trow,GR,B,0);type_rows.append(trow)
        # make type row stand out
        for j in range(1,gcol+1): ws.cell(trow,j).fill=YL
    gr=r;ws.cell(gr,1,'Grand Total');fill_totals(gr,'rows',type_rows);year_grand(gr);style(gr,DK,W)
    ws.cell(gr+2,1,'Check: rows in Data sheet');ws.cell(gr+2,2,len(df))
    ws.cell(gr+3,1,'Check: Grand Total − Data rows (must be 0)');ws.cell(gr+3,2,f'={L(gcol)}{gr}-B{gr+2}')
    ws.freeze_panes='B5';ws.column_dimensions['A'].width=46
    for j in range(2,gcol+1): ws.column_dimensions[L(j)].width=9 if j<gcol-len(years) else 11
    ws.row_dimensions[hr].height=30
    wn=wb.create_sheet('Notes')
    for i,t in enumerate(notes,1): wn.cell(i,1,t)
    wn['A1'].font=Font(bold=True,size=14);wn.column_dimensions['A'].width=140
    wb.save(out);return months


def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('--input',required=True)
    ap.add_argument('--outdir',default='.')
    ap.add_argument('--diag-sheet',default='DIAGNOSIS FACILITY WISE')
    ap.add_argument('--enrol-sheet',default='ENROLMENT FACILITY WISE')
    ap.add_argument('--remap-from',default='(Non functional) KADANA',help='Diagnosing TBU label to replace')
    ap.add_argument('--remap-to',default='Lunawada',help='Replacement TBU label')
    a=ap.parse_args()
    os.makedirs(a.outdir,exist_ok=True)
    def remap(df):
        m=df.Spectrum_Diagnosing_TBU.str.strip().str.lower()==a.remap_from.strip().lower()
        if m.any():
            print(f'Remapped {int(m.sum())} rows: TBU {a.remap_from!r} -> {a.remap_to!r}; facilities:')
            print(df.loc[m,'Spectrum_Diagnosing_PHI'].value_counts().to_string())
        df.loc[m,'Spectrum_Diagnosing_TBU']=a.remap_to
        return df
    note_remap=f'TBU "{a.remap_from}" has been replaced by "{a.remap_to}" (all its private facilities now sit under {a.remap_to}). Facility names are NOT renamed, so a hospital spelt two ways stays on two rows.'
    # File 1
    d=remap(pd.read_excel(a.input,sheet_name=a.diag_sheet))
    d['type_']=d.Spectrum_Modified_Diagnosing_facility_type.str.strip().str.title()
    d['tal']=d.Spectrum_Diagnosing_TBU.str.strip();d['phi']=d.Spectrum_Diagnosing_PHI.str.strip()
    n1=['Notification Register – Diagnosing Facility Base','',
        f'Source: sheet "{a.diag_sheet}" ({len(d):,} notified patients).',
        'Columns: month of notification = month of Spectrum_Diagnosis_Date.',
        'Row level 1: Spectrum_Modified_Diagnosing_facility_type (Public first, then Private).',
        'Row level 2 (Taluka): Spectrum_Diagnosing_TBU.  Row level 3 (sub-taluka): Spectrum_Diagnosing_PHI.',note_remap]
    build(d,'Notification Register – Diagnosing Facility Base (Mahisagar)',os.path.join(a.outdir,'Notification_Register_Diagnosing_Facility_Base.xlsx'),n1)
    # File 2
    e=remap(pd.read_excel(a.input,sheet_name=a.enrol_sheet))
    hf_tb=e['HF REGISTER.TB Unit'];hf_phc=e['HF REGISTER.Name of PHC(HWC)']
    blank=hf_tb.isna()|hf_phc.isna()
    e['type_']=np.where(blank,'Private','Public')
    e['tal']=np.where(blank,e.Spectrum_Diagnosing_TBU.str.strip(),hf_tb.str.strip())
    e['phi']=np.where(blank,e.Spectrum_Diagnosing_PHI.str.strip(),hf_phc.str.strip())
    outside=int((blank&(e.Spectrum_Modified_Diagnosing_facility_type.str.title()=='Public')).sum())
    n2=['Notification Register – Enrolment Facility Base','',
        f'Source: sheet "{a.enrol_sheet}" ({len(e):,} notified patients), already linked to HF REGISTER on enrolment facility.',
        'Columns: month of notification = month of Spectrum_Diagnosis_Date.',
        'Public block: Taluka = HF REGISTER.TB Unit; sub-taluka = HF REGISTER.Name of PHC(HWC) of the enrolment facility.',
        'Private block: used whenever the HF REGISTER match is blank. Taluka = Spectrum_Diagnosing_TBU; sub-taluka = Spectrum_Diagnosing_PHI.',
        f'Note: {int(blank.sum())} patients fell into the Private block (HF REGISTER match blank); {outside} of them were diagnosed at a Public facility but enrolled outside the district.',
        note_remap]
    build(e,'Notification Register – Enrolment Facility Base (Mahisagar)',os.path.join(a.outdir,'Notification_Register_Enrolment_Facility_Base.xlsx'),n2)
    print('Public/Private diag:',d.type_.value_counts().to_dict(),' enrol:',e.type_.value_counts().to_dict(),' outside-district public->private:',outside)
    print('Now run: python /mnt/skills/public/xlsx/scripts/recalc.py <file> 120 on each output.')

if __name__=='__main__': main()
