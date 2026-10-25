#!/usr/bin/env python3
"""Mahisagar TPT report builder.
Usage: python build_tpt_report.py --input master.xlsx --start 2026-01-01 --end 2026-10-31 --outdir /mnt/user-data/outputs
Sheets written: Summary, Patients, Contacts, ZERO PERSONS, TREATMENT NOT INITIATED, Notes
"""
import argparse, os
import pandas as pd, numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
ap=argparse.ArgumentParser()
ap.add_argument('--input',required=True);ap.add_argument('--start',required=True);ap.add_argument('--end',required=True)
ap.add_argument('--register',default=None,help='optional TPT REGISTER workbook: adds the month-wise register sheet');ap.add_argument('--outdir',default='.');ap.add_argument('--sheet',default=None,help='default: NOTIFICATION VS TPT, else FINAL SHEET FOR TPT ANALYSIS')
a=ap.parse_args()
f=a.input;START,END=a.start,a.end
xl=pd.ExcelFile(f)
sheet=a.sheet or ('NOTIFICATION VS TPT' if 'NOTIFICATION VS TPT' in xl.sheet_names else 'FINAL SHEET FOR TPT ANALYSIS')
df=pd.read_excel(xl,sheet_name=sheet)
if 'TPT.Spectrum_Enrollment_Date' in df.columns: P='TPT.'
else:
    pn=[c for c in df.columns if c.endswith('Person Name')][0];P=pn[:-len('Person Name')]
df['dx0']=0

df['dx']=pd.to_datetime(df['Spectrum_Diagnosis_Date'],errors='coerce')
df['tpt_start']=pd.to_datetime(df[P+'Spectrum_Treatment_initiation_Date'],errors='coerce')
d=df[(df['Site_of_TBDisease'].astype(str).str.strip()=='Pulmonary')&(df.dx>=START)&(df.dx<=END)].copy()
d['month']=d.dx.dt.strftime('%b-%y'); d['mnum']=d.dx.dt.month
d['taluka']=d['Spectrum_Current_TBU'].str.strip()
d['fac']=d['Spectrum_Current_HF_Type'].astype(str).str.strip().str.title()   # skill v3.0 rule 2: CURRENT HF type
d['enr']=pd.to_datetime(d[P+'Spectrum_Enrollment_Date'],errors='coerce') if P+'Spectrum_Enrollment_Date' in d.columns else pd.NaT
d['has_c']=d[P+'Person Name'].notna()      # skill v3.0 rule 3: persons attached = Person Name NOT NULL
d['given']=d.tpt_start.notna()&d.has_c
base=d.drop_duplicates('Episode_ID')[['Episode_ID','Patient_Name','month','mnum','taluka','fac','Name of PHC(HWC)','Spectrum_Current_HF','PrimaryPhone']]
n=d.groupby('Episode_ID').agg(n_contacts=('has_c','sum'),n_given=('given','sum')).reset_index()
pt=base.merge(n,on='Episode_ID'); pt['zero']=pt.n_contacts==0
pt['Name of PHC(HWC)']=pt['Name of PHC(HWC)'].fillna('Not mapped'); pt['Spectrum_Current_HF']=pt['Spectrum_Current_HF'].fillna('Not mapped')
months=[p.strftime('%b-%y') for p in pd.period_range(START,END,freq='M')]
months=[m for m in months if m in set(pt.month)]
# contacts
c=d[d.has_c].copy()
c['Spectrum_Enrollment_Date']=c['enr']

# styles
bl=PatternFill('solid',fgColor='1F4E78');lb=PatternFill('solid',fgColor='BDD7EE');gr=PatternFill('solid',fgColor='C6EFCE');rd=PatternFill('solid',fgColor='FFC7CE')
lg=PatternFill('solid',fgColor='E2F0D9');ly=PatternFill('solid',fgColor='FFF2CC')
W=Font(bold=True,color='FFFFFF');B=Font(bold=True);ctr=Alignment(horizontal='center',vertical='center',wrap_text=True)
th=Side(style='thin',color='999999');bd=Border(left=th,right=th,top=th,bottom=th)
def hdr(ws,row,vals,fill=bl,font=W):
    for i,v in enumerate(vals,1):
        x=ws.cell(row,i,v);x.fill=fill;x.font=font;x.alignment=ctr;x.border=bd
def widths(ws,w):
    for i,x in enumerate(w,1): ws.column_dimensions[L(i)].width=x

wb=Workbook(); ws=wb.active; ws.title='Summary'
mets=['Pulmonary Patients','Zero Contact','TPT Contacts','Enrolment %','Given TPT','TPT Given %']
cols=months+['Total']
ws.cell(1,1,f'Mahisagar District – TPT Notification Summary (Pulmonary TB patients diagnosed {pd.Timestamp(START).strftime('%b')}–{pd.Timestamp(END).strftime('%b %Y')})').font=Font(bold=True,size=14)
ws.cell(2,1,'Taluka').fill=bl;ws.cell(2,2,'Facility').fill=bl
ws.cell(1,12,'Pulmonary ONLY | Extra Pulmonary EXCLUDED | Facility Type = Spectrum_Current_HF_Type').font=Font(italic=True)
hdr(ws,3,['Taluka','Facility']+mets*len(cols))
for i,m in enumerate(cols):
    c0=3+i*6; ws.merge_cells(start_row=2,start_column=c0,end_row=2,end_column=c0+5)
    x=ws.cell(2,c0,m);x.fill=bl;x.font=W;x.alignment=ctr
ws.cell(2,1).font=W;ws.cell(2,2).font=W
tal=['Balasinor','Kadana','Khanpur','Lunawada','Santrampur','Virpur']
rows=[('Mahisagar Total','All',None,None),('All - Public','Public',None,'Public'),('All - Private','Private',None,'Private')]
rows+=[(t,'All',t,None) for t in tal]+[(t,'Public',t,'Public') for t in tal]+[(t,'Private',t,'Private') for t in tal]
rows=rows[:3]+[(t,'All',t,None) for t in tal]+[(t,'Public',t,'Public') for t in tal]+[(t,'Private',t,'Private') for t in tal]
def mk(sub,m):
    s=sub if m=='Total' else sub[sub.month==m]
    return s
r=4
for lab,fl,t,fc in rows:
    sub=pt
    if t: sub=sub[sub.taluka==t]
    if fc: sub=sub[sub.fac==fc]
    rowfill=lb if (fl=='All' and lab!='Mahisagar Total' and t) else (lg if fl=='Public' else ly if fl=='Private' else None)
    ws.cell(r,1,lab);ws.cell(r,2,fl)
    for cc in (1,2):
        ws.cell(r,cc).border=bd
        if rowfill: ws.cell(r,cc).fill=rowfill
        ws.cell(r,cc).font=B
    for i,m in enumerate(cols):
        s=mk(sub,m);np_=len(s);z=int(s.zero.sum());nc=int(s.n_contacts.sum());ng=int(s.n_given.sum())
        en=nc/(np_*3)*100 if np_ else None; gv=ng/nc*100 if nc else None
        vals=[np_,z,nc,None if en is None else round(en,1),ng,None if gv is None else round(gv,1)]
        for j,v in enumerate(vals):
            x=ws.cell(r,3+i*6+j,v if v is not None else '-');x.border=bd;x.alignment=Alignment(horizontal='center')
            if rowfill: x.fill=rowfill
            if j==3 and en is not None:
                if en>=66: x.fill=gr
                elif en<33: x.fill=rd
            if j==5 and gv is not None:
                if gv>=90: x.fill=gr
                elif gv<80: x.fill=rd
            if j in(3,5) and v is not None: x.number_format='0.0"%"'
    r+=1
ws.freeze_panes='C4';widths(ws,[18,10]+[11]*(6*len(cols)));ws.row_dimensions[3].height=32

# Patients
w2=wb.create_sheet('Patients');hdr(w2,1,['Episode ID','Patient Name','Month','Taluka','Facility Type (Current HF)','PHC / HWC','Health Facility (AAM)','Contacts Enrolled','TPT Given','Zero Contact'])
for i,x in enumerate(pt.sort_values(['taluka','mnum']).itertuples(),2):
    for j,v in enumerate([x.Episode_ID,x.Patient_Name,x.month,x.taluka,x.fac,x[7],x[8],int(x.n_contacts),int(x.n_given),'YES' if x.zero else 'No'],1):
        y=w2.cell(i,j,v)
        if x.zero: y.fill=rd
w2.auto_filter.ref=f'A1:J{len(pt)+1}';w2.freeze_panes='A2';widths(w2,[16,26,10,14,16,26,30,16,12,13])

# Contacts
w3=wb.create_sheet('Contacts');hdr(w3,1,['Index Episode ID','Contact Name','Age','Gender','Key Population','Enrolment Date','TPT Start Date','Regimen','Outcome','Month','Taluka','Facility'])
c=c.sort_values(['mnum','taluka'])
for i,x in enumerate(c.itertuples(index=False),2):
    xs=x._asdict() if hasattr(x,'_asdict') else None
rows_c=c[['Episode_ID',P+'Person Name',P+'Age',P+'Gender',P+'KeyPopulation','Spectrum_Enrollment_Date','tpt_start',P+'Regimen Type',P+'Treatment_Outcome','month','taluka','fac']].values.tolist()
for i,row in enumerate(rows_c,2):
    for j,v in enumerate(row,1):
        if isinstance(v,float) and np.isnan(v): v=None
        if isinstance(v,pd.Timestamp): v=v.to_pydatetime().date()
        if v is pd.NaT: v=None
        w3.cell(i,j,v)
w3.auto_filter.ref=f'A1:L{len(rows_c)+1}';w3.freeze_panes='A2';widths(w3,[16,28,6,10,18,14,14,10,28,9,13,10])

# ZERO PERSONS
w4=wb.create_sheet('ZERO PERSONS');w4.sheet_properties.tabColor='FF0000'
z=pt[pt.zero].sort_values(['taluka','Name of PHC(HWC)','mnum'])
w4.merge_cells('A1:I1');t1=w4.cell(1,1,f'⚠ ACTION REQUIRED – {len(z)} Patients with ZERO Contacts Identified');t1.font=Font(bold=True,size=14,color='FFFFFF');t1.fill=PatternFill('solid',fgColor='C00000');t1.alignment=ctr
hdr(w4,2,['Sr.','Episode ID','Patient Name','Month','Taluka','Facility Type','PHC / HWC Name','Facility (Sub-centre / AAM)','Phone'])
r=3;sr=1
for t,g in z.groupby('taluka'):
    w4.merge_cells(start_row=r,start_column=1,end_row=r,end_column=9)
    x=w4.cell(r,1,f'{t}  –  {len(g)} patients');x.fill=lb;x.font=B;r+=1
    for x in g.itertuples():
        for j,v in enumerate([sr,x.Episode_ID,x.Patient_Name,x.month,x.taluka,x.fac,x[7],x[8],x.PrimaryPhone],1):
            y=w4.cell(r,j,v);y.border=bd
        w4.cell(r,6).fill=lg if x.fac=='Public' else ly
        r+=1;sr+=1
w4.freeze_panes='A3';widths(w4,[6,16,28,10,13,13,26,30,14])


# TREATMENT NOT INITIATED (contacts enrolled but TPT start date blank)
w6=wb.create_sheet('TREATMENT NOT INITIATED',4);w6.sheet_properties.tabColor='FFC000'
ni=d[d.has_c&d.tpt_start.isna()].copy()
ni['Spectrum_Enrollment_Date']=ni['enr'];ni=ni.sort_values(['taluka','Name of PHC(HWC)','mnum','Episode_ID'])
w6.merge_cells('A1:N1');t6=w6.cell(1,1,f'ACTION REQUIRED – {len(ni)} Enrolled Contacts with TPT / Treatment NOT Initiated');t6.font=Font(bold=True,size=14,color='FFFFFF');t6.fill=PatternFill('solid',fgColor='C55A11');t6.alignment=ctr
hdr(w6,2,['Sr.','Index Episode ID','Index Patient','Index Phone','Month','Taluka','Facility Type','PHC / HWC','Contact Name','Age','Gender','Contact Phone','Contact Enrolment Date','Contact Outcome'])
for i,x in enumerate(ni.to_dict('records'),1):
    vals=[i,x['Episode_ID'],x['Patient_Name'],x['PrimaryPhone'],x['month'],x['taluka'],x['fac'],x['Name of PHC(HWC)'] if pd.notna(x['Name of PHC(HWC)']) else 'Not mapped',x[P+'Person Name'],x[P+'Age'],x[P+'Gender'],x.get(P+'Primary Phone'),x['Spectrum_Enrollment_Date'],x[P+'Treatment_Outcome']]
    for j,v in enumerate(vals,1):
        if isinstance(v,float) and np.isnan(v): v=None
        if isinstance(v,pd.Timestamp): v=v.to_pydatetime().date()
        if v is pd.NaT: v=None
        w6.cell(2+i,j,v).border=bd
w6.auto_filter.ref=f'A2:N{len(ni)+2}';w6.freeze_panes='A3';widths(w6,[6,16,26,13,9,13,12,24,26,6,9,13,14,30])
idx_ni=pt_all=d.drop_duplicates('Episode_ID')
print('contacts not initiated',len(ni),'index pts without own tx start',int(idx_ni['Spectrum_Treatment_Initiation_Date'].isna().sum()), ni[P+'Treatment_Outcome'].value_counts(dropna=False).to_dict())

# Notes
w5=wb.create_sheet('Notes')
tot=pt;nc=int(tot.n_contacts.sum())
notes=['Mahisagar TPT Notification Summary – Methodology','',
f'Source: {os.path.basename(f)}, sheet "{sheet}".',
f'Period: Spectrum_Diagnosis_Date from {START} to {END}. Latest month is partial if data ends mid-month.',
'Filter: Site_of_TBDisease = "Pulmonary" (strict equality; Extra Pulmonary, blank and NaN excluded from numerator and denominator).',
f'Generated: {pd.Timestamp.today():%d-%b-%Y}','',
'Pulmonary Patients = distinct Episode_ID. Zero Contact = patients with no contact enrolled.',
'TPT Contacts (Persons Attached) = rows where TPT.Person Name is not null.',
'Enrolment % = TPT Contacts ÷ (Patients × 3) × 100.  Given TPT = contacts with TPT treatment initiation date.  TPT Given % = Given ÷ Contacts × 100.',
'Colour key: Enrolment % green ≥66, red <33; TPT Given % green ≥90, red <80. Blue = taluka rows, light green = Public, light yellow = Private.',
'Month = month of index patient diagnosis. Taluka = Spectrum_Current_TBU. Facility type = CURRENT health facility type, Spectrum_Current_HF_Type (Public/Private).',
'TREATMENT NOT INITIATED sheet = enrolled contacts whose TPT treatment initiation date is blank (line list with index patient, PHC and contact details). ZERO PERSONS = index patients with no contact enrolled.','Recent months may look low because of pending data entry.']
for i,t in enumerate(notes,1): w5.cell(i,1,t)
w5['A1'].font=Font(bold=True,size=14);w5.column_dimensions['A'].width=150
# ---------------- TPT REGISTER month-wise sheet (optional) ----------------
if a.register:
    rx=pd.ExcelFile(a.register)
    reg=pd.read_excel(rx,sheet_name='TPT REGISTER' if 'TPT REGISTER' in rx.sheet_names else rx.sheet_names[0])
    reg['enr']=pd.to_datetime(reg['Spectrum_Enrollment_Date'],errors='coerce')
    reg=reg[(reg.enr>=START)&(reg.enr<=END)].copy()
    reg['linked']=reg['Episode ID Index'].notna()                                   # linked to an index patient
    reg['tpt']=reg['Type of treatment'].astype(str).str.strip().str.upper().eq('TPT')  # TPT actually given
    reg['mper']=reg.enr.dt.to_period('M')
    other_tx=int(((reg['Type of treatment'].astype(str).str.strip().str.upper().isin(['DSTB','DRTB']))).sum())
    wr=wb.create_sheet('TPT REGISTER MONTHWISE',1);wr.sheet_properties.tabColor='7030A0'
    wr.cell(1,1,'TPT Register – contacts enrolled and TPT given, month-wise (month of contact enrolment)').font=Font(bold=True,size=14)
    wr.cell(2,1,'All contacts in the TPT register, split by whether the contact is linked to an index patient (Episode ID Index present) or not linked (blank).').font=Font(italic=True)
    groups=[('Month',1,1),('All contacts enrolled',2,4),('Linked to an index patient',5,7),('NOT linked to an index patient',8,10),('Not linked as % of enrolled',11,11)]
    for lab,c1,c2 in groups:
        if c2>c1: wr.merge_cells(start_row=3,start_column=c1,end_row=3,end_column=c2)
        x=wr.cell(3,c1,lab);x.fill=bl;x.font=W;x.alignment=ctr
        for cc in range(c1,c2+1): wr.cell(3,cc).fill=bl;wr.cell(3,cc).border=bd
    hdr(wr,4,['']+['Enrolled','TPT given','TPT given %']*3+['% of enrolled'])
    wr.merge_cells('A3:A4');wr.merge_cells('K3:K4')
    def seg(g): n=len(g);t=int(g.tpt.sum());return n,t,(round(100*t/n,1) if n else None)
    def put(r,lab,g,fill=None,bold=False):
        vals=[lab]
        for sub in (g,g[g.linked],g[~g.linked]): vals+=list(seg(sub))
        vals.append(round(100*(~g.linked).sum()/len(g),1) if len(g) else None)
        for j,v in enumerate(vals,1):
            x=wr.cell(r,j,v if v is not None else '-');x.border=bd
            if j>1: x.alignment=Alignment(horizontal='center')
            if fill: x.fill=fill
            if bold: x.font=B
            if j in (4,7,10) and v is not None:
                x.number_format='0.0"%"'
                if not fill: x.fill=gr if v>=90 else (rd if v<80 else ly)
            if j==11 and v is not None: x.number_format='0.0"%"'
    r=5
    last=reg.mper.max()
    for m in sorted(reg.mper.dropna().unique()):
        put(r,m.strftime('%b-%y')+('*' if m==last and pd.Timestamp(END)>pd.Timestamp.today() else ''),reg[reg.mper==m]);r+=1
    put(r,'TOTAL',reg,fill=lb,bold=True)
    r+=2
    for t in ['Rows = month of Spectrum_Enrollment_Date of the contact. Contacts enrolled = every row of the TPT REGISTER sheet in the period (count of '+str(len(reg))+').',
              'Linked = Episode ID Index is filled (contact attached to an index patient). Not linked = Episode ID Index is blank (contact enrolled directly).',
              'TPT given = Type of treatment is TPT. TPT given % = TPT given / contacts enrolled in that month. Colour: green >=90%, red <80%.',
              f'{other_tx} contacts in the register started DSTB/DRTB treatment (found to have TB) – they are enrolled but NOT counted as TPT given.',
              'This sheet counts every contact in the register; the Summary sheet counts only contacts of pulmonary index patients, so its figures differ.',
              '* latest month is part-month (enrolments up to the register date).']:
        wr.cell(r,1,t).font=Font(italic=True,size=9);r+=1
    widths(wr,[12,11,11,12,11,11,12,11,11,12,13]);wr.freeze_panes='B5';wr.row_dimensions[3].height=34;wr.row_dimensions[4].height=22
    print('register sheet:',len(reg),'contacts; linked',int(reg.linked.sum()),'not linked',int((~reg.linked).sum()),'TPT given',int(reg.tpt.sum()))

os.makedirs(a.outdir,exist_ok=True)
out=os.path.join(a.outdir,f"Mahisagar_TPT_Report_{pd.Timestamp(START).strftime('%b')}_{pd.Timestamp(END).strftime('%b_%Y')}.xlsx");wb.save(out);print('Saved',out)
print(d.taluka.value_counts().to_dict());print(len(pt),int(pt.zero.sum()),nc,int(pt.n_given.sum()),months)
print(pt.groupby('month').agg(p=('Episode_ID','count'),z=('zero','sum'),c=('n_contacts','sum'),g=('n_given','sum')).loc[months])
print(c.Spectrum_Enrollment_Date.notna().mean(), (pt['Name of PHC(HWC)']=='Not mapped').sum())
