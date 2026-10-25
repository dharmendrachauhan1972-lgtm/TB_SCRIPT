"""Mahisagar district review - data extraction. Usage: python3 extract_data.py <uploads_dir> <out.json> <asof YYYY-MM-DD>
Takes ~4-5 min (large registers). Run in the FOREGROUND with timeout 295 (background jobs get killed)."""
import sys, json, os, pandas as pd, numpy as np, openpyxl
U=sys.argv[1] if len(sys.argv)>1 else '/mnt/user-data/uploads'; P=lambda f: os.path.join(U,f)
OUT=sys.argv[2] if len(sys.argv)>2 else '/tmp/R2.json'
T=['Balasinor','Kadana','Khanpur','Lunawada','Santrampur','Virpur']
ASOF=pd.Timestamp(sys.argv[3]) if len(sys.argv)>3 else pd.Timestamp.today().normalize()   # register extract date (pass YYYY-MM-DD)
def nz(s):
    s=str(s).strip().title()
    return {'Lunavada':'Lunawada','(Non Functional) Kadana':'Lunawada'}.get(s,s)
CUT25=ASOF-pd.DateOffset(years=1)   # same-period comparison with last year (works for a part-month)
R={'asof':str(ASOF.date())}
def rd(f,sh,**k): return pd.read_excel(P(f),sheet_name=sh,**k)

# ---------- A. presumptive / NAAT / microscopy / xray ----------
tg=rd('cbnaat_target.xlsx','cbnaat target'); tgt={nz(a):float(b) for a,b in zip(tg.taluka,tg['cbnaat target']) if 'mahisagar' not in str(a).lower()}
p=rd('PRESUMPTIVE_REGISTER_FOR_XRAY___NAAT__MICROSCOPY.xlsx','PHC WISE COMPILED')
p['tal']=p.Spectrum_Enrolment_TBU.map(nz); p['ed']=pd.to_datetime(p.Spectrum_Enrollment_Date,errors='coerce')
p['dn']=pd.to_datetime(p.Date_NAAT,errors='coerce'); p['dm']=pd.to_datetime(p.Date_Microscopy,errors='coerce')
Y=2026
def ach(kind):
    out={}
    for t in T:
        g=p[p.tal==t]; a=[]
        for m in range(1,13):
            if kind=='pres': c=((g.dn.dt.year==Y)&(g.dn.dt.month==m))|((g.dm.dt.year==Y)&(g.dm.dt.month==m))
            elif kind=='naat': c=(g.dn.dt.year==Y)&(g.dn.dt.month==m)
            else: c=(g.dm.dt.year==Y)&(g.dm.dt.month==m)
            a.append(int(c.sum()))
        out[t]={'t':tgt[t],'a':a}
    out['Total (District)']={'t':sum(tgt[t] for t in T),'a':[sum(out[t]['a'][i] for t in T) for i in range(12)]}
    return out
R['pres']=ach('pres'); R['naat']=ach('naat'); R['micro']=ach('micro')
q=p[(p.ed.dt.year==Y)].copy(); q['m']=q.ed.dt.month
R['xray']={'xray':{}}
for t in T+['District']:
    g=q if t=='District' else q[q.tal==t]
    R['xray']['xray'][t]=[[int((g[g.m==m].Xray_Offered=='Yes').sum()),int((g.m==m).sum())] for m in range(1,13)]
tested=q.dn.notna()|q.dm.notna(); abn=q[q.Result_XRay_Positive=='Yes']; off=q[q.Xray_Offered=='Yes']
sym=q.Symptom.fillna('').str.replace(',','').str.strip()=='Asymptomatic'
R['xray'].update({'enrolled':len(q),'tested':int(tested.sum()),'abn':len(abn),'abn_tested':int((abn.dn.notna()|abn.dm.notna()).sum()),
  'offered':len(off),'result_recorded':int(((off.Result_XRay_Positive=='Yes')|(off.Result_XRay_Negative=='Yes')).sum()),'asym_tested':int((sym&tested).sum()),'asym_n':int(sym.sum())})
print('A done',len(p),flush=True)

# ---------- B. notification (diag + enrol base) ----------
nt=rd('monthly_notification_target.xlsx','Sheet1'); ntm={nz(a):float(b) for a,b in zip(nt.taluka,nt['monthly target']) if 'mahisagar' not in str(a).lower()}
tu=rd('NOTIFICATION_TARGET_TU_WISE_MAHISAGAR_DISTRICT_YEAR_2026.xlsx',0); R['ntgt_pubpriv']={nz(a):[float(b),float(c)] for a,b,c in zip(tu.iloc[:,0],tu.iloc[:,1],tu.iloc[:,2]) if 'mahisagar' not in str(a).lower()}
R['ntgt_month']=ntm
dg=rd('NOTIFICATION.xlsx','DIAGNOSIS FACILITY WISE'); en=rd('NOTIFICATION.xlsx','ENROLMENT FACILITY WISE')
dg['type']=dg.Spectrum_Modified_Diagnosing_facility_type.astype(str).str.strip().str.title().replace({'Nan':'Private'}); dg['tal']=dg.Spectrum_Diagnosing_TBU.map(nz)
hfu=en['HF REGISTER.TB Unit']; pub=hfu.notna()&(hfu.astype(str).str.strip()!='')
en['type']=np.where(pub,'Public','Private'); en['tal']=np.where(pub,hfu.map(nz),en.Spectrum_Diagnosing_TBU.map(nz))
for d in (dg,en):
    d['dt']=pd.to_datetime(d.Spectrum_Diagnosis_Date); d['y']=d.dt.dt.year; d['m']=d.dt.dt.month
N={}
for k,d in (('diag',dg),('enrol',en)):
    N[k]={}
    for t in T+['District']:
        g=d if t=='District' else d[d.tal==t]
        N[k][t]={'m26':[int(((g.y==2026)&(g.m==m)).sum()) for m in range(1,13)],'ytd25':int(((g.y==2025)&(g['dt']<=CUT25)).sum()),
          'pub26':int(((g.y==2026)&(g.type=='Public')).sum()),'prv26':int(((g.y==2026)&(g.type=='Private')).sum()),
          'pub25ytd':int(((g.y==2025)&(g['dt']<=CUT25)&(g.type=='Public')).sum()),'prv25ytd':int(((g.y==2025)&(g['dt']<=CUT25)&(g.type=='Private')).sum()),
          'pubm':[int(((g.y==2026)&(g.m==m)&(g.type=='Public')).sum()) for m in range(1,13)],'prvm':[int(((g.y==2026)&(g.m==m)&(g.type=='Private')).sum()) for m in range(1,13)]}
N['target_annual']={t:ntm[t]*12 for t in T}; N['target_month']=ntm
a=dg[dg.y==2026][['Episode_ID','tal','type']].merge(en[en.y==2026][['Episode_ID','tal','type']],on='Episode_ID',suffixes=('_d','_e'))
N['out_of_diag']={t:int(((a.tal_d==t)&(a.tal_e!=t)).sum()) for t in T}; N['into_enrol']={t:int(((a.tal_e==t)&(a.tal_d!=t)).sum()) for t in T}
N['prv2pub']=int(((a.type_d=='Private')&(a.type_e=='Public')).sum())
R['notif']=N
print('B done',flush=True)

# ---------- C. current register: outcome / delay / not initiated ----------
d=rd('CURRENT_NOTIFICATION_REGISTER_FOR_TREATMENT_NOT_INITED_AND_OUTCOME_PENIDNGR.xlsx','CURRENT NOTIFICATION REGISTER')
d['dx']=pd.to_datetime(d.Spectrum_Diagnosis_Date); d['ti']=pd.to_datetime(d.Treatment_Initiation_Date); d['tal']=d.Spectrum_Current_TBU.map(nz)
REG=d.dx.max().normalize(); R['reg_asof']=str(REG.date()); R['reg_n']=len(d)
ds=d[(d.Type_of_TB_regimen=='2HRZE/4HRE')&(d.Site_of_TBDisease=='Pulmonary')].copy()
def cat(o):
    if pd.isna(o): return 'P'
    if o in('CURED','TREATMENT_COMPLETE'): return 'S'
    if o=='DIED': return 'D'
    if o=='TREATMENT_REGIMEN_CHANGED': return 'R'
    return 'O'
ds['c']=ds.Treatment_Outcome.map(cat)
cut=REG-pd.DateOffset(months=9); YR=cut.year if cut.month==12 else cut.year-1
co=ds[ds.dx.dt.year==YR]
R['out']={'cohort_year':int(YR),'tbu':{t:co[co.tal==t].c.value_counts().to_dict() for t in T},'all':co.c.value_counts().to_dict(),
  'mon':[ds[(ds.dx.dt.year==YR)&(ds.dx.dt.month==m)].c.value_counts().to_dict() for m in range(1,13)],
  'early':[ds[(ds.dx.dt.year==YR+1)&(ds.dx.dt.month==m)].c.value_counts().to_dict() for m in range(1,4)]}
a=d.dropna(subset=['ti']).copy(); a=a[a.dx>=f'{YR+1}-01-01']; a['gap']=(a.ti-a.dx).dt.days
R['gap']={'n':len(a),'late':int((a.gap>7).sum()),'med':float(a.gap.median()),'tbu':{t:[int((a[a.tal==t].gap>7).sum()),int((a.tal==t).sum())] for t in T}}
ni=d[d.ti.isna()&(d.dx>=f'{YR+1}-01-01')&d.Treatment_Outcome.isna()]; old=ni[ni.dx<REG-pd.Timedelta(days=7)]
R['ni']={'n':len(ni),'old':len(old),'tbu':ni.tal.value_counts().to_dict(),'tbu_old':old.tal.value_counts().to_dict()}
print('C done',flush=True)

# ---------- D. IP / CP ----------
lab=rd('PATIENT_LAB_REGI_IP_AND_CP_PENIDNG_LIST.xlsx','PATIENT LAB REGISTR')
lab['dx']=pd.to_datetime(lab.DiagnosisDate); lab['ti']=pd.to_datetime(lab.DateOf_TB_Treatment_Initiation)
l=lab[(lab.RegimenType=='2HRZE/4HRE')&(lab.SiteOfDisease=='Pulmonary')].copy()
l['unit']=l.CurrentFacilityTBU.astype(str).str.strip().str.upper().replace({'KADANA':'LUNAWADA'}).str.title()
def fup(col,n1,n2,start):
    e=l.copy(); gap=(ASOF-e.ti).dt.days
    due=(e.Treatment_Outcome.isna()&(gap>n1))|(e.Treatment_Outcome.notna()&(gap>n2)); e=e[due].copy(); e['done']=e[col].notna()
    e['cm']=e.dx.dt.to_period('M'); months=list(pd.period_range(start,ASOF.strftime('%Y-%m'),freq='M'))
    units=sorted(e.unit.unique()); res={'months':[str(m) for m in months],'units':{},'tot':{}}
    for u in units+['District']:
        g=e if u=='District' else e[e.unit==u]
        res['units'][u]=[[int((g.cm==m).sum()),int(g[g.cm==m].done.sum())] for m in months]
        res['tot'][u]=[len(g),int(g.done.sum())]
    res['all_due_not_done']=int((~e.done).sum()); return res
R['ip2']=fup('Microscopy_FollowUpOfDSTB_ENDIP_FinalInterpretation',55,58,'2026-01')
R['cp2']=fup('Microscopy_FollowUpOfDSTB_ENDCP_FinalInterpretation',165,168,'2025-09')
print('D done',flush=True)

# ---------- E. PTFU ----------
pt=rd('PTFU_CUMULATVE.xlsx','NOTIFICATION VS PTFU')
pt['od']=pd.to_datetime(pt.Date_of_Treatment_Outcome,errors='coerce'); pt['tal']=pt.Spectrum_Current_TBU.map(nz)
ok=pt[pt.Treatment_Outcome.isin(['CURED','TREATMENT_COMPLETE'])].copy()
R['ptfu_moc']=pt['PTFU.MonthOfCohort'].astype(str).value_counts().head(8).to_dict()
mc=ok['PTFU.MonthOfCohort'].astype(str).str.extract(r'(\d+)')[0]; ok['moc']=pd.to_numeric(mc,errors='coerce'); ok['resp']=ok['PTFU.Response'].notna()
PT={}
for mm in [(pd.Period(ASOF,'M')-k).strftime('%b-%y') for k in (3,2,1,0)]:   # latest 4 months up to the as-of month
    R_=pd.Period(pd.to_datetime('01-'+mm,format='%d-%b-%y'),'M'); rowsd={}
    for t in T+['District']:
        den=set(); num=set()
        for k in (6,12,18,24):
            c=R_-k; sel=ok[(ok.od.dt.to_period('M')==c)&((ok.tal==t) if t!='District' else True)]
            ids=set(sel.Episode_ID); den|={(i,k) for i in ids}
            num|={(i,k) for i in set(sel[(sel.moc==k)&sel.resp].Episode_ID)}
        rowsd[t]=[len(den),len(num)]
    PT[mm]=rowsd
R['ptfu']={t:[[mm,PT[mm][t][0],PT[mm][t][1],round(100*PT[mm][t][1]/PT[mm][t][0]) if PT[mm][t][0] else 0] for mm in PT] for t in T+['District']}
print('E done',flush=True)

# ---------- F. DTC ----------
cu=rd('DTC_CUMULATIVE.xlsx','CURRENT NOTIFICATION',usecols=['Spectrum_Current_TBU','Episode_ID','Spectrum_Diagnosis_Date'])
cu['dx']=pd.to_datetime(cu.Spectrum_Diagnosis_Date); cu['tal']=cu.Spectrum_Current_TBU.map(nz)
da=rd('DTC_CUMULATIVE.xlsx','DTC',usecols=['PatientId','Assessment Type'])
tm={'baseline':'Baseline','followUp1stMonth':'1st FU','followUp2ndMonth':'2nd FU','followUp3rdMonth':'3rd FU','followUp4thMonth':'4th FU','followUp5thMonth':'5th FU','followUp6thMonth':'6th FU'}
da['t']=da['Assessment Type'].map(tm); doneset={k:set(g.PatientId) for k,g in da.dropna(subset=['t']).groupby('t')}
fus=['Baseline','1st FU','2nd FU','3rd FU','4th FU','5th FU','6th FU']; off_={'Baseline':0,'1st FU':56,'2nd FU':84,'3rd FU':112,'4th FU':140,'5th FU':168,'6th FU':196}
DT={}
for mi,mn in enumerate(['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][:ASOF.month]):
    S=pd.Timestamp(2026,mi+1,1); E=S+pd.offsets.MonthEnd(0); DT[mn]={}
    for t in T+['District']:
        row=[]
        for f in fus:
            if f=='Baseline': lo,hi=S,min(E,REG)
            else: lo=S-pd.Timedelta(days=off_[f]); hi=lo+pd.Timedelta(days=28)
            e=cu[(cu.dx>=lo)&(cu.dx<=hi)]
            if t!='District': e=e[e.tal==t]
            row.append([int(e.Episode_ID.isin(doneset[f]).sum()),len(e)])
        DT[mn][t]=row
R['dtc2']={'fus':fus,'data':DT}
print('F done',flush=True)

# ---------- G. TPT ----------
tp=rd('tb_preventve_therapy_master.xlsx','NOTIFICATION VS TPT',usecols=['Spectrum_Current_TBU','Spectrum_Modified_Diagnosing_facility_type','Episode_ID','Spectrum_Diagnosis_Date','Site_of_TBDisease','TPT.Person Name','TPT.Spectrum_Treatment_initiation_Date'])
s=tp.Site_of_TBDisease.astype(str).str.upper(); tp=tp[s.str.contains('PULMONARY',na=False)&~s.str.contains('EXTRA',na=False)].copy()
tp['dx']=pd.to_datetime(tp.Spectrum_Diagnosis_Date); tp=tp[tp.dx.dt.year==2026]; tp['m']=tp.dx.dt.month; tp['tal']=tp.Spectrum_Current_TBU.map(nz)
tp['has']=tp['TPT.Person Name'].notna(); tp['given']=tp['TPT.Spectrum_Treatment_initiation_Date'].notna()&tp.has
tp['prv']=tp.Spectrum_Modified_Diagnosing_facility_type.astype(str).str.upper().str.contains('PRIVATE')|tp.Spectrum_Modified_Diagnosing_facility_type.isna()
def tpt(g):
    pts=g.Episode_ID.nunique(); zc=int((g.groupby('Episode_ID').has.any()==False).sum()); con=int(g.has.sum()); gv=int(g.given.sum()); return [pts,zc,con,gv]
TP={t:tpt(tp[tp.tal==t]) for t in T}; TP['Mahisagar Total']=tpt(tp); TP['All - Public']=tpt(tp[~tp.prv]); TP['All - Private']=tpt(tp[tp.prv])
TP['monthly']=[tpt(tp[tp.m==m]) for m in range(1,ASOF.month+1)]; TP['tal_month']={t:[tpt(tp[(tp.tal==t)&(tp.m==m)]) for m in range(1,ASOF.month+1)] for t in T}
R['tpt']=TP
print('G done',flush=True)

# ---------- H. comorbidity / beneficiary / ABHA ----------
c=rd('CO_MORBIDITY_REGISTER_FOR_PENIDNG_HIV_AND_DIABETS_.xlsx','co morbidity register'); c['tal']=c.Spectrum_Current_TBU.map(nz)
hivp=c.HIVTestStatus.isna()|(c.HIVTestStatus=='Unknown'); dmp=c.DiabetesStatus.isna()|(c.DiabetesStatus=='Unknown'); dm=c[c.DiabetesStatus=='Diabetic']; tb=c[c.CurrentTobaccoUser=='Positive']
R['co']={'n':len(c),'hiv_pend':int(hivp.sum()),'dm_pend':int(dmp.sum()),'hiv_pos':int(c.HIVTestStatus.isin(['Positive','Reactive']).sum()),'dm':len(dm),'dm_init':int(dm.DateOf_AntiDiabeticInitiation.notna().sum()),
 'tob':len(tb),'tob_link':int((tb.LinkedForCessation=='Yes').sum()),'alc':int((c.HO_AlcoholIntake=='Yes').sum()),'alc_link':int(((c.HO_AlcoholIntake=='Yes')&(c.LinkedForDeaddiction=='Yes')).sum()),
 'tbu':{t:[int((c.tal==t).sum()),int(hivp[c.tal==t].sum()),int(dmp[c.tal==t].sum())] for t in T},
 'dm_tbu':{t:[int(((c.tal==t)&(c.DiabetesStatus=='Diabetic')).sum()),int(((c.tal==t)&(c.DiabetesStatus=='Diabetic')&c.DateOf_AntiDiabeticInitiation.notna()).sum())] for t in T}}
b=rd('BENIFICIARY_REGISTER.xlsx','BENIFICIARY REGISTER'); b['tal']=b.CurrentFacilityTBU.map(nz); bb=b[b.ForegoBenefits!='Yes']
R['ben']={t:bb[bb.tal==t].Beneficiary_Status.fillna('EMPTY').value_counts().to_dict() for t in T}
R['ben_pay']={'paid':float(b.TotalAmountPaid_Nikshay.sum()),'elig':float(b.TotalBenefitAmountEligible.sum())}
R['ben_reject_top']=b.RejectionReason.value_counts().head(5).to_dict()
pm=rd('PMTBMBA_REGISTER_FOR_ABHASEEDING_AND_CONSENT_FOR_KIT.xlsx','PMTBMBA REGISTER'); pm['tal']=pm.Spectrum_Current_TBU.map(nz)
R['abha']={t:[int(((pm.tal==t)&(pm['ABHA generated/seeded']=='Yes')).sum()),int((pm.tal==t).sum())] for t in T}
R['consent']={t:[int(((pm.tal==t)&(pm.PatientConsent=='Consented')).sum()),int((pm.tal==t).sum())] for t in T}
R['kits']={'received':int(pm.Total_Kits_Received_By_Patient.notna().sum()),'n':len(pm)}
json.dump(R,open(OUT,'w'),default=str,indent=1); print('ALL DONE')
