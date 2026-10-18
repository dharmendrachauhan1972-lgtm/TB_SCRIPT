#!/usr/bin/env python3
"""Mahisagar monthly review extractor.

Usage: python3 extract_monthly.py <uploads_dir> <out.json> [--month YYYY-MM]

Computes, for ONE review month M, each indicator for three periods:
  M (last month), M-1 (previous month), and M one year earlier,
by taluka and district, straight from the raw Nikshay/Spectrum registers
(so no pre-built analysis workbook is needed). Anything it cannot compute is
listed in "not_computed" - the deck must show "n/a", never an invented number.
Default month = the calendar month before today.
"""
import sys, os, glob, json, argparse, datetime as dt
import pandas as pd

T = ['Balasinor', 'Kadana', 'Khanpur', 'Lunawada', 'Santrampur', 'Virpur']
# Monthly CBNAAT/presumptive target by taluka (user-stated; Lunavada == Lunawada)
CB_MONTHLY = {'Lunawada': 763.385, 'Khanpur': 292.295, 'Kadana': 424.14,
              'Santrampur': 889.5875, 'Balasinor': 402.65, 'Virpur': 260.0}
# Notification ANNUAL target; monthly = annual / 12
NOTIF_ANNUAL = {'Lunawada': 1243, 'Santrampur': 769, 'Balasinor': 440,
                'Virpur': 202, 'Kadana': 191, 'Khanpur': 168}

ap = argparse.ArgumentParser()
ap.add_argument('uploads', nargs='?', default='/mnt/user-data/uploads')
ap.add_argument('out', nargs='?', default='/tmp/monthly.json')
ap.add_argument('--month', help='review month YYYY-MM (default: previous calendar month)')
a = ap.parse_args()

today = dt.date.today()
M = pd.Period(a.month, 'M') if a.month else (pd.Period(today, 'M') - 1)
PER = {'last': M, 'prev': M - 1, 'lastyear': M - 12}
R = {'month': str(M), 'periods': {k: str(v) for k, v in PER.items()},
     'targets': {'cb_monthly': CB_MONTHLY, 'notif_annual': NOTIF_ANNUAL},
     'warnings': [], 'not_computed': []}


def find(pattern):
    hits = sorted(set(glob.glob(os.path.join(a.uploads, pattern)) +
                      glob.glob(os.path.join(a.uploads, pattern.lower())) +
                      glob.glob(os.path.join(a.uploads, pattern.upper()))))
    hits = [h for h in hits if not os.path.basename(h).startswith('~$')]
    return hits[0] if hits else None


def load(pattern, sheet):
    f = find(pattern)
    if not f:
        raise FileNotFoundError(f'no file matching {pattern}')
    return pd.read_excel(f, sheet_name=sheet)


def tal(df, col):
    """Kadana (non-functional) TBU is reported under Lunawada by the source; keep names as in data."""
    return df[col].astype(str).str.strip().str.title()


def pct(n, d):
    return None if not d else round(100.0 * n / d, 1)


def guard(name, fn):
    try:
        fn()
    except Exception as e:  # keep going; report what was skipped
        R['not_computed'].append({'module': name, 'reason': f'{type(e).__name__}: {e}'})


def check_complete(name, dates):
    last = dates.dropna().max()
    if pd.isna(last):
        R['warnings'].append(f'{name}: no valid dates found'); return
    end = M.end_time
    if last < end - pd.Timedelta(days=3):
        R['warnings'].append(f'{name}: latest record is {last.date()}, before the end of {M}; '
                             f'the review month may be partial.')


# ---------- 1. Presumptive: tested, NAAT, microscopy, X-ray ----------
def presumptive():
    p = load('*PRESUMPTIVE_REGISTER*', 'PRESUMPTIVE REGISTER')
    p['ed'] = pd.to_datetime(p['Spectrum_Enrollment_Date'], errors='coerce')
    check_complete('Presumptive register', p['ed'])
    p['per'] = p['ed'].dt.to_period('M'); p['tbu'] = tal(p, 'Spectrum_Enrolment_TBU')
    tested = p['Date_NAAT'].notna() | p['Date_Microscopy'].notna()
    abn = p['Result_XRay_Positive'].eq('Yes')
    asym = p['Symptom'].fillna('').astype(str).str.replace(',', '').str.strip().eq('Asymptomatic')
    out = {}
    for label, per in PER.items():
        out[label] = {}
        for t in T + ['District']:
            m = (p['per'] == per) & (True if t == 'District' else p['tbu'] == t)
            enr = int(m.sum())
            row = {'enrolled': enr, 'tested': int((m & tested).sum()),
                   'naat': int((m & p['Date_NAAT'].notna()).sum()),
                   'micro': int((m & p['Date_Microscopy'].notna()).sum()),
                   'xray_offered': int((m & p['Xray_Offered'].eq('Yes')).sum()),
                   'abnormal_xray': int((m & abn).sum()),
                   'abnormal_tested': int((m & abn & tested).sum()),
                   'asym_tested': int((m & asym & tested).sum())}
            tgt = sum(CB_MONTHLY.values()) if t == 'District' else CB_MONTHLY.get(t)
            row['target'] = round(tgt, 1) if tgt else None
            row['tested_pct_target'] = pct(row['tested'], tgt)
            row['naat_pct_target'] = pct(row['naat'], tgt)
            row['xray_offered_pct'] = pct(row['xray_offered'], enr)
            row['tested_pct_enrolled'] = pct(row['tested'], enr)
            row['abnormal_tested_pct'] = pct(row['abnormal_tested'], row['abnormal_xray'])
            out[label][t] = row
    R['presumptive'] = out


# ---------- 2. Notification: diagnosing-facility base and enrolment-facility base ----------
def notification():
    res = {}
    for kind, pat in [('diag', '*Notification_Register_Diagnosing*'), ('enrol', '*Notification_Register_Enrolment*')]:
        d = load(pat, 'Data')
        d['dx'] = pd.to_datetime(d['Spectrum_Diagnosis_Date'], errors='coerce')
        check_complete(f'Notification ({kind} base)', d['dx'])
        d['per'] = d['dx'].dt.to_period('M'); d['tal'] = tal(d, 'Taluka')
        res[kind] = {}
        for label, per in PER.items():
            res[kind][label] = {}
            for t in T + ['District']:
                m = (d['per'] == per) & (True if t == 'District' else d['tal'] == t)
                n = int(m.sum())
                tgt = (sum(NOTIF_ANNUAL.values()) if t == 'District' else NOTIF_ANNUAL[t]) / 12.0
                res[kind][label][t] = {'notified': n, 'public': int((m & d['Type'].eq('Public')).sum()),
                                       'private': int((m & d['Type'].eq('Private')).sum()),
                                       'monthly_target': round(tgt, 1), 'pct_target': pct(n, tgt)}
    R['notification'] = res


# ---------- 3. Current notification register: treatment start + matured-cohort outcome ----------
def current_register():
    d = load('*CURRENT_NOTIFICATION_REGISTER*', 'CURRENT NOTIFICATION REGISTER')
    d['dx'] = pd.to_datetime(d['Spectrum_Diagnosis_Date'], errors='coerce')
    d['ti'] = pd.to_datetime(d['Treatment_Initiation_Date'], errors='coerce')
    d['per'] = d['dx'].dt.to_period('M'); d['tbu'] = tal(d, 'Spectrum_Current_TBU')
    check_complete('Current notification register', d['dx'])
    d['gap'] = (d['ti'] - d['dx']).dt.days
    asof = d['dx'].max()
    start = {}
    for label, per in PER.items():
        start[label] = {}
        for t in T + ['District']:
            m = (d['per'] == per) & (True if t == 'District' else d['tbu'] == t)
            n = int(m.sum()); started = m & d['ti'].notna()
            not_init = m & d['ti'].isna() & d['Treatment_Outcome'].isna()
            if asof is not pd.NaT:  # only call it 'not initiated' once >7 days have passed
                not_init = not_init & (d['dx'] < asof - pd.Timedelta(days=7))
            start[label][t] = {'diagnosed': n, 'started': int(started.sum()),
                               'late_gt7': int((started & (d['gap'] > 7)).sum()),
                               'late_pct_of_started': pct(int((started & (d['gap'] > 7)).sum()), int(started.sum())),
                               'not_initiated': int(not_init.sum()),
                               'median_gap_days': (None if not started.any() else float(d.loc[started & m, 'gap'].median()))}
    R['treatment_start'] = start

    # Outcome: DS-TB, 2HRZE/4HRE, Pulmonary only; only cohorts notified >= 9 months ago.
    # Review month M -> matured cohort month is M-9; compare with the month before it and the year before.
    ds = d[(d['Type_of_TB_regimen'] == '2HRZE/4HRE') & (d['Site_of_TBDisease'] == 'Pulmonary')].copy()

    def cat(o):
        if pd.isna(o): return 'Pending'
        if o in ('CURED', 'TREATMENT_COMPLETE'): return 'Success'
        if o == 'DIED': return 'Died'
        if o == 'TREATMENT_REGIMEN_CHANGED': return 'Regimen changed'
        return 'Other'
    ds['cat'] = ds['Treatment_Outcome'].map(cat)
    cohorts = {'last': M - 9, 'prev': M - 10, 'lastyear': M - 21}
    out = {'cohort_months': {k: str(v) for k, v in cohorts.items()}}
    for label, per in cohorts.items():
        out[label] = {}
        for t in T + ['District']:
            g = ds[(ds['per'] == per) & (True if t == 'District' else ds['tbu'] == t)]
            n = len(g); c = g['cat'].value_counts().to_dict()
            out[label][t] = {'denominator': n, 'success': c.get('Success', 0), 'died': c.get('Died', 0),
                             'regimen_changed': c.get('Regimen changed', 0), 'other': c.get('Other', 0),
                             'pending': c.get('Pending', 0), 'success_pct': pct(c.get('Success', 0), n),
                             'died_pct': pct(c.get('Died', 0), n), 'pending_pct': pct(c.get('Pending', 0), n)}
    R['outcome_matured_cohort'] = out


guard('presumptive', presumptive)
guard('notification', notification)
guard('current_register', current_register)

R['not_computed'] += [
    {'module': 'follow_up_IP_CP_PTFU_DTC_TPT', 'reason': 'read the monthly column from the existing analysis workbooks '
     '(IP/CP, PTFU, DTC, TPT); last-year values only if a 2025 workbook is supplied'},
    {'module': 'comorbidity_and_NPY', 'reason': 'registers are snapshots with no month field: show current status only, no monthly comparison'}]

os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
json.dump(R, open(a.out, 'w'), indent=1, default=str)
print('wrote', a.out, '| review month', M, '| warnings:', len(R['warnings']),
      '| skipped modules:', [x['module'] for x in R['not_computed'] if x['module'] in ('presumptive', 'notification', 'current_register')])
