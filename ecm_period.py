#!/usr/bin/env python3
"""
ecm_period.py — ECM Workforce System, canonical classifier.

Turns a monthly employee master (.xlsx) into one ECM "period record" that matches
exactly what the dashboard's in-browser classifier produces. Use it to validate a
file before upload, to back-fill history in bulk, or to push straight to the
Apps Script backend without opening the browser.

It honours your own classification, in the same order the dashboard does:
  1. SUMMARY sheet  — rows like "Q1-Labour Operations | 1344" -> totals used directly.
  2. ECM COLUMN     — a column whose values are Q1..Q4 (by header name OR by content)
                      -> your split is kept; cost summed from salary if present.
  3. KEYWORD        — last resort: derive Q1..Q4 from Type / Designation / Department.

Usage:
    python ecm_period.py master.xlsx --label "Jul 26" --cadence monthly
    python ecm_period.py master.xlsx --label "Jul 26" --cadence monthly \
        --post https://script.google.com/macros/s/XXXX/exec --token CHANGE-ME-TOKEN
"""
import argparse, json, re, sys

# --- keyword fallback lists (footwear/RMG tuned; edit to fit your org) ---
SUPPORT_DEPTS = {'account & finiance', 'accountant', 'administration',
                 'human resources management', 'merchandising', 'procurement', 'sales'}
LAB_SUPPORT = ['loader', 'cleaner', 'sweeper', 'seweep', 'security', 'fire',
               'driver', 'delivery', 'vehicle', 'fork lift', 'plumber', 'helper',
               'store keeper', 'encoder', 'medical', 'garden']

# --- flexible column detection (header aliases, normalised) ---
COLMAP = {
    'type':  ['type', 'emptype', 'employeetype', 'empcategory', 'category', 'staffcategory',
              'workforcetype', 'class', 'workertype', 'employeecategory'],
    'desig': ['designame', 'designation', 'desig', 'jobtitle', 'title', 'position', 'role', 'designationname'],
    'dept':  ['deptname', 'department', 'dept', 'departmentname', 'division', 'section', 'functionalarea'],
    'gross': ['ogross', 'gross', 'grosssalary', 'salary', 'grosspay', 'totalsalary', 'grossamount',
              'monthlysalary', 'grosswage', 'grosssal', 'ctc', 'netpay', 'basicgross'],
    'quad':  ['ecm', 'quadrant', 'classification', 'segment', 'ecmsegment', 'ecmquadrant',
              'quad', 'ecmclass', 'ecmcategory', 'ecmtype'],
}


def norm(s):
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9 ]', ' ', str(s).lower())).strip()


def hkey(h):
    return re.sub(r'[^a-z0-9]', '', str('' if h is None else h).lower())


def to_num(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r'[, ]', '', str('' if v is None else v))
    return float(s) if re.fullmatch(r'-?\d+(\.\d+)?', s) else None


def parse_quad(v):
    """Return 'Q1'..'Q4' from a label like 'Q2-Labour Operations Support', else None."""
    if v is None:
        return None
    s = str(v).lower()
    m = re.search(r'q\s*([1-4])(?![0-9])', s)
    if m:
        return 'Q' + m.group(1)
    m = re.match(r'\s*([1-4])\s*[-:.\s]', s)
    if m and re.search(r'labour|labor|worker|staff|operation|support', s):
        return 'Q' + m.group(1)
    lab = bool(re.search(r'labour|labor|worker', s))
    stf, sup, ops = 'staff' in s, 'support' in s, 'operation' in s
    if stf and sup: return 'Q4'
    if stf and ops: return 'Q3'
    if lab and sup: return 'Q2'
    if lab and ops: return 'Q1'
    return None


def resolve_headers(cols):
    m = {}
    for c in cols:
        k = hkey(c)
        for canon, aliases in COLMAP.items():
            if k in aliases and canon not in m:
                m[canon] = c
    return m


def find_quad_column(df):
    sample = df.head(400)
    best, best_hits = None, 0
    for c in df.columns:
        hits = sum(1 for v in sample[c] if parse_quad(v))
        if hits > best_hits:
            best_hits, best = hits, c
    return best if best_hits >= len(sample) * 0.6 else None


def classify(typ, dept, desig):
    """Keyword fallback: Q1 Labour-Ops, Q2 Labour-Support, Q3 Staff-Ops, Q4 Staff-Support."""
    is_staff = norm(typ) == 'staff'
    d, g = norm(dept), norm(desig)
    if is_staff:
        func = 'Support' if (d in SUPPORT_DEPTS or d in ('ware house', 'transport')) \
            else ('Operations' if d in ('production', 'maintenance', 'quality assurance') else 'Support')
    else:
        if any(k in g for k in LAB_SUPPORT):
            func = 'Support'
        elif d in ('production', 'quality assurance'):
            func = 'Operations'
        else:
            func = 'Support'
    return {('Staff', 'Operations'): 'Q3', ('Staff', 'Support'): 'Q4',
            ('Worker', 'Operations'): 'Q1', ('Worker', 'Support'): 'Q2'}[
        ('Staff' if is_staff else 'Worker', func)]


def detect_summary(sheets):
    """If a sheet is an aggregate (a handful of Q1..Q4 rows with counts), return head dict."""
    for name, raw in sheets.items():
        df = raw.copy()
        df.columns = range(df.shape[1])          # treat as raw grid
        head, seen, rows_with_quad = {}, set(), 0
        for _, row in df.iterrows():
            q = next((parse_quad(c) for c in row if parse_quad(c)), None)
            if not q:
                continue
            rows_with_quad += 1
            nums = [to_num(c) for c in row if to_num(c) is not None and to_num(c) < 1e7]
            if nums and q not in seen:
                head[q] = int(round(max(nums)))
                seen.add(q)
        if len(seen) >= 3 and rows_with_quad <= 12:
            for q in ('Q1', 'Q2', 'Q3', 'Q4'):
                head.setdefault(q, 0)
            return head
    return None


def period_record(head, label, cost, M):
    tot = sum(head.values())
    pay = sum(cost.values())
    staff_layer = head['Q3'] + head['Q4']
    return {
        'label': label,
        'total': tot,
        'head': head,
        'cost': {q: round(cost[q]) for q in cost},
        'payroll': round(pay) if pay else (0 if any(cost.values()) else None),
        'tiers': {'W': head['Q1'] + head['Q2'], 'S': max(0, staff_layer - M), 'M': M, 'total': tot},
        'opsSupport': round((head['Q1'] + head['Q3']) / max(1, head['Q2'] + head['Q4']), 2),
        'q3pct': round(head['Q3'] / max(1, tot) * 100, 1),
        'q4pct': round(head['Q4'] / max(1, tot) * 100, 1),
        'directIndirect': round(head['Q1'] / max(1, head['Q2']), 1),
        'recon': False,
    }


def build_period(path, label):
    import pandas as pd
    sheets = pd.read_excel(path, sheet_name=None, header=None)

    # 1) summary?
    summ = detect_summary(sheets)
    if summ:
        cost = {q: 0 for q in summ}
        rec = period_record(summ, label, cost, 0)
        rec['payroll'] = None
        rec['_mode'] = 'summary'
        return rec

    # re-read with headers for roster handling; pick a usable sheet
    framed = pd.read_excel(path, sheet_name=None)
    picked = None
    for name, df in framed.items():
        if df.empty:
            continue
        m = resolve_headers(df.columns)
        quad = m.get('quad') or find_quad_column(df)
        if quad or all(k in m for k in ('type', 'desig', 'dept')):
            picked = (df, m, quad); break
        picked = picked or (df, m, None)
    if picked is None:
        sys.exit('ERROR: no data rows found in the workbook.')
    df, m, quad = picked

    gross = m.get('gross')
    if quad:                                       # 2) explicit classification column
        df = df[df[quad].map(lambda v: parse_quad(v) is not None)]
        if df.empty:
            sys.exit(f'ERROR: classification column "{quad}" had no Q1-Q4 values.')
        df = df.assign(ECM=df[quad].map(parse_quad))
        mode = f'ecm-column:{quad}'
    else:                                          # 3) keyword fallback
        for k in ('type', 'desig', 'dept'):
            if k not in m:
                sys.exit('ERROR: no classification column found and the columns needed to derive one '
                         f'are missing ({k}). Add an ECM column (Q1-Q4) or Type/Designation/Department.')
        df = df.assign(ECM=df.apply(lambda r: classify(r[m['type']], r[m['dept']], r[m['desig']]), axis=1))
        mode = 'keyword'

    head = {q: int((df['ECM'] == q).sum()) for q in ('Q1', 'Q2', 'Q3', 'Q4')}
    if gross:
        g = df[gross].fillna(0).map(lambda v: to_num(v) or 0)
        cost = {q: float(g[df['ECM'] == q].sum()) for q in head}
    else:
        cost = {q: 0 for q in head}
    desig_col = m.get('desig')
    M = int(df[desig_col].map(lambda v: 'manager' in norm(v)).sum()) if desig_col else 0

    rec = period_record(head, label, cost, M)
    if not gross:
        rec['payroll'] = None
    rec['_mode'] = mode
    return rec


def main():
    ap = argparse.ArgumentParser(description='Build an ECM period record from a master xlsx.')
    ap.add_argument('xlsx', help='path to the monthly employee master (.xlsx)')
    ap.add_argument('--label', required=True, help='period label, e.g. "Jul 26"')
    ap.add_argument('--cadence', default='monthly', choices=['monthly', 'quarterly', 'yearly'])
    ap.add_argument('--post', help='Apps Script /exec URL to push the record to')
    ap.add_argument('--token', help='upload token (must match the backend TOKEN)')
    a = ap.parse_args()

    period = build_period(a.xlsx, a.label)
    mode = period.pop('_mode', 'keyword')
    print(f'[classification mode] {mode}', file=sys.stderr)
    print(json.dumps(period, indent=2, ensure_ascii=False))

    if period['q3pct'] < 10.0:
        print(f'\n[!] Q3 share {period["q3pct"]}% is below the 10% healthy floor — '
              'supervision/QC cover is thin for this period.', file=sys.stderr)

    if a.post:
        import urllib.request
        body = json.dumps({'token': a.token or '', 'cadence': a.cadence, 'period': period}).encode()
        req = urllib.request.Request(a.post, data=body,
                                     headers={'Content-Type': 'text/plain;charset=utf-8'})
        with urllib.request.urlopen(req) as r:
            print('\nbackend:', r.read().decode())


if __name__ == '__main__':
    main()
