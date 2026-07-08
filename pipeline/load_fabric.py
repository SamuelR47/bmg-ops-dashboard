"""
load_fabric.py — Land the Fabric exports into operations_trending.db and recompute
GlazeGrade. Replaces the manual Ops Metrics Hub for the 8 automated metrics.

INPUTS (CSV exported from fabric_export.sql, placed in the project folder):
  ops_metrics.csv    pc, period, metric, raw_value     (8 metrics, store x month)
  sales.csv          pc, period, sales, sales_py
  store_master.csv   pc, name, market, dm, rd, format, status, open_date, close_date
  manual_inputs.csv  pc, period, metric, raw_value     (Accuracy, Cert SL/AGM — optional)

WHAT IT DOES
  - Loads the raw feed tables (ops_metric_monthly, sales_monthly) and upserts store_master
    (preserves existing lat/lon).
  - Recomputes metric_score (L3M avg points + raw), store_period.gg (per-period GlazeGrade),
    and store_snapshot.new_gg/old_gg/weakest/sales_4w/annual_sales via scoring.py.
  - PRESERVES sss / drag / pf_sss / events / GM / cannibalization fields (those come from the
    proforma + other feeds, not this pull).
  - Honors metric_definitions.json exclude_periods. Auto-detects L3M = latest 3 non-excluded
    periods present (override with --l3m 2026-02,2026-03,2026-04).

USAGE
  python load_fabric.py [--dir <folder with csvs>] [--db operations_trending.db] [--l3m a,b,c]
  Then refresh the dashboard/xlsx:  python generate_outputs.py   (and build_glazegrade_xlsx.py)
"""
import argparse, csv, os, sqlite3, sys
from collections import defaultdict
import scoring as sc

HERE = os.path.dirname(os.path.abspath(__file__))

def read_csv(path):
    if not os.path.exists(path):
        return None
    with open(path, newline='', encoding='utf-8-sig') as f:
        return list(csv.DictReader(f))

def to_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--dir', default=os.path.dirname(HERE))   # project folder (csvs live here)
    ap.add_argument('--db',  default=os.path.join(HERE, 'operations_trending.db'))
    ap.add_argument('--l3m', default=None, help='comma-separated periods, e.g. 2026-02,2026-03,2026-04')
    args = ap.parse_args()

    defs = sc.load_defs(os.path.join(HERE, 'metric_definitions.json'))
    byname = {m['name']: m for m in defs['metrics']}
    active = [m['name'] for m in defs['metrics'] if m.get('active')]
    exclude = set(defs.get('exclude_periods', []))

    ops = read_csv(os.path.join(args.dir, 'ops_metrics.csv'))
    if ops is None:
        sys.exit('ops_metrics.csv not found in ' + args.dir)
    manual = read_csv(os.path.join(args.dir, 'manual_inputs.csv')) or []
    sales = read_csv(os.path.join(args.dir, 'sales.csv')) or []
    master = read_csv(os.path.join(args.dir, 'store_master.csv')) or []

    con = sqlite3.connect(args.db); cur = con.cursor()

    # ---- 1. RAW FEED TABLES ----
    cur.execute('DELETE FROM ops_metric_monthly')
    rows = [(int(r['pc']), r['period'], r['metric'], to_float(r['raw_value']))
            for r in (ops + manual) if r.get('pc') and r.get('metric')]
    cur.executemany('INSERT OR REPLACE INTO ops_metric_monthly VALUES (?,?,?,?)', rows)

    if sales:
        cur.execute('DELETE FROM sales_monthly')
        cur.executemany('INSERT OR REPLACE INTO sales_monthly VALUES (?,?,?,?)',
            [(int(r['pc']), r['period'], to_float(r['sales']), to_float(r['sales_py'])) for r in sales if r.get('pc')])

    # ---- 2. STORE MASTER (upsert, preserve lat/lon) ----
    for r in master:
        pc = int(r['pc'])
        ll = cur.execute('SELECT lat, lon FROM store_master WHERE pc=?', (pc,)).fetchone() or (None, None)
        cur.execute('''INSERT INTO store_master(pc,name,market,dm,rd,format,lat,lon) VALUES(?,?,?,?,?,?,?,?)
                       ON CONFLICT(pc) DO UPDATE SET name=excluded.name, market=excluded.market,
                       dm=excluded.dm, rd=excluded.rd, format=excluded.format''',
                    (pc, r.get('name'), r.get('market'), r.get('dm'), r.get('rd'), r.get('format'), ll[0], ll[1]))

    fmt = {pc: f for pc, f in cur.execute('SELECT pc, format FROM store_master')}

    # ---- 3. SCORE: points per (pc,period,metric); GlazeGrade per (pc,period) ----
    raw = defaultdict(dict)   # (pc,period) -> {metric: raw}
    for r in (ops + manual):
        if not r.get('pc'):
            continue
        pc, per, met, val = int(r['pc']), r['period'], r['metric'], to_float(r['raw_value'])
        if met in byname and per not in exclude:
            raw[(pc, per)][met] = val

    pts = {}  # (pc,period) -> {metric: points}
    for (pc, per), mv in raw.items():
        pts[(pc, per)] = {m: sc.score_points(byname[m], v) for m, v in mv.items() if sc.score_points(byname[m], v) is not None}

    all_periods = sorted({per for (_, per) in raw})
    if args.l3m:
        L3M = args.l3m.split(',')
    else:
        L3M = all_periods[-3:]
    prior = all_periods[-6:-3] if len(all_periods) >= 6 else []
    print('Periods present:', all_periods)
    print('L3M:', L3M, '| prior:', prior)

    # store_period.gg (preserve sss/drag/pf_sss)
    for (pc, per), pm in pts.items():
        gg = sc.glazegrade(pm, fmt.get(pc), defs)
        ex = cur.execute('SELECT 1 FROM store_period WHERE pc=? AND period=?', (pc, per)).fetchone()
        if ex:
            cur.execute('UPDATE store_period SET gg=? WHERE pc=? AND period=?', (gg, pc, per))
        else:
            cur.execute('INSERT INTO store_period(pc,period,gg,sss,drag,pf_sss) VALUES(?,?,?,?,?,?)',
                        (pc, per, gg, None, None, None))

    # metric_score (L3M avg points + avg raw) and snapshot headline
    pcs = sorted(fmt)
    for pc in pcs:
        weak = []
        for m in active:
            ps = [pts[(pc, p)][m] for p in L3M if m in pts.get((pc, p), {})]
            rs = [raw[(pc, p)][m] for p in L3M if raw.get((pc, p), {}).get(m) is not None]
            score = round(sum(ps)/len(ps), 2) if ps else None
            rawv  = round(sum(rs)/len(rs), 4) if rs else None
            cur.execute('INSERT OR REPLACE INTO metric_score(pc,metric,score,raw_value) VALUES(?,?,?,?)',
                        (pc, m, score, rawv))
            if score is not None:
                weak.append((m, score))
        def avg_gg(window):
            vs = [sc.glazegrade(pts[(pc, p)], fmt.get(pc), defs) for p in window if (pc, p) in pts]
            vs = [v for v in vs if v is not None]
            return round(sum(vs)/len(vs), 2) if vs else None
        new_gg, old_gg = avg_gg(L3M), avg_gg(prior)
        weakest = ', '.join('%s(%.1f)' % (m, s) for m, s in sorted(weak, key=lambda x: x[1])[:2])
        srow = cur.execute('SELECT 1 FROM store_snapshot WHERE pc=?', (pc,)).fetchone()
        if srow:
            cur.execute('UPDATE store_snapshot SET new_gg=?, old_gg=?, weakest=? WHERE pc=?',
                        (new_gg, old_gg, weakest, pc))

    # sales_4w / annual_sales from sales_monthly
    for pc in pcs:
        srows = cur.execute('SELECT period, sales FROM sales_monthly WHERE pc=? ORDER BY period', (pc,)).fetchall()
        if srows:
            sales_4w = srows[-1][1]
            annual = sum(s for _, s in srows[-13:] if s is not None)
            cur.execute('UPDATE store_snapshot SET sales_4w=?, annual_sales=? WHERE pc=?', (sales_4w, annual, pc))

    con.commit()
    # report
    n_ops = cur.execute('SELECT COUNT(*) FROM ops_metric_monthly').fetchone()[0]
    cov = cur.execute('''SELECT period, COUNT(DISTINCT metric) FROM ops_metric_monthly
                         GROUP BY period ORDER BY period''').fetchall()
    print('Loaded ops rows:', n_ops, '| stores:', len(pcs))
    print('Coverage (period: #metrics):', dict(cov))
    print('Done. Next: python generate_outputs.py  &&  python build_glazegrade_xlsx.py')
    con.close()

if __name__ == '__main__':
    main()
