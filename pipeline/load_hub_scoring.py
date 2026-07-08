"""load_hub_scoring.py — load definitions + raw metrics (clipped to store lifecycle),
recompute Points + GlazeGrade. Usage: python load_hub_scoring.py <hub.xlsx> <db> <defs.json> <lifecycle.json>"""
import sys, json, sqlite3, importlib.util, openpyxl
from collections import defaultdict
HUB,DB,CFG = sys.argv[1],sys.argv[2],sys.argv[3]
LIFE = sys.argv[4] if len(sys.argv)>4 else None
spec=importlib.util.spec_from_file_location('scoring','/tmp/scoring.py'); sc=importlib.util.module_from_spec(spec); spec.loader.exec_module(sc)
defs=sc.load_defs(CFG); byname={m['name']:m for m in defs['metrics']}
life={int(k):v for k,v in json.load(open(LIFE)).items()} if LIFE else {}
EXCLUDE_PER=set(defs.get('exclude_periods',[]))
def valid(pc,per):
    L=life.get(pc)
    if not L or not L.get('open'): return True
    if L.get('status') not in ('Open', None): return False   # exclude Sold/Closed entirely
    if per < L['open']: return False
    if L.get('close') and per > L['close']: return False
    return True

wb=openpyxl.load_workbook(HUB,data_only=True)
fmt={r[0]:r[6] for r in list(wb['Store Master'].iter_rows(values_only=True))[4:] if r[0] is not None}
con=sqlite3.connect(DB); cur=con.cursor()
cur.executescript("""
DROP TABLE IF EXISTS metric_definition; DROP TABLE IF EXISTS ops_metric_raw; DROP TABLE IF EXISTS glazegrade_period;
DROP TABLE IF EXISTS store_lifecycle;
CREATE TABLE metric_definition(name TEXT PRIMARY KEY,direction TEXT,cuts TEXT,unit TEXT,
  weight_dt_v1 REAL,weight_in_v1 REAL,weight_dt_v2 REAL,weight_in_v2 REAL,active INTEGER,placeholder INTEGER,note TEXT);
CREATE TABLE ops_metric_raw(pc INTEGER,period TEXT,metric TEXT,value REAL,points INTEGER,PRIMARY KEY(pc,period,metric));
CREATE TABLE glazegrade_period(pc INTEGER,period TEXT,gg REAL,n_metrics INTEGER,PRIMARY KEY(pc,period));
CREATE TABLE store_lifecycle(pc INTEGER PRIMARY KEY,status TEXT,open_month TEXT,close_month TEXT);
""")
for m in defs['metrics']:
    cur.execute("INSERT INTO metric_definition VALUES(?,?,?,?,?,?,?,?,?,?,?)",
        (m['name'],m['direction'],json.dumps(m['cuts']),m['unit'],m['weight_dt_v1'],m['weight_in_v1'],
         m['weight_dt_v2'],m['weight_in_v2'],int(m['active']),int(m.get('placeholder',False)),m.get('note')))
for pc,L in life.items():
    cur.execute("INSERT INTO store_lifecycle VALUES(?,?,?,?)",(pc,L.get('status'),L.get('open'),L.get('close')))

pts=defaultdict(dict)
for r in list(wb['Metrics Hub'].iter_rows(values_only=True))[4:]:
    if r[0] is None: continue
    pc,period,metric,val=r[0],r[3],r[4],r[5]
    if metric not in byname or not byname[metric]['active']: continue
    if period in EXCLUDE_PER: continue
    if not valid(pc,period): continue
    p=sc.score_points(byname[metric],val)
    cur.execute("INSERT OR REPLACE INTO ops_metric_raw VALUES(?,?,?,?,?)",(pc,period,metric,val,p))
    if p is not None: pts[(pc,period)][metric]=p
for (pc,period),mp in pts.items():
    cur.execute("INSERT OR REPLACE INTO glazegrade_period VALUES(?,?,?,?)",(pc,period,sc.glazegrade(mp,fmt.get(pc),defs),len(mp)))
con.commit()
n=lambda q:cur.execute(q).fetchone()[0]
print('metric_definition:',n("SELECT COUNT(*) FROM metric_definition"),'| ops_metric_raw:',n("SELECT COUNT(*) FROM ops_metric_raw"),
      '| glazegrade_period:',n("SELECT COUNT(*) FROM glazegrade_period"),'| lifecycle:',n("SELECT COUNT(*) FROM store_lifecycle"))
con.close()
