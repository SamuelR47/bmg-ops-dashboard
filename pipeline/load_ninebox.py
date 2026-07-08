"""load_ninebox.py — parse 9-Box report into DB table `ninebox`, deterministically link
each GM/DM to a store and its L3M GlazeGrade. Usage: python load_ninebox.py <xlsx> <db>"""
import sys, sqlite3, openpyxl, re
XL, DB = sys.argv[1], sys.argv[2]
SUFFIX={'jr','sr','ii','iii','iv'}
def norm(n):
    n=str(n).strip()
    if ',' in n:
        last,first=[x.strip() for x in n.split(',',1)]
        return (first.split()[0]+' '+last).lower()
    return n.lower()
def surname(name):
    name=str(name)
    part = name.split(',')[0] if ',' in name else name
    toks=[t for t in re.split(r'\s+', part) if t.lower().strip('.') not in SUFFIX]
    return (toks[-1] if toks else part).lower().strip('.')
def role_of(t):
    t=(t or '').lower()
    if 'district' in t: return 'DM'
    if 'training manager' in t: return 'CTM'
    return 'GM'
def boxnum(p):
    p=str(p or ''); m=re.match(r'\s*(\d)',p)
    if m: return int(m.group(1))
    if 'consistent star' in p.lower(): return 1
    return None
DM_KEY={'pamela barner':'Pam','matthew brown':'Matthew B','brandie byrd':'Brandie','johnny colorado':'Johnny',
 'gena huckaby':'Gena','jamie knox nalvarte':'Jamie','john lancaster':'John L','sandra mccraw':'Sandy',
 'kimberly mcdaniel':'Kim','eusebio nalvarte':'Chevi','andrew rup':'Andrew','gwen wyman':'Gwen'}

wb=openpyxl.load_workbook(XL,data_only=True)
ws=wb['9 Box Chart']; rows=list(ws.iter_rows(values_only=True)); hdr=rows[0]
idx={h:i for i,h in enumerate(hdr) if h}
def g(r,c): i=idx.get(c); return r[i] if i is not None else None

con=sqlite3.connect(DB); con.row_factory=sqlite3.Row; cur=con.cursor()
# stores: current_gm + gg + dm
stores=[dict(pc=r['pc'],name=r['name'],dm=r['dm'],gm=(r['current_gm'] or ''),gg=r['new_gg'])
        for r in cur.execute("SELECT m.pc,m.name,m.dm,s.current_gm,s.new_gg FROM store_master m JOIN store_snapshot s ON s.pc=m.pc").fetchall()]
store_by_gmname={ s['gm'].lower():s for s in stores if s['gm'] and s['gm']!='NO GM' }

# parse ninebox people
people=[]
for r in rows[1:]:
    nm=g(r,'Payroll Name')
    if not nm: continue
    people.append(dict(name=str(nm), nname=norm(nm), surname=surname(nm), role=role_of(g(r,'Job Title Description')),
        title=g(r,'Job Title Description'), box=boxnum(g(r,'9-Box Placement')), placement=str(g(r,'9-Box Placement')),
        potential=g(r,'Potential Rating Name'), performance=g(r,'Performance Rating Name'),
        loss_risk=g(r,'Loss Risk Rating Name'), loss_impact=g(r,'Loss Impact Rating Name'),
        key_talent=g(r,'Key Talent Rating Name'), critical_role=g(r,'Critical Role Rating Name'),
        reports_to=g(r,'Reports To Legal Name'), bu=g(r,'Business Unit Code'), home_cost=g(r,'Home Cost Number Description'),
        placement_date=str(g(r,'Placement Date:')) if g(r,'Placement Date:') else None))

# ---- deterministic GM->store linkage ----
for p in people:
    p['store_pc']=None; p['store_name']=None; p['gg']=None; p['match']=None
    if p['role']=='DM':
        p['dm_key']=DM_KEY.get(p['nname'])
        ss=[s['gg'] for s in stores if s['dm']==p['dm_key'] and s['gg'] is not None]
        if ss: p['gg']=round(sum(ss)/len(ss),2); p['match']='dm-portfolio'
        continue
    p['dm_key']=None
    s=store_by_gmname.get(p['nname'])   # exact normalized name
    if s:
        p['store_pc'],p['store_name'],p['gg'],p['match']=s['pc'],s['name'],s['gg'],'exact-name'

# unmatched store GMs and unmatched ninebox GM/CTM
matched_gm_names={p['nname'] for p in people if p['match']=='exact-name'}
unmatched_stores=[s for s in stores if s['gm'] and s['gm']!='NO GM' and s['gm'].lower() not in matched_gm_names]
unmatched_nb=[p for p in people if p['role'] in ('GM','CTM') and p['match'] is None]
# surname counts (1:1 uniqueness)
from collections import Counter
nb_sn=Counter(p['surname'] for p in unmatched_nb)
st_sn=Counter(surname(s['gm']) for s in unmatched_stores)
recon=[]
for p in unmatched_nb:
    sn=p['surname']
    if nb_sn[sn]==1 and st_sn[sn]==1:
        s=next(s for s in unmatched_stores if surname(s['gm'])==sn)
        p['store_pc'],p['store_name'],p['gg'],p['match']=s['pc'],s['name'],s['gg'],'unique-surname'
        recon.append((p['name'],s['gm'],s['name'],s['gg']))

# ---- write table ----
cur.executescript("""DROP TABLE IF EXISTS ninebox;
CREATE TABLE ninebox(name TEXT,nname TEXT,title TEXT,role TEXT,box INTEGER,placement TEXT,
 potential TEXT,performance TEXT,loss_risk TEXT,loss_impact TEXT,key_talent TEXT,critical_role TEXT,
 reports_to TEXT,business_unit TEXT,home_cost TEXT,dm_key TEXT,placement_date TEXT,
 store_pc INTEGER,store_name TEXT,gg REAL,match TEXT);""")
for p in people:
    cur.execute("INSERT INTO ninebox VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(
        p['name'],p['nname'],p['title'],p['role'],p['box'],p['placement'],p['potential'],p['performance'],
        p['loss_risk'],p['loss_impact'],p['key_talent'],p['critical_role'],p['reports_to'],p['bu'],
        str(p['home_cost']) if p['home_cost'] else None,p['dm_key'],p['placement_date'],
        p['store_pc'],p['store_name'],p['gg'],p['match']))
con.commit()
print('ninebox rows:',len(people))
gmlinked=sum(1 for p in people if p['role'] in('GM','CTM') and p['gg'] is not None)
print('GM/CTM linked to a store score:',gmlinked)
print()
print('=== Reconciled via unique surname (NEW) ===')
for r in recon: print('   9box:%-26s  store GM:%-22s  @ %-18s gg=%s'%(r[0],r[1],r[2],r[3]))
print()
print('=== Store GMs still unlinked (no deterministic 9-box match) ===')
linked_pcs={p['store_pc'] for p in people if p['store_pc'] is not None}
allsn=Counter(p['surname'] for p in people if p['role'] in('GM','CTM'))
for s in stores:
    if s['gm'] and s['gm']!='NO GM' and s['pc'] not in linked_pcs:
        reason='surname absent in 9-box' if allsn[surname(s['gm'])]==0 else 'surname ambiguous (>1 in 9-box)'
        print('   %-22s @ %-20s — %s'%(s['gm'],s['name'],reason))
con.close()
