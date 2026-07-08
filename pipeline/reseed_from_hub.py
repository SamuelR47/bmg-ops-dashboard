import openpyxl, json, importlib.util, sys
from collections import defaultdict
spec=importlib.util.spec_from_file_location('scoring','/tmp/scoring.py'); sc=importlib.util.module_from_spec(spec); spec.loader.exec_module(sc)
HUB=sys.argv[1] if len(sys.argv)>1 else '/tmp/hub.xlsx'
LIFE=sys.argv[2] if len(sys.argv)>2 else '/tmp/store_lifecycle.json'
defs=sc.load_defs('/tmp/metric_definitions.json'); byname={m['name']:m for m in defs['metrics']}
ACTIVE=[m['name'] for m in defs['metrics'] if m['active']]
EXCLUDE_PER=set(defs.get('exclude_periods',[]))
life={int(k):v for k,v in json.load(open(LIFE)).items()}
import os as _os
ll_path='/tmp/latlon.json'
LATLON={int(k):v for k,v in (json.load(open(ll_path)).items() if _os.path.exists(ll_path) else [])}

def valid(pc,per):
    L=life.get(pc)
    if not L or not L.get('open'): return True
    if L.get('status') not in ('Open', None): return False
    if per < L['open']: return False
    if L.get('close') and per > L['close']: return False
    return True

wb=openpyxl.load_workbook(HUB,data_only=True)
def rowsof(name,hdr=3):
    r=list(wb[name].iter_rows(values_only=True)); h={c:i for i,c in enumerate(r[hdr])}; return h,[x for x in r[hdr+1:] if x[0] is not None]

h,sm=rowsof('Store Master')
master={}
for r in sm:
    if str(r[h['Active']]).upper() in ('TRUE','Y','YES','1'):
        master[r[h['PC']]]={'pc':r[h['PC']],'name':r[h['Store Name']],'market':r[h['Geography']],
            'dm':r[h['DM']],'rd':r[h['RD']],'format':r[h['Format']]}

# DM override from Store Contact List (DM name only)
import os
_ov='/tmp/dm_override.json'
if os.path.exists(_ov):
    _dm=json.load(open(_ov)); _dm={int(k):v for k,v in _dm.items()}
    for _pc,_m in master.items():
        if _pc in _dm: _m['dm']=_dm[_pc]

h,asss=rowsof('Adjusted SSS')
adj={}
for r in asss:
    pc=r[h['PC']]; obs=r[h['Observed SSS']]; baseline=r[h['Adjusted SSS']]
    adj[pc]=dict(actual_sss=obs,baseline_sss=baseline,
        net_drag=(baseline-obs) if (baseline is not None and obs is not None) else 0,
        new_gg=r[h['New GG']],old_gg=r[h['Old GG']],ext_count=r[h['# Active Events']] or 0,
        weakest=r[h['Weakest Metrics']] or '',current_gm=r[h['Current GM']],
        gm_changes=r[h['GM Changes']] or 0,category=r[h['Category']])

h,gms=rowsof('GM Summary')
gm={r[h['PC']]:dict(current_gm=r[h['Current GM']],gm_changes=r[h['GM Changes Since 2024']] or 0,
    gm_tenure=r[h['Current Tenure (Yrs)']] or 0) for r in gms}

h,ev=rowsof('Event Impacts L3M')
events=defaultdict(list)
for r in ev:
    events[r[h['PC']]].append(dict(new_store=r[h['Competitor/New Store']],event_type=r[h['Type']],
        date_opened=str(r[h['Date Opened']]),impact=r[h['Impact']],net=r[h['Net SSS Impact']],src=r[h['Source']]))

h,gh=rowsof('GM History')
hist=defaultdict(list)
for r in gh:
    pc=r[h.get('PC',0)]; per=r[h.get('Period')] if 'Period' in h else None
    nm=r[h.get('GM Name',h.get('GM'))] if ('GM Name' in h or 'GM' in h) else None
    if pc and per and nm and valid(pc,per): hist[pc].append((per,nm))
def collapse(seq):
    seq=sorted(seq); out=[]
    for per,nm in seq:
        if out and out[-1]['gm']==nm: out[-1]['end']=per
        else: out.append({'gm':nm,'start':per,'end':per})
    return out

# raw metrics -> points, CLIPPED to lifecycle window
h,mh=rowsof('Metrics Hub')
ptsP=defaultdict(dict); rawP=defaultdict(dict)
for r in mh:
    pc,per,metric,val=r[h['PC']],r[h['Period']],r[h['Metric']],r[h['Value']]
    if metric not in byname or not byname[metric]['active']: continue
    if per in EXCLUDE_PER: continue
    if not valid(pc,per): continue
    p=sc.score_points(byname[metric],val)
    if p is not None: ptsP[(pc,per)][metric]=p
    rawP[(pc,per)][metric]=val

h,tv=rowsof('Trend View')
base_per={(r[h['PC']],r[h['Period']]):r[h['Baseline SSS']] for r in tv}
all_periods=sorted({per for (pc,per) in ptsP})

v12={s['pc']:s for s in json.load(open('/tmp/v12_data.json'))['stores']}
def v12trend(pc): return {t['period']:t for t in v12.get(pc,{}).get('trend',[])}

L3M=['2026-02','2026-03','2026-04']
stores=[]
for pc,m in master.items():
    a=adj.get(pc,{}); g=gm.get(pc,{}); L=life.get(pc,{})
    metrics={}; mvals={}
    for met in ACTIVE:
        ps=[ptsP[(pc,p)][met] for p in L3M if met in ptsP.get((pc,p),{})]
        rs=[rawP[(pc,p)][met] for p in L3M if (pc,p) in rawP and rawP[(pc,p)].get(met) is not None]
        metrics[met]=round(sum(ps)/len(ps),2) if ps else None
        mvals[met]=round(sum(rs)/len(rs),4) if rs else None
    vt=v12trend(pc); trend=[]
    for per in all_periods:
        if not valid(pc,per): continue
        ggv=sc.glazegrade(ptsP.get((pc,per),{}), m['format'], defs)
        if ggv is None and per not in vt: continue
        old=vt.get(per,{})
        trend.append({'period':per,'gg':ggv,'sss':old.get('sss'),'drag':old.get('drag'),
            'pf_sss':base_per.get((pc,per),old.get('pf_sss'))})
    # RECOMPUTED_HEADLINE: derive new/old GlazeGrade + weakest from pipeline (config-driven), not Hub
    _tg={t['period']:t['gg'] for t in trend if t['gg'] is not None}
    def _avg(ps):
        vs=[_tg[p] for p in ps if p in _tg]
        return round(sum(vs)/len(vs),2) if vs else None
    L3M_W=['2026-02','2026-03','2026-04']; PRIOR_W=['2025-11','2025-12','2026-01']
    _new_gg=_avg(L3M_W); _old_gg=_avg(PRIOR_W)
    _pm=[(met,metrics[met]) for met in metrics if metrics[met] is not None]
    _weak=', '.join('%s(%.1f)'%(met,val) for met,val in sorted(_pm,key=lambda x:x[1])[:2]) if _pm else (a.get('weakest','') or '')
    sv=v12.get(pc,{})
    stores.append({'pc':pc,'name':m['name'],'market':m['market'],'dm':m['dm'],'rd':m['rd'],'format':m['format'],
        'actual_sss':a.get('actual_sss',0) or 0,'baseline_sss':a.get('baseline_sss',0) or 0,'net_drag':a.get('net_drag',0) or 0,
        'new_gg':(_new_gg if _new_gg is not None else a.get('new_gg')),'old_gg':(_old_gg if _old_gg is not None else a.get('old_gg')),
        'sales_4w':sv.get('sales_4w',0),'annual_sales':sv.get('annual_sales',0),
        'n_events':len(events.get(pc,[])),'ext_count':a.get('ext_count',0) or 0,
        'weakest':_weak,'current_gm':g.get('current_gm') or a.get('current_gm'),
        'gm_changes':g.get('gm_changes',a.get('gm_changes',0)),'gm_tenure':g.get('gm_tenure',0),
        'metrics':metrics,'metric_values':mvals,'trend':trend,
        'events':events.get(pc,[]),'gm_timeline':collapse(hist.get(pc,[])),'category':a.get('category'),'lat':(LATLON.get(pc) or [None,None])[0],'lon':(LATLON.get(pc) or [None,None])[1],
        'open_month':L.get('open'),'status':L.get('status')})
json.dump({'stores':stores},open('/tmp/hub_data.json','w'))
# report
pre=sum(1 for s in stores for t in s['trend'] if s.get('open_month') and t['period']<s['open_month'])
print('stores:',len(stores),'| total trend rows:',sum(len(s['trend']) for s in stores),'| pre-open grades remaining:',pre)
