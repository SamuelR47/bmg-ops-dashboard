"""
scoring.py — GlazeGrade scoring engine (config-driven).
Reads metric_definitions.json. Two functions:
  score_points(metric, value) -> int 1-5     (threshold binning)
  glazegrade(points_by_metric, store_format)  -> weighted score, renormalized over present metrics
Switch the active weight scheme via 'scheme_in_use' in the config (definition_v1 | proposed_v2).
Validated: reproduces Hub v08 Points (24,932/24,932) and Trend View GlazeGrade (2,419/2,419).
"""
import json, os
HERE=os.path.dirname(os.path.abspath(__file__))

def load_defs(path=None):
    return json.load(open(path or os.path.join(HERE,'metric_definitions.json')))

def score_points(metric_def, value):
    if value is None or metric_def.get('cuts') is None: return None
    cuts=metric_def['cuts']; d=metric_def['direction']
    if d=='high_good':
        s=1
        for c in cuts:
            if value>=c: s+=1
        return s
    elif d=='low_good':
        s=5
        for c in cuts:
            if value>=c: s-=1
        return s
    return None

def _weight(m, fmt, scheme):
    suf = 'v1' if scheme=='definition_v1' else 'v2'
    key = ('weight_dt_'+suf) if fmt=='DT' else ('weight_in_'+suf)
    return m.get(key,0) or 0

def glazegrade(points_by_metric, store_format, defs=None):
    defs=defs or load_defs()
    scheme=defs['scheme_in_use']
    byname={m['name']:m for m in defs['metrics']}
    num=den=0.0
    for metric,p in points_by_metric.items():
        if p is None or metric not in byname: continue
        w=_weight(byname[metric], store_format, scheme)
        if w<=0: continue
        num+=p*w; den+=w
    return round(num/den,2) if den else None
