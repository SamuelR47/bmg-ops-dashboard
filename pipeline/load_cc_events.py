#!/usr/bin/env python3
"""Load the cc_event table (competitor / cannibalization / new-store events) from cc_events.csv.

Usage: python load_cc_events.py [cc_events.csv] [operations_trending.db]

cc_events.csv is the source of truth for the events shown on the Ops dashboard. `src` tags
each event's origin (e.g. 'dind_26w_proforma' = model-derived, 'manual' = analyst-entered).
Run AFTER build_db.py to apply the curated CSV. Postgres-clean → lifts to Supabase."""
import sys, csv, sqlite3
CSVP = sys.argv[1] if len(sys.argv) > 1 else "cc_events.csv"
DB   = sys.argv[2] if len(sys.argv) > 2 else "operations_trending.db"
def f(x):
    return float(x) if (x or "").strip() not in ("", "None") else None
rows = [r for r in csv.DictReader(open(CSVP, newline="", encoding="utf-8-sig")) if (r.get("pc") or "").strip()]
con = sqlite3.connect(DB); cur = con.cursor()
cur.execute("DELETE FROM cc_event")
cur.executemany("INSERT INTO cc_event(pc,new_store,event_type,date_opened,impact,net,src) VALUES(?,?,?,?,?,?,?)",
    [(int(r["pc"]), r.get("new_store"), r.get("event_type"), r.get("date_opened") or None,
      f(r.get("impact")), f(r.get("net")), r.get("src") or "manual") for r in rows])
con.commit()
print("cc_event rows loaded:", cur.execute("SELECT COUNT(*) FROM cc_event").fetchone()[0])
con.close()
