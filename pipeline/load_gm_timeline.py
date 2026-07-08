#!/usr/bin/env python3
"""Load the gm_timeline table from gm_timeline.csv (the source of truth for GM tenure).

Usage: python load_gm_timeline.py [gm_timeline.csv] [operations_trending.db]

Run AFTER build_db.py (which recreates the table from the legacy v12 snapshot) to apply the
curated CSV. Edit gm_timeline.csv ad-hoc to record a GM change. `source` distinguishes
manual entries from a future Fabric GM feed. Postgres-clean → lifts to Supabase."""
import sys, csv, sqlite3, os
CSVP = sys.argv[1] if len(sys.argv) > 1 else "gm_timeline.csv"
DB   = sys.argv[2] if len(sys.argv) > 2 else "operations_trending.db"
rows = [r for r in csv.DictReader(open(CSVP, newline="", encoding="utf-8-sig")) if (r.get("pc") or "").strip()]
con = sqlite3.connect(DB); cur = con.cursor()
if "source" not in [c[1] for c in cur.execute("PRAGMA table_info(gm_timeline)")]:
    cur.execute("ALTER TABLE gm_timeline ADD COLUMN source TEXT DEFAULT 'manual'")
cur.execute("DELETE FROM gm_timeline")
cur.executemany("INSERT INTO gm_timeline(pc,gm,start,end,source) VALUES(?,?,?,?,?)",
    [(int(r["pc"]), r["gm"], r["start"] or None, r["end"] or None, r.get("source") or "manual") for r in rows])
con.commit()
print("gm_timeline rows loaded:", cur.execute("SELECT COUNT(*) FROM gm_timeline").fetchone()[0])
con.close()
