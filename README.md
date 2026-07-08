# BMG Operations Diagnostic Dashboard

Live operations dashboard for the Bluemont Group donut portfolio (GlazeGrade,
same-store sales, store diagnostics).

## What gets served
`index.html` is the finished, self-contained dashboard. GitHub Pages serves this
one file at the site URL. All data is baked into the file at build time, so there
is no server or database to run for hosting.

## Data currency
Data is **monthly**. The current build covers through **May 2026**. A new month is
added only after the operations scorecard publishes for that period.

## How to refresh (monthly, ~10 min) — from Pete's runbook
1. Pull the 3 Fabric queries (`fabric_export.sql`) from Gold_warehouse and export to
   `ops_metrics.csv`, `sales.csv`, `store_master.csv`.
2. Update `manual_inputs.csv` with the period's Accuracy and Cert SL/AGM values
   (the 2 metrics that are not in Fabric).
3. Rebuild:
   ```
   cd pipeline
   python load_fabric.py            # builds operations_trending.db, recomputes GlazeGrade
   python generate_outputs.py operations_trending.db dashboard_template.html ../index.html glazegrade.xlsx
   ```
4. Commit the updated `index.html`. GitHub Pages redeploys automatically.

## What is intentionally NOT in this repo (see .gitignore)
- `operations_trending.db` — large and fully rebuildable from the pipeline.
- `fabric-mcp/` — carries Fabric connection/auth details; keep out of git.

## Later: full automation
Point `load_fabric.py` at the Fabric SQL endpoint directly (service login) to remove
the manual CSV export, then run the rebuild on a monthly GitHub Actions schedule.
