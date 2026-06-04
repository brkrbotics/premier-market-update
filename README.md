# Premier Market Update — auto-publishing dashboard

Live dashboard: **https://brkrbotics.github.io/premier-market-update/**
(iframed into https://www.premierportal.ca/market-update)

## To update each month — upload one PDF (works from your phone)

1. Go to the **`reports`** folder in this repo:
   https://github.com/brkrbotics/premier-market-update/tree/main/reports
2. Click **Add file → Upload files**.
3. Drop in the new TRREB Market Watch PDF, named **`mwYYMM.pdf`**
   (e.g. June 2026 → `mw2606.pdf`). **The name matters** — `YY` = year, `MM` = month.
4. Click **Commit changes**.

That's it. A GitHub Action automatically:
- extracts the report,
- appends a row to all 8 tabs of `mw-historic-data.xlsx` (recreating the YoY and
  Avg MOI formulas),
- regenerates `index.html`,
- commits the results.

The live page updates within ~1–2 minutes. Nothing to touch on Wix.

You can also trigger a run manually: **Actions → Update market dashboard → Run workflow**.

## What's in here

| File | Purpose |
|---|---|
| `reports/` | Upload the monthly `mwYYMM.pdf` here (this is what triggers everything) |
| `pipeline.py` | Extract → append workbook → regenerate dashboard |
| `extract_trreb.py` | Parses the TRREB PDF tables to CSV |
| `mw-historic-data.xlsx` | The 8-tab historic dataset (source of truth) |
| `template.html` | Dashboard template (`{{DATA}}` placeholder) |
| `index.html` | The generated dashboard that GitHub Pages serves |
| `.github/workflows/update.yml` | The automation |

## Checking on a run

**Actions** tab → newest "Update market dashboard" run. Green ✓ = published.
If it's red, open the run to see which step failed (usually a misnamed PDF or an
unexpected report layout → the extractor finds ≠ 41 region rows).
