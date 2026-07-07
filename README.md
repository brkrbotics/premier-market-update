# Premier Market Update — auto-publishing dashboard

Live dashboard: **https://brkrbotics.github.io/premier-market-update/**
(iframed into https://www.premierportal.ca/market-update and
https://brokerbotics.ai/marketupdate)

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
- drafts the per-region agent commentary from the new numbers (Claude API),
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
| `generate_commentary.py` | Drafts the per-region agent commentary via the Claude API |
| `commentary.json` | The current month's commentary (generated; period-stamped) |
| `mw-historic-data.xlsx` | The 8-tab historic dataset (source of truth) |
| `template.html` | Dashboard template (`{{DATA}}` and `{{COMMENTARY}}` placeholders) |
| `index.html` | The generated dashboard that GitHub Pages serves |
| `.github/workflows/update.yml` | The automation |

## Agent commentary

Each month the Action drafts a short "what this means for your clients" note under
the KPI cards and under each chart, per region, written for real estate agents and
grounded only in that month's extracted numbers.

- Requires a repo secret named **`ANTHROPIC_API_KEY`**
  (Settings → Secrets and variables → Actions). No key = the step skips quietly and
  the dashboard publishes without new commentary.
- The commentary step can never block the data: it runs `continue-on-error`, and the
  page hides any commentary whose month doesn't match the data (`commentary.json`
  carries a `period` stamp the template checks).
- To redo a month by hand: run `python generate_commentary.py --force` locally with
  the key in the environment, or edit `commentary.json` directly and re-run the
  workflow (Actions → Run workflow).
- The same page is embedded on premierportal.ca and brokerbotics.ai, so it stays
  neutral: no brokerage or company branding on the page itself.

## Checking on a run

**Actions** tab → newest "Update market dashboard" run. Green ✓ = published.
If it's red, open the run to see which step failed (usually a misnamed PDF or an
unexpected report layout → the extractor finds ≠ 41 region rows).
