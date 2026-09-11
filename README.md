# The Melting Pod — Analytics Dashboard

An interactive, cross-platform analytics dashboard for **The Melting Pod** podcast, built from
consolidated Megaphone / Spotify / Apple Podcasts / YouTube exports.

## What's here
- **`index.html`** — the dashboard. Self-contained: open it directly in any browser (double-click,
  no server, no internet connection needed). Chart.js is vendored inline, and all data is embedded
  in the page at build time.
- `CLAUDE.md` — the full project brief: data inventory, metric definitions, and the comparability
  caveats the dashboard must surface.
- `PROMPT.md` — the original task brief given to Claude Code.
- `data/consolidated/` — the 12-tab analytics workbook (source of truth for screenshot-derived
  tables: Platform Summary, Apple Episodes, Apple Top Cities, YouTube Videos/Traffic/Funnel).
- `data/raw/` — the original per-platform CSV exports (source for the true time series: Megaphone
  daily downloads, Spotify streams/engagement/completion/retention, Megaphone Technology).
- `data/screenshots/` — platform screenshots that had no CSV export (already transcribed into the
  consolidated workbook).
- `templates/index_template.html` — the HTML/CSS/JS shell, with a `/*__DASHBOARD_DATA__*/`
  placeholder for the data and a `/*__CHARTJS_LIB__*/` placeholder for the vendored chart library.
- `vendor/chart.umd.js` — Chart.js 4.4.4 (MIT), vendored so the dashboard has zero external
  network dependencies.
- `scripts/build_data.py` — the build step (see below).

## Build step
`index.html` is generated, not hand-edited. The build step is one dependency-light Python script:

```
pip install openpyxl
python3 scripts/build_data.py
```

It reads the raw CSVs and the consolidated workbook, assembles one JSON blob, and inlines it (plus
the vendored Chart.js) into `templates/index_template.html` to produce `index.html`. Re-run it any
time the source data changes — the output is fully deterministic from the inputs.

## Viewing the dashboard
- Simplest: open `index.html` directly in a browser — no server required.
- Hosted: enable **GitHub Pages** (Settings → Pages → deploy from branch) for a shareable URL.

## Refreshing each month
1. Drop the new Megaphone/Spotify CSV exports into `data/raw/` (same filename patterns), and any
   new screenshots into `data/screenshots/`.
2. Update `data/consolidated/…xlsx` with the new screenshot-derived figures (Platform Summary,
   Apple Episodes, Apple Top Cities, YouTube Videos/Traffic Sources/Funnel) — these have no CSV
   export and must be transcribed by hand from the new screenshots, same as the initial build.
3. Re-run `python3 scripts/build_data.py` to regenerate `index.html`.

Keep the caveats in `CLAUDE.md` in mind when transcribing: don't invent trend lines for Apple/
YouTube snapshots, and compute the "Other platforms" download residual from Megaphone's own
per-app download report, not native play counts.
