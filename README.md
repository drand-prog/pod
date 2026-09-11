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
- `scripts/fetch_megaphone.py` — pulls Megaphone's downloads + technology-performance reports via
  its API (see [Automation](#automation) below — currently a stub for the actual report calls).
- `.github/workflows/rebuild-dashboard.yml` — auto-rebuilds `index.html` whenever `data/` changes.
- `.github/workflows/megaphone-refresh.yml` — scheduled Megaphone pull + PR (see below).

## Build step
`index.html` is generated, not hand-edited. The build step is one dependency-light Python script:

```
pip install openpyxl
python3 scripts/build_data.py
```

It reads the raw CSVs and the consolidated workbook, assembles one JSON blob, and inlines it (plus
the vendored Chart.js) into `templates/index_template.html` to produce `index.html`. Re-run it any
time the source data changes — the output is fully deterministic from the inputs. Everything
Megaphone-derived (totals, period labels, the "Other platforms" residual, the caveats that quote
those numbers) is computed fresh from the current `data/raw/` CSVs on every build, not hardcoded —
so a refreshed CSV can't leave stale numbers behind elsewhere on the page.

## Viewing the dashboard
- Simplest: open `index.html` directly in a browser — no server required.
- Hosted: enable **GitHub Pages** (Settings → Pages → deploy from branch), or the project's
  existing Vercel import, for a shareable URL.

## Automation

Automation potential differs by platform, since only Megaphone has a documented API for pulling
your own analytics — Apple Podcasts Connect and Spotify for Creators don't expose one. See the
`.github/workflows/` files for the two pieces:

**1. Auto-rebuild (`rebuild-dashboard.yml`) — fully working.** Any push that touches `data/raw/`,
`data/consolidated/`, `templates/`, `vendor/`, or `scripts/build_data.py` triggers a rebuild that
commits the regenerated `index.html` back to that branch automatically. This removes "remember to
run the build script" from the monthly refresh, including for Apple/Spotify — drop the exported
CSVs and updated workbook, push, and the dashboard rebuilds itself.

**2. Scheduled Megaphone pull (`megaphone-refresh.yml`) — scaffolded, one piece unfinished.** Runs
monthly (or on demand via `workflow_dispatch`), calls `scripts/fetch_megaphone.py`, and opens a PR
with the refreshed data. It needs three repo secrets first (Settings → Secrets and variables →
Actions): `MEGAPHONE_API_TOKEN` (from Megaphone User Settings), `MEGAPHONE_NETWORK_ID`, and
optionally `MEGAPHONE_PODCAST_ID`. **`fetch_megaphone.py`'s two report-fetching functions currently
raise `NotImplementedError`** — the endpoint for Megaphone's downloads/technology-performance
*reports* (as opposed to its podcast/episode content API, which is implemented and confirmed)
couldn't be verified from the environment this was built in. To finish it: check the Reporting
section at [developers.megaphone.fm](https://developers.megaphone.fm/) for the exact endpoint and
response shape, or ask Megaphone support to enable the
[Metrics Export Service](https://support.megaphone.fm/en/articles/2678649-metrics-export-service)
(hourly exports to an S3 bucket you control) and adapt the script to read from S3 instead.

## Refreshing each month
1. Drop the new Megaphone/Spotify CSV exports into `data/raw/` (same filename patterns), and any
   new screenshots into `data/screenshots/` — or let `megaphone-refresh.yml` do the Megaphone half
   once it's finished (see above).
2. Update `data/consolidated/…xlsx` with the new screenshot-derived figures (Platform Summary,
   Apple Episodes, Apple Top Cities, YouTube Videos/Traffic Sources/Funnel) — these have no CSV
   export and must be transcribed by hand from the new screenshots, same as the initial build.
3. Push. `rebuild-dashboard.yml` regenerates `index.html` automatically — no local build step needed.

Keep the caveats in `CLAUDE.md` in mind when transcribing: don't invent trend lines for Apple/
YouTube snapshots, and compute the "Other platforms" download residual from Megaphone's own
per-app download report, not native play counts (the build script does this automatically for the
Megaphone-derived numbers; the Apple/YouTube screenshot figures still need a careful human transcribing them).
