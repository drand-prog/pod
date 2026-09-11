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
- `scripts/fetch_megaphone.py` — pulls new finalized days from Megaphone's Delivery Export (S3)
  and merges them into `data/raw/` (see [Automation](#automation) below).
- `.github/workflows/rebuild-dashboard.yml` — auto-rebuilds `index.html` whenever `data/` changes.
- `.github/workflows/megaphone-refresh.yml` — daily Megaphone pull + PR (see below).

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

Automation potential differs by platform, since only Megaphone offers a way to pull your own
analytics programmatically — Apple Podcasts Connect and Spotify for Creators don't expose one, and
Megaphone itself has no REST reporting endpoint either; the only supported mechanism is its
**Delivery & Impressions Export** service (see below). See the `.github/workflows/` files for the
two pieces:

**1. Auto-rebuild (`rebuild-dashboard.yml`) — fully working.** Any push that touches `data/raw/`,
`data/consolidated/`, `templates/`, `vendor/`, or `scripts/build_data.py` triggers a rebuild that
commits the regenerated `index.html` back to that branch automatically. This removes "remember to
run the build script" from the monthly refresh, including for Apple/Spotify — drop the exported
CSVs and updated workbook, push, and the dashboard rebuilds itself.

**2. Daily Megaphone pull (`megaphone-refresh.yml`) — fully working, once set up.** Megaphone
pushes gzipped, newline-delimited JSON "delivery export" files to an S3 bucket you own — one file
per day, updated hourly, finalized (won't change again) by ~11:00 UTC the following morning. This
workflow runs daily, downloads any newly finalized day(s), aggregates them (downloads = rows where
`delivery_type == "download"`; per-app breakdown = grouped by `normalized_user_agent`; "download
reach" = distinct hashed listener IDs per day — see the caveat about that last one in
`scripts/fetch_megaphone.py`'s docstring, it's an inference, not something Megaphone's docs state
outright), and **merges** the results into the existing `data/raw/` history via a PR. Merges,
not replaces, because the export has **no backfill** — Megaphone only sends data from the date the
bucket was wired up, forward, so the script always builds on top of whatever's already there rather
than trying to regenerate history from scratch.

One detail worth knowing: because the workflow runs daily and reuses the same branch name
(`automated/megaphone-refresh`), a second day's run before you've merged the first just adds a
commit to the *same* open PR rather than opening a new one — so you won't get a flood of PRs, just
one that stays current until you merge it.

To set it up:
1. **Create an S3 bucket** (if you don't have one) — any AWS region, Block Public Access on.
2. **Ask Megaphone support** (Live Chat, per their docs — typically 1–2 weeks) to enable the
   **Delivery & Impressions Export** service, pointed at that bucket. You only need the delivery
   file, not impressions (that one's ad-verification data).
3. **Create a read-only IAM user** scoped to that bucket alone, e.g.:
   ```json
   {
     "Version": "2012-10-17",
     "Statement": [{
       "Effect": "Allow",
       "Action": ["s3:GetObject", "s3:ListBucket"],
       "Resource": ["arn:aws:s3:::YOUR-BUCKET-NAME", "arn:aws:s3:::YOUR-BUCKET-NAME/*"]
     }]
   }
   ```
   Generate an access key for it.
4. **Add repo secrets** (Settings → Secrets and variables → Actions): `AWS_ACCESS_KEY_ID`,
   `AWS_SECRET_ACCESS_KEY`, `AWS_REGION`, `MEGAPHONE_S3_BUCKET` (and `MEGAPHONE_S3_PREFIX` only if
   Megaphone writes under a prefix rather than the bucket root).

Until setup is complete the workflow will simply fail at the S3 call — nothing runs against data
that isn't there yet, and once the bucket starts receiving files, the very next scheduled run picks
up everything finalized since setup (again, no earlier backfill is possible).

## Refreshing each month
Megaphone refreshes itself daily once set up (see Automation above) — merge its automated PR
whenever convenient. Apple and Spotify still need a human each month, since neither platform
exposes an API:
1. Drop the new Spotify CSV exports into `data/raw/` (same filename patterns), and any new Apple/
   YouTube screenshots into `data/screenshots/`.
2. Update `data/consolidated/…xlsx` with the new screenshot-derived figures (Platform Summary,
   Apple Episodes, Apple Top Cities, YouTube Videos/Traffic Sources/Funnel) — these have no CSV
   export and must be transcribed by hand from the new screenshots, same as the initial build.
3. Push. `rebuild-dashboard.yml` regenerates `index.html` automatically — no local build step needed.

Keep the caveats in `CLAUDE.md` in mind when transcribing: don't invent trend lines for Apple/
YouTube snapshots, and compute the "Other platforms" download residual from Megaphone's own
per-app download report, not native play counts (the build script does this automatically for the
Megaphone-derived numbers; the Apple/YouTube screenshot figures still need a careful human transcribing them).
