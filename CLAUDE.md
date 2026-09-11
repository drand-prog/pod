# The Melting Pod — Cross-Platform Analytics Dashboard

This repository holds consolidated listening/viewing analytics for **The Melting Pod**
(an immigration-policy podcast) across four platforms, plus the context needed to turn
it into a dashboard. This file is the brief; read it fully before building.

## Goal

Build an interactive dashboard from `data/consolidated/The_Melting_Pod_Analytics_Consolidated.xlsx`.
Prefer a **single self-contained `index.html`** (charts via a CDN library or inlined JS) so it
opens in any browser and can be hosted on GitHub Pages with no build step. If a build step is
used, keep it minimal and document it in `README.md`.

## Data inventory

### Consolidated workbook — `data/consolidated/…xlsx` (12 tabs)
1. **Platform Summary** — cross-platform KPI matrix (Spotify / Apple / Other platforms / YouTube) with sourced term definitions and the download residual.
2. **Megaphone Downloads** — daily downloads + download reach, 2026-01-01 to 2026-07-15 (time series).
3. **Megaphone Technology** — downloads broken out by listening app (Apple 6,200; Spotify 1,003; long tail), 2026-01-01 to 2026-07-16.
4. **Spotify Streams** — daily streams, all-time (time series).
5. **Spotify Engagement** — daily consumption hours, avg consumption, comments, followers (time series).
6. **Spotify Completion Rates** — per-episode completion %.
7. **Spotify WoW Retention** — week-over-week retention % (time series).
8. **Apple Episodes** — 11 episodes: listeners, engaged listeners, plays, avg consumption.
9. **Apple Top Cities** — listeners by city (Apple).
10. **YouTube Videos** — per-video views, watch time, subscribers, impressions, CTR.
11. **YouTube Traffic Sources** — views/impressions/watch time by source.
12. **YouTube Funnel** — impressions → views → watch time.

### Raw source exports — `data/raw/`
CSV exports the workbook was built from. Use these for programmatic reads rather than
re-parsing the xlsx if easier. `Technology_Performance…csv` is the per-app download report.

### Screenshots — `data/screenshots/`
Platform figures that have no CSV export were transcribed from these into the workbook.
Manifest (only labeled values were transcribed — never numbers read off line/bar charts):
- `Apple_screenshot_1.png` — Apple overview KPIs (Followers 487, Listeners 587, Engaged 409, Plays 10.4K, Time Listened ~1,217 h).
- `Apple_screenshot_3.png` — Apple episodes table.
- `YouTube_screenshot_1.png` — YouTube overview (Views 7,543, Watch time 241.5 h, Subscribers +178).
- `YouTube_screenshot_3.png` — YouTube impressions→watch-time funnel.
- `YouTube_screenshot_4.png` — YouTube traffic sources.
- Three source screenshots are NOT in this repo (Apple listeners-by-city, YouTube per-video
  table, Spotify header). Their data already lives in tabs **Apple Top Cities**, **YouTube Videos**,
  and **Platform Summary** (Spotify 283 followers / 1,934 all-time plays). Treat the workbook as
  the source of truth for those.

## Metric definitions (each platform's own wording — keep these accurate on the dashboard)
- Apple **Plays** = "the total number of times people pressed play on your episode."
- Apple **Listeners** = "the total number of people that listened to or watched your show" (unique; not additive across episodes).
- Apple **Engaged Listeners** = people who listened/watched "at least 20 minutes or 40% of an episode."
- Apple **Followers** = "the total number of people following your show."
- Spotify **Plays** (standard effective 11 Jun 2026) = an episode "watched or listened to for at least thirty seconds."
- YouTube **Views** = times a video was watched (unique + repeat); **Watch time** = total hours watched; **Impressions** = thumbnail shows; **CTR** = share of impressions that led to a view; **Average view duration** = watch time ÷ views.
- Megaphone/IAB **Download** = "a unique file request that was downloaded" (complete or partial), de-duplicated by IP + user agent within 24h.
- Megaphone dashboard, "Growth on Spotify" panel — tooltips, quoted verbatim (confirmed Sep 2026):
  - **Plays** = "The number of times any episode of this show was watched or listened to for at least 30 seconds on Spotify during the selected time period." **Confirmed (not just inferred) to be a verbatim passthrough of Spotify's own "Performance" Plays export**: cross-checked day-by-day against Spotify's own `TheMeltingPod_Performance…csv` export for the full Jan 1–Sep 11, 2026 range — 253 of 254 days matched exactly (one day off by 1, almost certainly an export-timing difference). Megaphone is relaying Spotify's own numbers here, not independently measuring. This is still a *different* figure from the Platform Summary's Spotify Plays row, though — that one is the all-platform-visible, all-time cumulative total (currently 2,250, summed from Spotify's own per-episode export), not a Megaphone-relayed daily series for a specific window.
  - **Confirmed reach by plays** = "The number of distinct people who actively watched or listened to any episode of your podcast on Spotify." Same confirmation: matches Spotify's own `Performance` "Audience" column exactly across all 254 days (0 mismatches).
  - **Downloads** = "The total number of downloads for all episodes for this podcast across all platforms." (Includes Spotify's downloads, same figure the Megaphone Technology tab breaks out by app.)
  - **Downloads reach** = "The total number of households, or IP addresses, that downloaded an episode of your podcast." (Confirms the interpretation already used for the "Download reach" column in the daily downloads CSV and for aggregating the S3 delivery export in `scripts/fetch_megaphone.py` — unique households/IPs, not additive across days.)
Sources: podcasters.apple.com/support/5392-listener-analytics ; newsroom.spotify.com (11 Jun 2026) ; support.google.com/youtube ; IAB Tech Lab Podcast Measurement (Megaphone certification) ; Megaphone dashboard UI tooltips (Growth on Spotify panel, screenshotted by the project owner Sep 2026) ; Spotify's own `Performance` export (screenshotted/exported by the project owner, cross-checked Sep 2026).

## Critical caveats — surface these on the dashboard, do NOT smooth them away
1. **Different units, not interchangeable:** Megaphone counts *downloads*, Apple/Spotify count *plays*, Spotify separately reports *streams*, YouTube counts *views*. Do not sum or directly compare across these as if identical.
2. **Different time windows:** Spotify & Apple figures are all-time; YouTube is Jan 1 – Jul 14, 2026; Megaphone covers Jan 1 – Jul 15/16. Label periods; don't imply a shared window.
3. **Plays ≠ downloads:** Apple "plays" (~10,380) exceed Megaphone's *total* downloads (9,585) because a play is a playback event (repeats included) while a download is one de-duplicated file request.
4. **"Other platforms" residual (the correct method):** use Megaphone's own per-app DOWNLOAD figures, not native play counts. Other = Megaphone total (9,632) − Apple downloads (6,200) − Spotify downloads (1,003) = **2,429**. This is the ONLY valid basis for that residual; subtracting native plays gives a meaningless negative.
5. **Snapshots vs. trends:** only Megaphone daily downloads, Spotify daily streams/engagement, and Spotify WoW retention are true time series → chart them. Apple and YouTube figures are single point-in-time snapshots → render as KPI cards / rank bars, not fabricated trend lines.
6. **Two Megaphone totals differ slightly:** app report 9,632 (through Jul 16) vs daily tab 9,585 (through Jul 15) — one extra day plus rounding. Use the app-report total for the residual so it ties out internally.

## Refresh workflow (design for this)
Each month Doug will drop new platform exports into `data/raw/` (and any new screenshots into
`data/screenshots/`) and re-run the dashboard build. Keep the parsing tolerant of the existing
file-name patterns and date ranges, and centralize the screenshot-derived figures (the ones with
no CSV) somewhere obvious to update by hand.

## Style / accuracy expectations (from the project owner)
- Do not present inferred or fabricated figures; if a value isn't in the data, show it as unavailable rather than guessing.
- Flag uncertainty and comparability limits in the UI rather than hiding them.
- Do not use the Center for Immigration Studies (CIS) as a source if any external context is ever added.
