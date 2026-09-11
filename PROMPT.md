Read CLAUDE.md in full first — it contains the data inventory, metric definitions, and the
comparability caveats that must be respected.

Then build an interactive analytics dashboard for The Melting Pod from
data/consolidated/The_Melting_Pod_Analytics_Consolidated.xlsx (fall back to the CSVs in data/raw/
where easier to parse).

Requirements:
- Output a single self-contained index.html that opens in any browser with no build step
  (charts via a CDN charting library or inlined JS are fine). If you add a build step, keep it
  minimal and document it in README.md.
- Top section: a cross-platform KPI strip mirroring the "Platform Summary" tab, including the
  "Other platforms" download residual computed the correct way (Megaphone total 9,632 − Apple
  6,200 − Spotify 1,003 = 2,429).
- Per-platform panels:
  - Megaphone: daily downloads trend + the by-app breakdown (Megaphone Technology tab).
  - Spotify: streams trend, engagement (consumption/followers) trend, week-over-week retention
    trend, and per-episode completion rates.
  - Apple: overview KPIs, episodes table, listeners-by-city bar chart.
  - YouTube: overview KPIs, per-video table, traffic sources, and the impressions→watch-time funnel.
- Chart the true time series (Megaphone daily downloads, Spotify daily streams/engagement, Spotify
  WoW retention). Render Apple/YouTube snapshot figures as KPI cards or rank bars — do NOT invent
  trend lines for point-in-time data.
- Keep the comparability caveats from CLAUDE.md visible in the UI (units differ: downloads vs plays
  vs streams vs views; time windows differ; plays are not downloads). Do not present any figure that
  isn't in the data; show unavailable values as such.

Before writing code, restate your plan and the panel layout. Then implement, open the dashboard,
verify the numbers against the workbook, and push a branch.
