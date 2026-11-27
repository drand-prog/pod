#!/usr/bin/env python3
"""
Builds the embedded data blob for index.html from the raw platform exports
and the consolidated analytics workbook.

Time series (true trends) are read from data/raw/*.csv, since those are the
platforms' own exports and are easiest to re-parse on a monthly refresh.
Screenshot-derived and snapshot tables (Platform Summary, Apple Episodes,
Apple Top Cities, YouTube Videos/Traffic Sources/Funnel) exist only in the
consolidated workbook (they were hand-transcribed from screenshots with no
CSV export), so those are read from data/consolidated/*.xlsx.

Usage:
    python3 scripts/build_data.py
Regenerates index.html from templates/index_template.html, replacing the
"/*__DASHBOARD_DATA__*/" placeholder with a JSON blob.
"""
import csv
import json
import re
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "raw"
XLSX_PATH = ROOT / "data" / "consolidated" / "The_Melting_Pod_Analytics_Consolidated.xlsx"
TEMPLATE_PATH = ROOT / "templates" / "index_template.html"
CHARTJS_PATH = ROOT / "vendor" / "chart.umd.js"
OUTPUT_PATH = ROOT / "index.html"


def parse_date(s):
    s = s.strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%b %d, %Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    raise ValueError(f"Unrecognized date format: {s!r}")


def fmt(n):
    return f"{n:,}"


def fmt_month_day(iso):
    d = datetime.strptime(iso, "%Y-%m-%d")
    return f"{d.strftime('%b')} {d.day}"


def fmt_duration(total_seconds):
    h, rem = divmod(total_seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"


def fmt_period(start_iso, end_iso):
    start = datetime.strptime(start_iso, "%Y-%m-%d")
    end = datetime.strptime(end_iso, "%Y-%m-%d")
    if start.year == end.year:
        return f"{fmt_month_day(start_iso)} – {fmt_month_day(end_iso)}, {end.year}"
    return f"{fmt_month_day(start_iso)}, {start.year} – {fmt_month_day(end_iso)}, {end.year}"


def extract_date_range_from_filename(path):
    """Best-effort: pull the two YYYY-MM-DD dates out of a report filename.

    Megaphone's exports bake the report's date range into the filename
    (e.g. "...-2026-01-01-2026-07-15.csv" or "..._2026-01-01_-_2026-07-16_.csv").
    A refreshed export is expected to follow the same convention with new
    dates; if a future export doesn't, callers fall back to another source
    of the range rather than failing the whole build.
    """
    found = re.findall(r"\d{4}-\d{2}-\d{2}", path.stem)
    if len(found) == 2:
        return found[0], found[1]
    return None


def extract_asof_date_from_filename(path):
    """Pull the last YYYY-MM-DD date out of a filename (single-date "as of" exports)."""
    found = re.findall(r"\d{4}-\d{2}-\d{2}", path.stem)
    return found[-1] if found else None


def read_csv_rows(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def find_one(pattern):
    """Locate a raw CSV by glob pattern instead of an exact filename.

    Megaphone's exported filenames bake in the report's date range
    (…-2026-01-01-2026-07-15.csv), which changes every refresh — including
    an automated one that writes a new file rather than overwriting the old.
    Glob so build_data.py doesn't need editing each month, and fail loudly
    if the refresh left more than one match lying around.
    """
    matches = sorted(RAW.glob(pattern))
    if not matches:
        raise FileNotFoundError(f"No file in {RAW} matches {pattern!r}")
    if len(matches) > 1:
        raise RuntimeError(
            f"Multiple files match {pattern!r}: {[m.name for m in matches]} — "
            "remove the stale one(s) before rebuilding."
        )
    return matches[0]


def build_megaphone_daily():
    # Megaphone's export is sorted by Downloads descending, not by date — sort
    # chronologically or the daily trend chart plots dates out of order.
    rows = read_csv_rows(find_one("Megaphone_podcast-downloads-performance-*.csv"))
    by_date = {
        parse_date(r["Date"]): {"downloads": int(r["Downloads"]), "reach": int(r["Download reach"])}
        for r in rows
    }

    # Fill any missing calendar dates (e.g. a reporting gap between two
    # separately-exported date ranges) with nulls rather than leaving them
    # out entirely. The chart's x-axis is categorical — it spaces labels
    # evenly regardless of the actual date deltas — so silently omitting
    # gap dates would draw a continuous line straight across a real gap,
    # implying data that doesn't exist. An explicit null date breaks the
    # line there instead (Chart.js doesn't span gaps by default).
    all_dates = sorted(by_date)
    start = datetime.strptime(all_dates[0], "%Y-%m-%d")
    end = datetime.strptime(all_dates[-1], "%Y-%m-%d")
    out = []
    d = start
    while d <= end:
        iso = d.strftime("%Y-%m-%d")
        if iso in by_date:
            out.append({"date": iso, **by_date[iso]})
        else:
            out.append({"date": iso, "downloads": None, "reach": None})
        d += timedelta(days=1)
    return out


def build_art19_daily():
    """Daily downloads from Art19 — the hosting platform the show migrated
    to from Megaphone in Sep 2026. Art19's own export reports
    'unique_listeners_count' where Megaphone reported 'download reach';
    plotted the same way (a secondary line) but the two aren't asserted to
    be the same methodology, just the closest equivalent each platform
    offers.
    """
    rows = read_csv_rows(find_one("Art19_downloads-daily-*.csv"))
    out = [
        {"date": r["time_unit"], "downloads": int(r["downloads"]), "listeners": int(r["unique_listeners_count"])}
        for r in rows
    ]
    out.sort(key=lambda d: d["date"])
    return out


def build_megaphone_technology():
    """Megaphone's by-app download breakdown.

    Unlike the daily downloads export, this report is queried for a fixed
    date range rather than auto-expanding, so a refresh adds a new
    non-overlapping window instead of extending the old one. Combines every
    Technology_Performance*.csv file present by summing each app's downloads
    across all windows, and reports any real gap between them explicitly
    rather than implying continuous since-launch coverage.
    """
    paths = sorted(RAW.glob("Technology_Performance*.csv"))
    if not paths:
        raise FileNotFoundError(f"No file in {RAW} matches 'Technology_Performance*.csv'")
    windows = []
    totals = {}
    for path in paths:
        date_range = extract_date_range_from_filename(path)
        if date_range is None:
            raise RuntimeError(f"Can't parse the date range from filename: {path.name}")
        windows.append(date_range)
        for r in read_csv_rows(path):
            app = r["APPLICATION"].strip()
            downloads = int(r["DOWNLOADS"].replace(",", ""))
            totals[app] = totals.get(app, 0) + downloads

    windows.sort()
    combined_total = sum(totals.values())
    out = [
        {"app": app, "downloads": downloads, "pct": round(downloads / combined_total, 4)}
        for app, downloads in sorted(totals.items(), key=lambda kv: -kv[1])
    ]

    gaps = []
    for (_, prev_end), (next_start, _) in zip(windows, windows[1:]):
        prev_end_dt = datetime.strptime(prev_end, "%Y-%m-%d")
        next_start_dt = datetime.strptime(next_start, "%Y-%m-%d")
        if (next_start_dt - prev_end_dt).days > 1:
            gap_start = (prev_end_dt + timedelta(days=1)).strftime("%Y-%m-%d")
            gap_end = (next_start_dt - timedelta(days=1)).strftime("%Y-%m-%d")
            gaps.append((gap_start, gap_end))

    return out, windows[0][0], windows[-1][1], gaps


def build_megaphone_episodes():
    """Real per-episode Megaphone downloads, cross-checked against a second,
    independently-exported report before trusting either.

    Megaphone's per-episode report gives lifetime-to-date downloads plus
    fixed-window downloads at 24 hours and 7 days after release — no
    first-30-days column exists, unlike the IAB industry benchmark this
    feeds into. First-7-days is the closest real (not derived/estimated)
    proxy: since downloads only accumulate over time, each episode's true
    first-30-days figure must be >= its first-7-days figure, making the
    7-day average a legitimate lower bound rather than a guess.
    """
    downloads_rows = read_csv_rows(find_one("Megaphone_top-episode-report-*.csv"))
    reach_rows = read_csv_rows(find_one("Megaphone_download-reach-by-episode-*.csv"))
    reach_by_title = {r["EPISODE"].strip(): r for r in reach_rows}

    out = []
    for r in downloads_rows:
        title = r["EPISODE"].strip()
        downloads = int(float(r["DOWNLOADS"]))
        first24h = int(float(r["FIRST 24 HOURS"]))
        first7d = int(float(r["FIRST 7 DAYS"]))
        to_date = int(float(r["TO DATE"]))
        if downloads == 0 and first24h == 0 and first7d == 0 and to_date == 0:
            # A real Megaphone-tracked episode never reads literal zero across
            # every single metric — this is Megaphone's report listing a
            # placeholder row for an episode released after the Art19
            # migration (it knows the episode exists from the feed, but has
            # no actual download data for it). It's also absent from the
            # reach-by-episode report entirely, which would otherwise raise
            # the "missing from the reach report" error below. Skip it here;
            # this episode's real data lives in ART19_EPISODE_30DAY instead.
            continue
        reach_row = reach_by_title.get(title)
        if reach_row is None:
            raise RuntimeError(f"Episode {title!r} is in the downloads report but missing from the reach report")
        reach_downloads = int(reach_row["DOWNLOADS"].replace(",", ""))
        if abs(reach_downloads - downloads) > 1:
            raise RuntimeError(
                f"DOWNLOADS mismatch for {title!r} between the two Megaphone episode reports: "
                f"{downloads} vs {reach_downloads} — one of them may be stale."
            )
        out.append(
            {
                "title": title,
                "published": parse_date(r["PUBLISHED"]),
                "downloadsToDate": downloads,
                "first24h": int(float(r["FIRST 24 HOURS"])),
                "first7d": int(float(r["FIRST 7 DAYS"])),
                "reach": int(reach_row["Download reach"].replace(",", "")),
                "reachPct": float(reach_row["% OF TOTAL"].rstrip("%")) / 100.0,
            }
        )
    return out


def build_megaphone_episode_first30d():
    """Real first-30-days-after-release downloads per episode.

    Megaphone's per-episode report has no first-30-days column of its own
    (see build_megaphone_episodes), but it can be filtered to an arbitrary
    date range and broken down by episode — so a report filtered to exactly
    [release date, release date + 29] gives each episode's own true 30-day
    figure. The project owner pulls one such report per episode once that
    episode's window has fully elapsed; each file covers one window and
    reports every episode's downloads *during that window* (not since each
    episode's own release), so only the row(s) whose PUBLISHED date equals
    the window's start date are that window's actual target episode(s) —
    everything else in the file is a different episode's incidental
    activity during someone else's window and must be ignored.

    Returns a dict of {episode title: first-30-days downloads}, containing
    only episodes whose window has actually been pulled so far — not all
    episodes will have one yet, and that's an honest gap, not an error.
    """
    out = {}
    for path in sorted(RAW.glob("Megaphone_episode-30day-report-*.csv")):
        window_range = extract_date_range_from_filename(path)
        if window_range is None:
            raise RuntimeError(f"Can't parse the 30-day window's date range from filename: {path.name}")
        window_start = window_range[0]
        rows = read_csv_rows(path)
        matched_any = False
        for r in rows:
            if parse_date(r["PUBLISHED"]) == window_start:
                out[r["EPISODE"].strip()] = int(float(r["DOWNLOADS"]))
                matched_any = True
        if not matched_any:
            raise RuntimeError(
                f"{path.name}: no episode was published on {window_start} (the window's own start date) — "
                "check that this report was filtered to the right episode's release date."
            )
    return out


def build_megaphone_geo():
    """Megaphone's own country breakdown of downloads, for a fixed window
    (queried by date range, not auto-expanding — see build_megaphone_technology).
    Separate from Spotify's and YouTube's own geo breakdowns: this one is
    downloads across every platform Megaphone sees, not plays/views on one
    specific platform.
    """
    path = find_one("Megaphone_geolocation-performance-*.csv")
    rows = read_csv_rows(path)
    total = sum(int(r["DOWNLOADS"].replace(",", "")) for r in rows)
    out = [
        {
            "country": r["Country"].strip(),
            "downloads": int(r["DOWNLOADS"].replace(",", "")),
            "pct": round(int(r["DOWNLOADS"].replace(",", "")) / total, 4),
        }
        for r in rows
    ]
    out.sort(key=lambda d: -d["downloads"])
    return out, extract_date_range_from_filename(path)


def build_megaphone_time_of_day():
    """Megaphone's downloads-by-day-and-hour breakdown, for the same kind of
    fixed window as build_megaphone_geo. Megaphone's own day/hour label
    (e.g. "Mon 12am") is split into day + time here; the frontend supplies
    the canonical Mon-Sun / 12am-11pm ordering for the heatmap grid, since
    that ordering is fixed and isn't itself data.
    """
    path = find_one("Megaphone_time-of-day-performance-*.csv")
    rows = read_csv_rows(path)
    out = []
    for r in rows:
        day, time = r["DAY / TIME"].split(" ", 1)
        out.append({"day": day, "time": time, "downloads": int(r["DOWNLOADS"])})
    return out, extract_date_range_from_filename(path)


def build_spotify_streams():
    rows = read_csv_rows(RAW / "Spotify_TheMeltingPod_Streams_all-time.csv")
    out = [{"date": parse_date(r["Date"]), "streams": int(r["Streams"])} for r in rows]
    out.sort(key=lambda d: d["date"])
    return out


def build_spotify_engagement():
    rows = read_csv_rows(RAW / "Spotify_TheMeltingPod_Engagement_all-time.csv")
    out = []
    for r in rows:
        followers = r["Followers"].strip()
        out.append(
            {
                "date": parse_date(r["Date"]),
                "consumptionHours": float(r["Consumption time (hours)"] or 0),
                "avgConsumptionHours": float(r["Average consumption time (hours)"] or 0),
                "comments": int(r["Comments"]) if r["Comments"].strip() else None,
                "followers": int(followers) if followers else None,
            }
        )
    out.sort(key=lambda d: d["date"])
    return out


def build_spotify_completion():
    rows = read_csv_rows(RAW / "Spotify_TheMeltingPod_EpisodeCompletionRates.csv")
    return [
        {
            "title": r["Episode title"].strip(),
            "completionPct": round(float(r["Completion rate (%)"]), 2),
            "publishDate": parse_date(r["Publish date"]),
        }
        for r in rows
    ]


def build_spotify_retention():
    # Filename bakes in a date range that changes every refresh, same
    # reason the Megaphone reports are looked up by glob instead of exact
    # name — avoids another hardcoded-filename edit next time.
    rows = read_csv_rows(find_one("Spotify_TheMeltingPod_WeekOverWeekRetention_*.csv"))
    out = [
        {"weekStart": parse_date(r["Week starting"]), "retentionPct": round(float(r["Retention rate (%)"]), 2)}
        for r in rows
    ]
    out.sort(key=lambda d: d["weekStart"])
    return out


def build_spotify_age_gender():
    # Percentages are already 0-100 (not 0-1 fractions) and each row's Total
    # is the sum of its own Male/Female/Non-binary/Not-specified columns —
    # confirmed for every row, so the frontend can render this as one
    # stacked bar per age bracket without recomputing a total.
    path = find_one("Spotify_TheMeltingPod_AgeByGender_*.csv")
    rows = read_csv_rows(path)
    out = [
        {
            "age": r["Age"].strip(),
            "totalPct": round(float(r["Total percentage"]), 2),
            "malePct": round(float(r["Male percentage"]), 2),
            "femalePct": round(float(r["Female percentage"]), 2),
            "nonbinaryPct": round(float(r["Non-binary percentage"]), 2),
            "notSpecifiedPct": round(float(r["Not specified percentage"]), 2),
        }
        for r in rows
    ]
    return out, extract_date_range_from_filename(path)


def build_spotify_audience_segments():
    path = find_one("Spotify_TheMeltingPod_AudienceSegments_*.csv")
    rows = read_csv_rows(path)
    out = [
        {
            "date": parse_date(r["Date"]),
            "total": int(r["Total"]),
            "returning": int(r["Returning"]),
            "new": int(r["New"]),
        }
        for r in rows
    ]
    out.sort(key=lambda d: d["date"])
    return out


def build_spotify_episode_plays():
    path = find_one("Spotify_TheMeltingPod_EpisodePlaysSincePublished_*.csv")
    rows = read_csv_rows(path)
    out = [
        {
            "title": r["Episode name"].strip(),
            "publishDate": parse_date(r["Publish date"]),
            "plays": int(r["Plays"]),
        }
        for r in rows
    ]
    out.sort(key=lambda d: d["plays"], reverse=True)
    return out, extract_asof_date_from_filename(path)


def build_spotify_geo():
    # Percentage is already 0-100, of all-time listeners by country (sums to
    # ~100 across the full list) — same shape as Apple's Top Cities panel,
    # just percentage-based instead of raw listener counts.
    path = find_one("Spotify_TheMeltingPod_GeoLocation_*.csv")
    rows = read_csv_rows(path)
    out = [{"country": r["Geo"].strip(), "pct": round(float(r["Percentage"]), 2)} for r in rows]
    out.sort(key=lambda d: d["pct"], reverse=True)
    return out, extract_asof_date_from_filename(path)


def build_youtube_device_type():
    path = find_one("YouTube_TheMeltingPod_DeviceType_*.csv")
    rows = read_csv_rows(path)
    total_row = None
    devices = []
    for r in rows:
        entry = {
            "device": r["Device type"].strip(),
            "views": int(r["Views"]),
            "watchTimeHours": float(r["Watch time (hours)"]),
            "avgPctViewed": round(float(r["Average percentage viewed (%)"]), 2),
        }
        if entry["device"] == "Total":
            total_row = entry
        else:
            devices.append(entry)
    return {"devices": devices, "total": total_row}, extract_asof_date_from_filename(path)


def build_youtube_geography():
    # YouTube's own export, not a screenshot — but still only breaks out a
    # handful of countries (small channel; likely below its per-country
    # reporting threshold for the rest), so the listed rows don't sum to the
    # Total the way Apple's truncated-screenshot cities don't either.
    path = find_one("YouTube_TheMeltingPod_Geography_*.csv")
    rows = read_csv_rows(path)
    total_row = None
    countries = []
    for r in rows:
        entry = {
            "country": r["Geography"].strip(),
            "views": int(r["Views"]),
            "avgViewDuration": r["Average view duration"].strip() or None,
            "watchTimeHours": float(r["Watch time (hours)"]),
        }
        if entry["country"] == "Total":
            total_row = entry
        else:
            countries.append(entry)
    countries.sort(key=lambda d: d["views"], reverse=True)
    return {"countries": countries, "total": total_row}, extract_asof_date_from_filename(path)


def build_youtube_audience_segments():
    # 28-day new/casual/regular viewer counts. Verified these three columns
    # sum exactly to the separate "Monthly audience" export for all 192
    # overlapping days — that file is a pure redundant total, so it isn't
    # read at all; this stacked series already reconstructs it as the sum
    # of the three bands.
    path = find_one("YouTube_TheMeltingPod_NewCasualRegularViewers_*.csv")
    rows = read_csv_rows(path)
    out = [
        {
            "date": parse_date(r["Date"]),
            "new": int(r["28-day new viewers"]),
            "casual": int(r["28-day casual viewers"]),
            "regular": int(r["28-day regular viewers"]),
        }
        for r in rows
    ]
    out.sort(key=lambda d: d["date"])
    return out


def build_youtube_subscribers_daily():
    path = find_one("YouTube_TheMeltingPod_SubscribersDaily_*.csv")
    rows = read_csv_rows(path)
    out = [{"date": parse_date(r["Date"]), "subscribers": int(r["Subscribers"])} for r in rows]
    out.sort(key=lambda d: d["date"])
    return out


# Full-length episode titles, used to separate YouTube's per-video exports
# into real episodes vs. short-form clips/Shorts (title-matched against this
# list rather than a duration cutoff, since the Trailer is legitimately short
# but is still a tracked episode everywhere else on the dashboard).
CANONICAL_EPISODE_TITLES = {
    "Welcome to The Melting Pod",
    "Biden and the Border (Part 1), with Charanya Krishnaswami",
    "Biden and the Border (Part 2), with Charanya Krishnaswami",
    "Trailer: What is The Melting Pod?",
    "The Heartbreak of Immigration Reform, with Esther Olavarria",
    "DACA and the Dreamer Movement, with Felicia Escobar Carrillo",
    "Who’s Afraid of the H-1B?",
    "The Truth About ICE, with Jason Houser",
    "News Digest: Green Card Chaos, Hunger Strikes, and Stern Dad Energy",
    "The Reality-Based Economics of Immigration, with Jenny Hunt",
    "A Winning Message on Immigration, with Beatriz Lopez",
    "Three Big Lies About Immigrants and Fraud, with Sonya Schwartz",
    "Confessions of an Immigration Bureaucrat, with Michael Valverde",
    "Asylum Is Not a Loophole, with Colleen Zengotitabengoa",
    "News Dive: The Great De-Legalization, H-1B Follies, and Midterm Predictions",
    "From Suing the Man to Serving the Man, with Lucas Guttentag",
    "The Diplomat, with Rena Bitter",
    "Midterm Special: Counting Wins, Drinking Tears, and Going Big",
    "Real Talk About Latino Voters, with Felicia Escobar Carrillo",
    "Truth, Lies, and Immigration Statistics, with Marc Rosenblum",
}


def build_youtube_clips():
    """Short-form YouTube clips/Shorts — everything in the per-video table
    that isn't a full tracked episode (see CANONICAL_EPISODE_TITLES). Kept
    separate from the main per-video panel, which has always been
    videos-only (Shorts excluded) since the original YouTube refresh.
    """
    path = find_one("YouTube_TheMeltingPod_TableData_*.csv")
    rows = read_csv_rows(path)
    out = []
    for r in rows:
        title = r["Video title"].strip()
        if r["Content"].strip() == "Total" or title in CANONICAL_EPISODE_TITLES:
            continue
        views = int(r["Views"]) if r["Views"].strip() else 0
        out.append(
            {
                "contentId": r["Content"].strip(),
                "title": title,
                "published": parse_date(r["Video publish time"]),
                "durationSeconds": int(r["Duration"]) if r["Duration"].strip() else None,
                "views": views,
                "engagedViews": int(r["Engaged views"]) if r["Engaged views"].strip() else 0,
                "watchTimeHours": float(r["Watch time (hours)"]) if r["Watch time (hours)"].strip() else 0.0,
            }
        )
    out.sort(key=lambda c: -c["views"])
    return out, extract_date_range_from_filename(path)


# Transcribed directly from Art19 dashboard screenshots (no CSV export
# offered for these views as of Nov 26, 2026). All figures are "all-time
# since the Sep 16 migration" unless noted otherwise — Art19 has no visibility
# into anything before the show moved to it from Megaphone.
ART19_GEO_SCREENSHOT_DATE = "2026-11-26"
ART19_GEO = [
    {"country": "United States", "downloads": 2674, "pct": 0.8961},
    {"country": "United Kingdom", "downloads": 70, "pct": 0.0235},
    {"country": "Canada", "downloads": 48, "pct": 0.0161},
    {"country": "Mexico", "downloads": 21, "pct": 0.007},
    {"country": "India", "downloads": 20, "pct": 0.0067},
    {"country": "Australia", "downloads": 14, "pct": 0.0047},
    {"country": "Japan", "downloads": 11, "pct": 0.0037},
    {"country": "Brazil", "downloads": 9, "pct": 0.003},
    {"country": "France", "downloads": 9, "pct": 0.003},
    {"country": "Spain", "downloads": 8, "pct": 0.0027},
]

# Art19's "Top Episodes" (Content tab), all-time since migration — only the
# top 10 rows were visible in the screenshot, so this is a partial list, not
# every episode. For episodes released before the Sep 16 migration (Heartbreak,
# Biden Part 1), this download count only reflects activity since the
# migration, not their true lifetime total — Megaphone already counted their
# earlier downloads, so these numbers understate those two episodes badly and
# are shown as a "since migration" metric, never combined with Megaphone's own
# per-episode figures.
ART19_TOP_EPISODES_SINCE_MIGRATION = [
    {"title": "From Suing the Man to Serving the Man, with Lucas Guttentag", "published": "2026-10-05", "downloads": 537, "pct": 0.18},
    {"title": "The Diplomat, with Rena Bitter", "published": "2026-10-21", "downloads": 488, "pct": 0.1635},
    {"title": "Midterm Special: Counting Wins, Drinking Tears, and Going Big", "published": "2026-11-07", "downloads": 429, "pct": 0.1438},
    {"title": "Real Talk About Latino Voters, with Felicia Escobar Carrillo", "published": "2026-11-11", "downloads": 389, "pct": 0.1304},
    {"title": "Truth, Lies, and Immigration Statistics, with Marc Rosenblum", "published": "2026-11-19", "downloads": 370, "pct": 0.124},
    {"title": "News Dive: The Great De-Legalization, H-1B Follies, and Midterm Predictions", "published": "2026-09-10", "downloads": 129, "pct": 0.0432},
    {"title": "Asylum Is Not a Loophole, with Colleen Zengotitabengoa", "published": "2026-09-04", "downloads": 112, "pct": 0.0375},
    {"title": "Confessions of an Immigration Bureaucrat, with Michael Valverde", "published": "2026-08-13", "downloads": 87, "pct": 0.0292},
    {"title": "The Heartbreak of Immigration Reform, with Esther Olavarria", "published": "2026-03-02", "downloads": 81, "pct": 0.0271},
    {"title": "Biden and the Border (Part 1), with Charanya Krishnaswami", "published": "2026-02-23", "downloads": 64, "pct": 0.0214},
]

# Art19's "Platforms" view (Analytics -> Audience), "All Time" (i.e. since the
# migration — Art19 has no data before it). The listening-app equivalent of
# Megaphone's own Technology Performance report, which only sees downloads
# still flowing through Megaphone's legacy feed URL — this is what's needed
# to extend the "Other platforms" residual to cover the whole year rather
# than just Megaphone's (incomplete, post-migration) side of it. Screenshot's
# own total (2,990) is a few downloads off the Art19 daily-chart total (2,984)
# — same small reconciliation gap seen elsewhere between two exports of
# Art19 data, not an error worth chasing further.
ART19_TECHNOLOGY = [
    {"app": "Apple Podcasts", "downloads": 1788},
    {"app": "Spotify", "downloads": 459},
    {"app": "Chrome", "downloads": 171},
    {"app": "Overcast", "downloads": 132},
    {"app": "Pocket Casts", "downloads": 100},
    {"app": "Stagefright", "downloads": 48},
    {"app": "AppleCoreMedia", "downloads": 36},
    {"app": "Edge", "downloads": 33},
    {"app": "Podcast Addict", "downloads": 28},
    {"app": "Downcast", "downloads": 27},
    {"app": "AntennaPod", "downloads": 23},
    {"app": "Podcasts", "downloads": 19},
    {"app": "Safari", "downloads": 17},
    {"app": "Amazon Music", "downloads": 12},
    {"app": "Podbean", "downloads": 11},
    {"app": "Player", "downloads": 10},
    {"app": "Podcast Republic", "downloads": 9},
    {"app": "Firefox", "downloads": 7},
    {"app": "Castbox", "downloads": 6},
    {"app": "PodEngine", "downloads": 5},
    {"app": "Podzone", "downloads": 5},
    {"app": "WordPress.com", "downloads": 5},
    {"app": "AirPodcasts", "downloads": 4},
    {"app": "ThePodcastApp", "downloads": 4},
    {"app": "fm.castbox.audiobook.radio.podcast", "downloads": 4},
    {"app": "iTunes", "downloads": 4},
    {"app": "podclone-bot", "downloads": 4},
    {"app": "python-httpx", "downloads": 4},
    {"app": "Amazon Echo", "downloads": 3},
    {"app": "OrcaBackfill", "downloads": 3},
    {"app": "Podscribe", "downloads": 3},
    {"app": "AudioAgiPodcastStereoProbe", "downloads": 2},
    {"app": "Earshot", "downloads": 1},
    {"app": "Google Home", "downloads": 1},
    {"app": "Google Podcasts", "downloads": 1},
    {"app": "Opera", "downloads": 1},
]

# From Art19's "Episode Comparison" tool: true 30-day-since-release downloads,
# usable for the cross-platform benchmark only for episodes released after the
# Sep 16 migration (pre-migration episodes' Art19 downloads are truncated —
# see above). "confirmed": True means the 30-day window has fully elapsed as
# of today (Nov 26, 2026); False means the episode hasn't hit day 30 yet, so
# the figure is downloads-to-date, a real lower bound on the eventual total,
# not a final number.
ART19_EPISODE_30DAY = [
    {"title": "From Suing the Man to Serving the Man, with Lucas Guttentag", "published": "2026-10-05", "downloads": 492, "confirmed": True},
    {"title": "The Diplomat, with Rena Bitter", "published": "2026-10-21", "downloads": 478, "confirmed": True},
    {"title": "Midterm Special: Counting Wins, Drinking Tears, and Going Big", "published": "2026-11-07", "downloads": 429, "confirmed": False},
    {"title": "Real Talk About Latino Voters, with Felicia Escobar Carrillo", "published": "2026-11-11", "downloads": 389, "confirmed": False},
    {"title": "Truth, Lies, and Immigration Statistics, with Marc Rosenblum", "published": "2026-11-19", "downloads": 370, "confirmed": False},
]


def build_youtube_clip_daily(title):
    """Daily engaged-views series for one specific clip, by exact title —
    only clips YouTube's Chart-data export happened to include (recent/
    trending ones) have day-by-day data; older long-tail clips don't.
    """
    path = find_one("YouTube_TheMeltingPod_ChartData_*.csv")
    rows = read_csv_rows(path)
    out = [
        {"date": r["Date"].strip(), "engagedViews": int(r["Engaged views"])}
        for r in rows
        if r["Video title"].strip() == title
    ]
    out.sort(key=lambda d: d["date"])
    return out


def sheet_rows(wb, name):
    ws = wb[name]
    return [row for row in ws.iter_rows(values_only=True) if any(c is not None for c in row)]


def cell_date(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    return v


def build_platform_summary(wb):
    rows = sheet_rows(wb, "Platform Summary")
    # header row is the one starting with "Metric"
    header_idx = next(i for i, r in enumerate(rows) if r[0] == "Metric")
    out = []
    for r in rows[header_idx + 1 :]:
        if r[0] is None or str(r[0]).startswith('"Other platforms"') or str(r[0]).startswith("Subtracting") or str(r[0]).startswith("Definition sources") or str(r[0]).startswith("•"):
            break
        out.append(
            {
                "metric": r[0],
                "spotify": r[1],
                "apple": r[2],
                "other": r[3],
                "youtube": r[4],
                "notes": r[5],
            }
        )
    return out


def build_apple_episodes(wb):
    rows = sheet_rows(wb, "Apple Episodes")
    header_idx = next(i for i, r in enumerate(rows) if r[0] == "Number")
    episodes = []
    for r in rows[header_idx + 1 :]:
        if str(r[0]).startswith("Sum of listed"):
            continue
        episodes.append(
            {
                "number": r[0],
                "name": r[1],
                "releaseDate": r[2],
                "duration": r[3],
                "listeners": r[4],
                "engagedListeners": r[5],
                "plays": r[6],
                "avgConsumption": r[7],
            }
        )
    # Computed from the episode rows rather than read from the sheet's own
    # SUM() cell: openpyxl doesn't evaluate formulas, only cached values, so
    # a formula cell edited programmatically (as opposed to by Excel/Sheets)
    # reads back as None. Summing here is correct regardless of how the
    # workbook was last saved.
    sum_plays = sum(ep["plays"] for ep in episodes)
    return {"episodes": episodes, "sumOfListedPlays": sum_plays}


def build_apple_top_cities(wb):
    rows = sheet_rows(wb, "Apple Top Cities")
    header_idx = next(i for i, r in enumerate(rows) if r[0] == "City")
    all_cities_total = None
    cities = []
    note = None
    for r in rows[header_idx + 1 :]:
        if r[0] == "All Cities":
            all_cities_total = r[1]
            continue
        if r[1] is None:
            # A footnote row (why the list doesn't sum to the total, a
            # display-floor caveat, etc.) rather than a city — recognized by
            # having no listener count, not by matching specific wording, so
            # whatever caveat is actually true this refresh gets surfaced
            # instead of a fixed "truncated" assumption baked into the code.
            note = r[0]
            continue
        cities.append({"city": r[0], "listeners": r[1]})
    return {"allCitiesTotal": all_cities_total, "cities": cities, "note": note}


def build_youtube_videos():
    """Per-video performance table — videos-only (Shorts/clips excluded, see
    CANONICAL_EPISODE_TITLES and build_youtube_clips), full lifetime (Jan 1
    onward). Replaces an earlier xlsx-snapshot version of this table with
    YouTube's own "Content" breakdown export (the same file that feeds the
    clip classification above), now that a real full-history pull exists —
    see build_youtube_clips's docstring for why titles, not content IDs,
    are the join key.

    One title ("Three Big Lies About Immigrants and Fraud...") appears twice
    in this export under two different content IDs — a near-empty duplicate
    (1 view, no publish date) alongside the real video's row. Rather than
    showing a phantom near-zero duplicate row, the one with more views wins.
    """
    path = find_one("YouTube_TheMeltingPod_TableData_*.csv")
    rows = read_csv_rows(path)
    videos_by_title = {}
    overview = None
    for r in rows:
        if r["Content"].strip() == "Total":
            # The channel-wide overview KPIs (videos + Shorts + everything
            # else combined) live in this same export's Total row — using it
            # instead of the Platform Summary workbook's hand-transcribed
            # snapshot means these cards share the same full-lifetime period
            # as the per-video table below, rather than quietly lagging it.
            overview = {
                "views": int(r["Views"]),
                "watchTimeHours": float(r["Watch time (hours)"]),
                "subscribersNet": int(r["Subscribers gained"]) - int(r["Subscribers lost"]),
                "impressions": int(r["Thumbnail impressions"]),
                "ctr": round(float(r["Thumbnail click-through rate (%)"]) / 100, 4),
                "avgViewDuration": fmt_duration(sum(int(p) * 60**i for i, p in enumerate(reversed(r["Average view duration"].strip().split(":"))))),
            }
            continue
        title = r["Video title"].strip()
        if title not in CANONICAL_EPISODE_TITLES:
            continue
        views = int(r["Views"]) if r["Views"].strip() else 0
        if title in videos_by_title and videos_by_title[title]["views"] >= views:
            continue
        duration_seconds = int(r["Duration"]) if r["Duration"].strip() else None
        videos_by_title[title] = {
            "title": title,
            "duration": fmt_duration(duration_seconds) if duration_seconds is not None else None,
            "views": views,
            "watchTimeHours": float(r["Watch time (hours)"]) if r["Watch time (hours)"].strip() else 0.0,
            "impressions": int(r["Thumbnail impressions"]) if r["Thumbnail impressions"].strip() else 0,
            "ctr": round(float(r["Thumbnail click-through rate (%)"]) / 100, 4) if r["Thumbnail click-through rate (%)"].strip() else None,
        }
    videos = sorted(videos_by_title.values(), key=lambda v: -v["views"])
    return {"videos": videos, "overview": overview}, extract_date_range_from_filename(path)


def build_youtube_traffic():
    """Views by discovery surface/traffic source — full lifetime (Jan 1
    onward), from YouTube's own CSV export. Replaces an earlier xlsx
    snapshot. viewsPct/watchTimePct (share of channel total) aren't in this
    export, so they're computed here from the export's own "Total" row.
    """
    path = find_one("YouTube_TheMeltingPod_TrafficSource_*.csv")
    rows = read_csv_rows(path)
    out = []
    for r in rows:
        out.append(
            {
                "source": r["Traffic source"].strip(),
                "impressions": int(r["Thumbnail impressions"]) if r["Thumbnail impressions"].strip() else None,
                "ctr": round(float(r["Thumbnail click-through rate (%)"]) / 100, 4) if r["Thumbnail click-through rate (%)"].strip() else None,
                "views": int(r["Views"]),
                "avgViewDuration": r["Average view duration"].strip() or None,
                "watchTimeHours": float(r["Watch time (hours)"]),
            }
        )
    total = next((s for s in out if s["source"] == "Total"), None)
    for s in out:
        s["viewsPct"] = round(s["views"] / total["views"], 4) if total and total["views"] else None
        s["watchTimePct"] = round(s["watchTimeHours"] / total["watchTimeHours"], 4) if total and total["watchTimeHours"] else None
    return out, extract_date_range_from_filename(path)


def build_youtube_funnel(wb):
    rows = sheet_rows(wb, "YouTube Funnel")
    header_idx = next(i for i, r in enumerate(rows) if r[0] == "Stage / metric")
    out = []
    for r in rows[header_idx + 1 :]:
        out.append({"stage": r[0], "value": r[1], "notes": r[2]})
    return out


def main():
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)

    # Megaphone is the one source this pipeline can refresh automatically
    # (scripts/fetch_megaphone.py), so its totals, period labels, and the
    # cross-platform "Other platforms" residual are all derived from the
    # current CSVs here rather than hardcoded — otherwise an automated
    # refresh would silently leave stale numbers in the KPI strip and
    # caveats text even though the underlying data/raw/ files were current.
    megaphone_daily = build_megaphone_daily()
    art19_daily = build_art19_daily()
    megaphone_tech, tech_start, tech_end, tech_gaps = build_megaphone_technology()
    megaphone_geo, megaphone_geo_range = build_megaphone_geo()
    megaphone_time_of_day, megaphone_tod_range = build_megaphone_time_of_day()

    daily_start, daily_end = megaphone_daily[0]["date"], megaphone_daily[-1]["date"]

    daily_tab_total = sum(d["downloads"] for d in megaphone_daily if d["downloads"] is not None)

    # Combined daily-downloads series for the trend chart: the show migrated
    # hosting platforms from Megaphone to Art19 in Sep 2026. Megaphone's own
    # export doesn't stop cleanly at the cutover, though — a backfilled
    # export (pulled during the migration itself, Jan 1-Sep 18) shows it
    # kept counting real downloads for a few more days (legacy RSS
    # subscribers still pointed at Megaphone's old feed URL, presumably),
    # overlapping Art19's first few days (Sep 16-18) rather than leaving a
    # gap. Megaphone's by-app Technology Performance report shows the same
    # pattern continuing for months after (see megaphone_tech above), so
    # this overlap is additive — two different, real audiences — not a
    # double count of the same downloads; overlapping days are summed
    # rather than one source silently overriding the other.
    art19_total = sum(r["downloads"] for r in art19_daily)
    art19_migration_date = art19_daily[0]["date"]
    mp_by_date = {d["date"]: d for d in megaphone_daily if d["downloads"] is not None}
    art19_by_date = {d["date"]: d for d in art19_daily}
    combined_dates = sorted(set(mp_by_date) | set(art19_by_date))
    combined_start = datetime.strptime(combined_dates[0], "%Y-%m-%d")
    combined_end = datetime.strptime(combined_dates[-1], "%Y-%m-%d")
    overlap_dates = sorted(set(mp_by_date) & set(art19_by_date))
    combined_daily = []
    d = combined_start
    while d <= combined_end:
        iso = d.strftime("%Y-%m-%d")
        mp_row = mp_by_date.get(iso)
        art19_row = art19_by_date.get(iso)
        if mp_row and art19_row:
            combined_daily.append(
                {
                    "date": iso,
                    "downloads": mp_row["downloads"] + art19_row["downloads"],
                    "reach": mp_row["reach"] + art19_row["listeners"],
                    "source": "both",
                }
            )
        elif mp_row:
            combined_daily.append({"date": iso, "downloads": mp_row["downloads"], "reach": mp_row["reach"], "source": "megaphone"})
        elif art19_row:
            combined_daily.append({"date": iso, "downloads": art19_row["downloads"], "reach": art19_row["listeners"], "source": "art19"})
        else:
            combined_daily.append({"date": iso, "downloads": None, "reach": None, "source": None})
        d += timedelta(days=1)

    app_report_total = sum(a["downloads"] for a in megaphone_tech)
    apple_downloads = next((a["downloads"] for a in megaphone_tech if a["app"] == "Apple Podcasts"), None)
    spotify_downloads = next((a["downloads"] for a in megaphone_tech if a["app"] == "Spotify"), None)
    if apple_downloads is None or spotify_downloads is None:
        raise RuntimeError(
            "Technology Performance report has no 'Apple Podcasts' or 'Spotify' row — "
            "can't compute the 'Other platforms' residual. Check the raw CSV."
        )
    other_downloads = app_report_total - apple_downloads - spotify_downloads

    # Combined-era version of the same breakdown: Megaphone's Technology
    # Performance report only ever sees the share of downloads still flowing
    # through its own legacy feed URL post-migration, so using it alone for
    # the whole year understates "Other platforms" by missing every app-level
    # download that moved to Art19's new feed. Art19's own "Platforms" view
    # (ART19_TECHNOLOGY) covers that gap — summed by app name with Megaphone's
    # (the two never overlap in time, same as the daily downloads chart).
    art19_tech_total = sum(a["downloads"] for a in ART19_TECHNOLOGY)
    combined_tech_totals = {}
    for a in megaphone_tech + ART19_TECHNOLOGY:
        combined_tech_totals[a["app"]] = combined_tech_totals.get(a["app"], 0) + a["downloads"]
    combined_tech_total = sum(combined_tech_totals.values())
    combined_apple_downloads = combined_tech_totals.get("Apple Podcasts", 0)
    combined_spotify_downloads = combined_tech_totals.get("Spotify", 0)
    combined_other_downloads = combined_tech_total - combined_apple_downloads - combined_spotify_downloads

    spotify_age_gender, age_gender_range = build_spotify_age_gender()
    spotify_episode_plays, episode_plays_asof = build_spotify_episode_plays()
    spotify_geo, spotify_geo_asof = build_spotify_geo()
    youtube_device, youtube_device_asof = build_youtube_device_type()
    youtube_geo, youtube_geo_asof = build_youtube_geography()
    youtube_clips, youtube_clips_range = build_youtube_clips()
    # Picked dynamically (highest-views clip) rather than hardcoded: with the
    # Sep 11-Nov 25 export this was always the Stephen Miller clip (the only
    # one with any activity in that window), but the full-lifetime export
    # shows several older clips have actually accumulated more lifetime views
    # — hardcoding would have kept calling Stephen Miller "top-performing"
    # after that stopped being true.
    top_clip_title = youtube_clips[0]["title"]
    top_clip_daily = build_youtube_clip_daily(top_clip_title)
    youtube_videos_data, youtube_videos_range = build_youtube_videos()
    youtube_traffic, youtube_traffic_range = build_youtube_traffic()

    periods = {
        "spotify": "All-time",
        # Apple's own workbook notes pin this down: the Apple Top Cities
        # sheet says "Full list from the Sep 2026 'See All' cities view",
        # and Apple Episodes lists exactly 15 episodes (the pre-migration
        # count) with the latest being the Sep 10-released one — so this
        # snapshot was taken shortly after that, before the Art19-era
        # episodes existed. Not "all-time" in the Spotify-trend-chart sense
        # (continuously updated since launch): it's a frozen point-in-time
        # pull that hasn't been refreshed since, missing 5 known episodes.
        "apple": "Sep 2026 snapshot (not refreshed since — missing the 5 episodes released after Sep 10, 2026)",
        # The overview KPI cards now come from the same full-lifetime Content
        # export's Total row as the per-video table (see build_youtube_videos),
        # so this period covers both. Only the Impressions->watch-time funnel
        # panel is still on its own older xlsx snapshot (see its own chart-sub
        # text) — YouTube Studio doesn't expose that breakdown any other way.
        "youtube": fmt_period(*youtube_videos_range) if youtube_videos_range else None,
        "youtubeVideoTable": fmt_period(*youtube_videos_range) if youtube_videos_range else None,
        "youtubeTraffic": fmt_period(*youtube_traffic_range) if youtube_traffic_range else None,
        "megaphoneDaily": fmt_period(daily_start, daily_end),
        # The combined chart spans both hosting platforms — kept separate
        # from megaphoneDaily above, which stays Megaphone-only since that's
        # what the "Two Megaphone totals differ" caveat further down is
        # actually about (Megaphone's own two internal reports, unrelated to
        # the Art19 migration).
        "downloadsCombined": fmt_period(combined_dates[0], combined_dates[-1])
        + f" (migrated from Megaphone to Art19 on {fmt_month_day(art19_migration_date)})",
        "art19MigrationDate": art19_migration_date,
        "megaphoneAppReport": fmt_period(tech_start, tech_end) + (
            " (gap: " + "; ".join(f"{fmt_month_day(g[0])}–{fmt_month_day(g[1])}" for g in tech_gaps) + " not included)"
            if tech_gaps else ""
        ),
        # Raw ISO end dates too, so the frontend can compute the gap between
        # the two Megaphone reports itself rather than parsing it back out
        # of the formatted strings above.
        "megaphoneDailyEnd": daily_end,
        "megaphoneAppReportEnd": tech_end,
        "spotifyAgeGender": fmt_period(*age_gender_range) if age_gender_range else "All-time",
        "spotifyEpisodePlaysAsOf": fmt_month_day(episode_plays_asof) if episode_plays_asof else None,
        "spotifyGeoAsOf": fmt_month_day(spotify_geo_asof) if spotify_geo_asof else None,
        "youtubeDeviceAsOf": fmt_month_day(youtube_device_asof) if youtube_device_asof else None,
        "youtubeGeoAsOf": fmt_month_day(youtube_geo_asof) if youtube_geo_asof else None,
        "megaphoneGeo": fmt_period(*megaphone_geo_range) if megaphone_geo_range else None,
        "megaphoneTimeOfDay": fmt_period(*megaphone_tod_range) if megaphone_tod_range else None,
        "youtubeClips": fmt_period(*youtube_clips_range) if youtube_clips_range else None,
        "art19GeoAsOf": fmt_month_day(ART19_GEO_SCREENSHOT_DATE),
    }

    platform_summary = build_platform_summary(wb)
    for row in platform_summary:
        if row["metric"] == "Downloads by app (Megaphone, IAB)":
            # Megaphone's own per-app download report was the sole source for
            # this row, but Megaphone-only is now a full year of undercounting
            # the post-migration period (its own Technology Performance report
            # only ever sees the legacy-feed trickle, not the much larger
            # volume that moved to Art19 — see ART19_TECHNOLOGY above). This
            # row's own "Notes" text was also still a hardcoded leftover
            # (9,632/6,200/1,003/2,429, from some earlier revision) rather
            # than ever being derived — replaced along with the values.
            row["metric"] = "Downloads by app (Megaphone + Art19)"
            row["spotify"] = combined_spotify_downloads
            row["apple"] = combined_apple_downloads
            row["other"] = combined_other_downloads
            row["notes"] = (
                "IAB / Megaphone 'Download' = \"a unique file request that was downloaded\" (complete or "
                "partial), de-duplicated by IP + user agent within 24h; Art19 reports its own equivalent "
                "concept, not independently verified to match IAB's exact methodology. These are each "
                "platform's own per-app/per-technology download counts, combined across the Megaphone era and "
                "the Art19 era since migration (they never overlap in time — see the daily downloads chart's "
                "migration caveat). 'Other platforms' = combined total "
                f"({fmt(combined_tech_total)}) − Apple ({fmt(combined_apple_downloads)}) − Spotify "
                f"({fmt(combined_spotify_downloads)}) = {fmt(combined_other_downloads)} (Chrome, Overcast, "
                "Safari, Pocket Casts, and the long tail). Megaphone excludes YouTube."
            )
        if row["metric"] == "Reporting period":
            # Same staleness problem as the row above, just undiscovered until
            # now: this row is hand-transcribed free text, so refreshing
            # Megaphone/YouTube data (which updates periods{} above and the
            # section headers) silently left this table row on its original
            # "Jan 1 - Jul 14, 2026" text. Derive it from periods{} instead so
            # it can't drift from the section headers again.
            # "Jan 1" elsewhere on the page is just the safe query-start date
            # used when pulling full-history exports, not a claim that
            # activity began then. The real first-activity dates, checked
            # directly against the raw data: Spotify and YouTube both show
            # their first real numbers on Feb 22, 2026 (matching the Feb 22
            # YouTube upload date); Megaphone shows 2 downloads on Feb 13,
            # 2026 — 10 days before the Feb 23 official episode-publish date
            # in its own episode metadata, almost certainly pre-launch test
            # traffic rather than real audience activity. Treated as noise:
            # this row uses Feb 22 uniformly rather than making Megaphone's
            # column inconsistent with the other three over a 2-download blip.
            launch_date = "2026-02-22"
            row["spotify"] = fmt_period(launch_date, tech_end)
            # Every other Apple cell in this table is a frozen Sep 2026
            # snapshot with no Megaphone/Art19 equivalent (downloads-by-app
            # can't infer plays, followers, or engagement) — those are
            # blanked further below rather than shown as stale. The one
            # Apple cell that *is* inferable (downloads, via the app
            # attributed to "Apple Podcasts" in Megaphone+Art19's combined
            # breakdown) covers the full year, so this cell shows that period
            # directly; the "downloads only" caveat itself lives in the
            # table's fine print below instead of crowding this cell.
            row["apple"] = fmt_period(launch_date, tech_end)
            row["youtube"] = fmt_period(launch_date, tech_end)
            row["notes"] = (
                f"All three use the show's real first-activity date ({fmt_month_day(launch_date)}), not Jan 1 "
                "(just the safe query-start date for full-history exports). Megaphone's own raw export actually "
                "shows 2 downloads on Feb 13 — 10 days before the Feb 23 episode-publish date in its own episode "
                f"metadata — treated as pre-launch test noise rather than real activity. Apple's downloads figure "
                "is full-year (via Megaphone + Art19's combined per-app breakdown — see the \"Downloads by app\" "
                f"row), but every other Apple metric is a {periods['apple']} with no Megaphone/Art19 equivalent, so "
                "those cells are blank rather than stale. Different periods — cross-platform totals are not "
                "strictly comparable."
            )
        if row["metric"] == "Plays / Streams / Views":
            # Same staleness class again: this Spotify cell is meant to be the
            # exact sum of the per-episode "Plays since published" export, not
            # a hand-copied figure — so a refresh of that export (new episodes,
            # accumulated plays) silently left this cell behind until now.
            row["spotify"] = sum(e["plays"] for e in spotify_episode_plays)
            row["youtube"] = youtube_videos_data["overview"]["views"]
        if row["metric"] == "Followers / Subscribers":
            row["youtube"] = youtube_videos_data["overview"]["subscribersNet"]
        if row["metric"] == "Listen / Watch time (hours)":
            row["youtube"] = round(youtube_videos_data["overview"]["watchTimeHours"], 1)
        if row["metric"] == "Impressions":
            # Same reasoning as the three rows above: this table and the
            # YouTube overview KPI cards both ultimately come from the same
            # Content export's Total row now (see build_youtube_videos), so
            # keep this cell's YouTube column tied to it directly rather than
            # letting it drift back to the old xlsx snapshot on a refresh.
            row["youtube"] = youtube_videos_data["overview"]["impressions"]

    def metric(name, col):
        """Look up one Platform Summary cell by metric label + platform column.

        The Apple/YouTube overview KPI cards mirror figures that already live
        in the Platform Summary sheet (screenshot-transcribed, same as
        everything else on that tab) — reading them from there instead of a
        separate set of hardcoded constants means updating the workbook is
        enough; there's no second copy of these numbers to remember to edit.
        """
        for row in platform_summary:
            if row["metric"].strip() == name:
                return row[col]
        raise RuntimeError(f"Platform Summary has no {name!r} row")

    apple_plays_header = metric("Plays / Streams / Views", "apple")
    apple_episodes_data = build_apple_episodes(wb)
    megaphone_episodes = build_megaphone_episodes()
    first30d_by_title = build_megaphone_episode_first30d()
    for e in megaphone_episodes:
        first30d = first30d_by_title.get(e["title"])
        if first30d is not None and first30d < e["first7d"]:
            raise RuntimeError(
                f"{e['title']!r}: first-30-days downloads ({first30d}) is less than first-7-days "
                f"({e['first7d']}) — downloads can't un-accumulate, so one of these reports is wrong."
            )
        e["first30d"] = first30d

    # Industry benchmark from a Captivate/IAB "Podcasting Stats" slide (Global/Dax,
    # @podcast411, July 2026): percentile thresholds for downloads-per-episode in
    # the first 30 days after release. Uses each episode's real first-30-days
    # figure where a windowed report has been pulled for it; for episodes whose
    # 30-day window hasn't been pulled yet (or hasn't fully elapsed), falls back
    # to that episode's first-7-days figure as a real lower bound — never a
    # fabricated estimate, since first-30-days can only be >= first-7-days.
    benchmark_brackets = [
        {"topPct": 50, "downloadsPerEp": 28},
        {"topPct": 20, "downloadsPerEp": 189},
        {"topPct": 10, "downloadsPerEp": 677},
        {"topPct": 5, "downloadsPerEp": 1884},
        {"topPct": 2, "downloadsPerEp": 5840},
        {"topPct": 1, "downloadsPerEp": 11716},
    ]
    # Art19 episodes (post-migration, Sep 16 onward) feed the same benchmark
    # alongside the Megaphone-era ones, screenshot-transcribed from Art19's
    # Episode Comparison tool (see ART19_EPISODE_30DAY above).
    #
    # Only episodes with a fully-elapsed, confirmed 30-day window are
    # included in the average and chart — a still-in-progress episode's
    # downloads-to-date (or a pre-30-day first-7-days fallback) is a real
    # lower bound, but blending partial figures into the average pulled it
    # down in a way that looked like a regression rather than what it
    # actually was: fewer confirmed episodes, not worse performance.
    total_episode_count = len(megaphone_episodes) + len(ART19_EPISODE_30DAY)
    confirmed_episode_values = [e["first30d"] for e in megaphone_episodes if e["first30d"] is not None] + [
        e["downloads"] for e in ART19_EPISODE_30DAY if e["confirmed"]
    ]
    exact_30d_count = len(confirmed_episode_values)
    our_blended_avg = round(sum(confirmed_episode_values) / exact_30d_count)
    # Lifetime-to-date average is Megaphone-only: Art19 episodes are all
    # recent enough (and Art19 hasn't exposed a lifetime figure the way
    # Megaphone's per-episode report does) that a "lifetime" number for them
    # would just be their current 30-day-or-less figure, not a real lifetime
    # total — mixing the two would understate it misleadingly.
    our_lifetime_avg = round(sum(e["downloadsToDate"] for e in megaphone_episodes) / len(megaphone_episodes), 1)
    episode_chart = sorted(
        [
            {"title": e["title"], "value": e["first30d"], "source": "megaphone"}
            for e in megaphone_episodes
            if e["first30d"] is not None
        ]
        + [
            {"title": e["title"], "value": e["downloads"], "source": "art19"}
            for e in ART19_EPISODE_30DAY
            if e["confirmed"]
        ],
        key=lambda x: -x["value"],
    )
    benchmark_rows = []
    inserted = False
    for b in benchmark_brackets:
        if not inserted and our_blended_avg < b["downloadsPerEp"]:
            benchmark_rows.append({"label": "The Melting Pod (30-day avg)", "downloadsPerEp": our_blended_avg, "ours": True})
            inserted = True
        benchmark_rows.append({"label": f"Top {b['topPct']}%", "downloadsPerEp": b["downloadsPerEp"], "ours": False})
    if not inserted:
        benchmark_rows.append({"label": "The Melting Pod (30-day avg)", "downloadsPerEp": our_blended_avg, "ours": True})

    data = {
        "generatedAt": datetime.utcnow().strftime("%Y-%m-%d"),
        "periods": periods,
        "platformSummary": platform_summary,
        "megaphone": {
            "daily": combined_daily,
            "byApp": megaphone_tech,
            "appReportTotal": app_report_total,
            "dailyTabTotal": daily_tab_total,
            "art19Total": art19_total,
            "episodes": megaphone_episodes,
            "geo": megaphone_geo,
            "timeOfDay": megaphone_time_of_day,
            "otherResidual": {
                "total": combined_tech_total,
                "apple": combined_apple_downloads,
                "spotify": combined_spotify_downloads,
                "other": combined_other_downloads,
            },
        },
        "art19": {
            "asOfDate": ART19_GEO_SCREENSHOT_DATE,
            "geo": ART19_GEO,
            "topEpisodesSinceMigration": ART19_TOP_EPISODES_SINCE_MIGRATION,
            "technology": sorted(
                [{"app": a["app"], "downloads": a["downloads"], "pct": round(a["downloads"] / art19_tech_total, 4)} for a in ART19_TECHNOLOGY],
                key=lambda a: -a["downloads"],
            ),
        },
        "spotify": {
            "streams": build_spotify_streams(),
            "engagement": build_spotify_engagement(),
            "completion": build_spotify_completion(),
            "wowRetention": build_spotify_retention(),
            "ageGender": spotify_age_gender,
            "audienceSegments": build_spotify_audience_segments(),
            "episodePlays": spotify_episode_plays,
            "geo": spotify_geo,
        },
        "apple": {
            "overview": {
                "followers": metric("Followers / Subscribers", "apple"),
                "listeners": metric("Listeners (unique)", "apple"),
                "engagedListeners": metric("Engaged listeners", "apple"),
                "playsHeader": apple_plays_header,
                "timeListenedHours": metric("Listen / Watch time (hours)", "apple"),
                "timeListenedFollowing": metric("— of which Following (Apple)", "apple"),
                "timeListenedNotFollowing": metric("— of which Not following (Apple)", "apple"),
            },
            **apple_episodes_data,
            "topCities": build_apple_top_cities(wb),
        },
        "youtube": {
            **youtube_videos_data,
            "trafficSources": youtube_traffic,
            "funnel": build_youtube_funnel(wb),
            "deviceType": youtube_device,
            "geo": youtube_geo,
            "audienceSegments": build_youtube_audience_segments(),
            "subscribersDaily": build_youtube_subscribers_daily(),
            "clips": youtube_clips,
            "clipHighlightTitle": top_clip_title,
            "clipHighlightDaily": top_clip_daily,
        },
        "caveats": [
            "Different units, not interchangeable: Megaphone counts downloads, Apple/Spotify count plays, Spotify separately reports streams, YouTube counts views. These are not summed or directly compared as if identical.",
            f"Different time windows: Spotify figures are all-time (continuously updated); Apple is a {periods['apple']}; YouTube is {periods['youtube']}; Megaphone's daily-downloads tab covers {periods['megaphoneDaily']} and its app (Technology) report covers {periods['megaphoneAppReport']}. Periods are labeled throughout — no shared window is implied.",
            f"Plays ≠ downloads: Apple plays (~{fmt(apple_plays_header)}) exceed Megaphone's total downloads ("
            + (
                fmt(app_report_total)
                if daily_tab_total == app_report_total
                else f"{fmt(min(daily_tab_total, app_report_total))}–{fmt(max(daily_tab_total, app_report_total))}"
            )
            + ") because a play is a playback event (repeats included) while a download is one de-duplicated file request.",
            f"“Other platforms” residual uses each platform's own per-app/per-technology DOWNLOAD figures, not native play counts, combined across the Megaphone era and the Art19 era since migration: combined total ({fmt(combined_tech_total)}) − Apple downloads ({fmt(combined_apple_downloads)}) − Spotify downloads ({fmt(combined_spotify_downloads)}) = {fmt(combined_other_downloads)}. (Megaphone-only: {fmt(app_report_total)} total, {fmt(other_downloads)} other — see below for why the combined figure is used instead.)",
            "Snapshots vs. trends: only Megaphone daily downloads, Spotify daily streams/engagement, and Spotify WoW retention are true time series. Apple and YouTube figures are single point-in-time snapshots, shown as KPI cards / rank bars — not fabricated trend lines.",
            (
                f"The app (Technology) report and the daily-downloads tab now agree exactly ({fmt(app_report_total)} through {fmt_month_day(tech_end)}, both reports) — "
                "that Megaphone-only figure, plus Art19's own Platforms breakdown, is combined for the \"Other\" residual below."
                if app_report_total == daily_tab_total and tech_end == daily_end
                else (
                    f"Two Megaphone totals differ: the app (Technology) report totals {fmt(app_report_total)} through "
                    f"{fmt_month_day(tech_end)}, while the daily-downloads tab totals {fmt(daily_tab_total)} through "
                    f"{fmt_month_day(daily_end)}"
                    + (
                        " — one extra day plus rounding."
                        if abs((datetime.strptime(daily_end, '%Y-%m-%d') - datetime.strptime(tech_end, '%Y-%m-%d')).days) <= 2
                        else (
                            ", because the daily-downloads tab hasn't been refreshed as recently as the app report."
                            if datetime.strptime(tech_end, '%Y-%m-%d') > datetime.strptime(daily_end, '%Y-%m-%d')
                            else ", because the app report hasn't been refreshed as recently as the daily downloads."
                        )
                    )
                    + " The app-report total, plus Art19's own Platforms breakdown, is combined for the “Other” residual."
                )
            ),
            f"YouTube's country breakdown only lists {len(youtube_geo['countries'])} countries totaling {fmt(sum(c['views'] for c in youtube_geo['countries']))} views, against a channel total of {fmt(youtube_geo['total']['views'] if youtube_geo['total'] else 0)} — YouTube Studio doesn't break out the long tail below its per-country reporting threshold, so the list is shown as-is rather than padded to 100%.",
            (
                f"The show migrated hosting platforms from Megaphone to Art19 on {fmt_month_day(art19_migration_date)}, 2026. "
                f"The daily downloads chart above combines both ({fmt(daily_tab_total)} Megaphone + {fmt(art19_total)} Art19 = {fmt(daily_tab_total + art19_total)} total), "
                "tagged by source — not asserted to be identical methodology across the switch, just two platforms both reporting downloads. "
                + (
                    f"Megaphone kept reporting a trickle of real downloads well past the migration date — likely legacy RSS subscribers still pointed at its old feed URL — "
                    f"overlapping Art19's own reporting for {len(overlap_dates)} days so far ({fmt_month_day(overlap_dates[0])}–{fmt_month_day(overlap_dates[-1])}); "
                    "those overlapping days are summed (both platforms were genuinely counting real, distinct downloads), not one overriding the other. "
                    if overlap_dates
                    else ""
                )
                + "The country/time-of-day breakdowns and per-episode downloads panels are still Megaphone-only, frozen at their last export "
                f"({fmt_month_day(tech_end)}), since no Art19 equivalents exist yet. The by-app breakdown chart is the same — but the \"Other platforms\" "
                "residual in the Platform Summary table above is not: it now combines Megaphone's Technology Performance report with Art19's own "
                "Platforms breakdown (screenshot-transcribed, since Art19 offers no CSV export for that view), covering the whole year across both eras."
            ),
        ],
        "sources": [
            "Apple Podcasts Connect — Listener analytics: podcasters.apple.com/support/5392-listener-analytics",
            "Spotify — “A New Standard for Podcast Plays” (11 Jun 2026): newsroom.spotify.com",
            "YouTube Help — Impressions & CTR / key metrics: support.google.com/youtube",
            "IAB Tech Lab Podcast Measurement — Megaphone certification (v2.2): compliance.iabtechnologylab.com",
        ],
        # Megaphone's dashboard has its own "Growth on Spotify" panel with
        # tooltip definitions distinct from (and not sourced from) any of
        # the platforms' own docs above — quoted verbatim per the project
        # owner's screenshots (Sep 2026). Confirmed (not just inferred) to
        # be a verbatim passthrough of Spotify's own "Performance" export:
        # cross-checked day-by-day for Jan 1-Sep 11, 2026 — Plays matched
        # 253/254 days exactly, Confirmed-reach-by-plays matched 254/254.
        # Still a different figure from the Platform Summary's Spotify
        # Plays row, though: that one is the all-time cumulative total
        # (2,250), not this daily series for a specific window.
        "megaphoneSpotifyDefinitions": [
            {
                "term": "Plays",
                "definition": "The number of times any episode of this show was watched or listened to for at least 30 seconds on Spotify during the selected time period.",
            },
            {
                "term": "Confirmed reach by plays",
                "definition": "The number of distinct people who actively watched or listened to any episode of your podcast on Spotify.",
            },
            {
                "term": "Downloads",
                "definition": "The total number of downloads for all episodes for this podcast across all platforms.",
            },
            {
                "term": "Downloads reach",
                "definition": "The total number of households, or IP addresses, that downloaded an episode of your podcast.",
            },
        ],
        # External referral events the project owner identified as likely
        # drivers of visible anomalies in the trend charts above — recorded
        # here (not inferred/guessed) so a reader hitting one of these spikes
        # doesn't mistake it for a data error. Kept to two verified cases;
        # each note describes only what's actually visible in the data,
        # without overstating a same-day match that isn't there (the Aug 29
        # case is a days-delayed uptick, not an immediate spike).
        "notableEvents": [
            {
                "date": "2026-05-24",
                "description": "Doug appeared on The Bulwark's YouTube channel. Megaphone downloads jumped from a ~20/day baseline to 354 on May 26, and YouTube subscribers rose from 48 to 164 (net +116) over the following week.",
                "url": "https://www.youtube.com/watch?v=D0Q7Y_1n3-g",
            },
            {
                "date": "2026-08-29",
                "description": "The Bulwark published an article referencing the show. No same-day spike, but Megaphone downloads rose to 263 on Sep 4 (vs. a ~30-50/day baseline) — a smaller, days-delayed uptick rather than an immediate jump.",
                "url": "https://www.thebulwark.com/p/trump-said-hed-target-illegal-immigration",
            },
        ],
        "downloadBenchmark": {
            "source": 'Captivate/IAB "Podcasting Stats — Median & Mean" (Global/Dax, @podcast411, July 2026); IAB numbers for episodes released on Captivate, downloads in the first 30 days after release',
            "rows": benchmark_rows,
            "ourBlendedAvg": our_blended_avg,
            "ourLifetimeAvg": our_lifetime_avg,
            "ourLifetimeEpisodeCount": len(megaphone_episodes),
            "ourEpisodeCount": total_episode_count,
            "ourExact30dCount": exact_30d_count,
            "episodeChart": episode_chart,
        },
    }

    # Blank the Platform Summary's other stale Apple cells now that the
    # Apple section's own overview (above, via metric()) has already
    # captured their real Sep-2026-snapshot values — this only changes what
    # this one table displays. None of these metrics have a Megaphone/Art19
    # equivalent (a download count can't infer plays, followers, listeners,
    # engagement, or watch time), unlike the "Downloads by app" row just
    # above, which is why that row alone keeps a real (and full-year) number.
    for row in platform_summary:
        if row["metric"].strip() in {
            "Followers / Subscribers",
            "Plays / Streams / Views",
            "Listeners (unique)",
            "Engaged listeners",
            "Listen / Watch time (hours)",
            "— of which Following (Apple)",
            "— of which Not following (Apple)",
        }:
            row["apple"] = "—"

    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    json_blob = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    # Guard against a literal "</script>" inside the JSON or vendored lib breaking the page.
    json_blob = json_blob.replace("</script>", "<\\/script>")
    chartjs_src = CHARTJS_PATH.read_text(encoding="utf-8").replace("</script>", "<\\/script>")

    output = template.replace("/*__CHARTJS_LIB__*/", chartjs_src)
    output = output.replace("/*__DASHBOARD_DATA__*/", json_blob)
    if "/*__CHARTJS_LIB__*/" in template and chartjs_src not in output:
        raise RuntimeError("Failed to inline vendored Chart.js")
    if json_blob not in output:
        raise RuntimeError("Placeholder /*__DASHBOARD_DATA__*/ not found in template")
    OUTPUT_PATH.write_text(output, encoding="utf-8")
    print(f"Wrote {OUTPUT_PATH} ({len(output):,} bytes) — {len(json_blob):,} bytes data, {len(chartjs_src):,} bytes vendored Chart.js")


if __name__ == "__main__":
    main()
