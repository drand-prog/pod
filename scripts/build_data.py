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
    total = sum(int(r["DOWNLOADS"]) for r in rows)
    out = [
        {"country": r["Country"].strip(), "downloads": int(r["DOWNLOADS"]), "pct": round(int(r["DOWNLOADS"]) / total, 4)}
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


def build_youtube_videos(wb):
    # Videos-only (Shorts excluded — no fresh per-Short breakdown exists) and
    # no per-video Subscribers figure in this export (dropped entirely,
    # rather than shown as an unavailable placeholder for every row).
    rows = sheet_rows(wb, "YouTube Videos")
    header_idx = next(i for i, r in enumerate(rows) if r[0] == "Video title")
    videos = []
    sum_row = None
    total_row = None
    for r in rows[header_idx + 1 :]:
        if str(r[0]).startswith("Sum of rows"):
            sum_row = r
            continue
        if str(r[0]).startswith("Channel total"):
            total_row = r
            continue
        videos.append(
            {
                "title": r[0],
                "duration": r[1],
                "views": r[2],
                "viewsPct": r[3],
                "watchTimeHours": r[4],
                "watchTimePct": r[5],
                "impressions": r[6],
                "ctr": r[7],
            }
        )
    return {
        "videos": videos,
        "sumOfRows": {"views": sum_row[2], "watchTimeHours": sum_row[4], "impressions": sum_row[6]} if sum_row else None,
        "channelTotal": {"views": total_row[2], "watchTimeHours": total_row[4], "impressions": total_row[6], "ctr": total_row[7]} if total_row else None,
    }


def build_youtube_traffic(wb):
    rows = sheet_rows(wb, "YouTube Traffic Sources")
    header_idx = next(i for i, r in enumerate(rows) if r[0] == "Traffic source")
    out = []
    for r in rows[header_idx + 1 :]:
        out.append(
            {
                "source": r[0],
                "impressions": r[1],
                "ctr": r[2],
                "views": r[3],
                "viewsPct": r[4],
                "avgViewDuration": r[5],
                "watchTimeHours": r[6],
                "watchTimePct": r[7],
            }
        )
    return out


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
    # hosting platforms from Megaphone to Art19 in Sep 2026, so Megaphone's
    # own daily export simply stops (last day: Sep 11) and Art19's own export
    # picks up separately (first day: Sep 16) — a real ~4-day gap during the
    # migration itself, filled with nulls like any other reporting gap
    # rather than glossed over. Each point is tagged with its source so the
    # frontend can render the two periods distinguishably rather than
    # implying one continuous measurement spanning a platform switch.
    art19_total = sum(r["downloads"] for r in art19_daily)
    art19_migration_date = art19_daily[0]["date"]
    mp_by_date = {d["date"]: d for d in megaphone_daily if d["downloads"] is not None}
    art19_by_date = {d["date"]: d for d in art19_daily}
    combined_dates = sorted(set(mp_by_date) | set(art19_by_date))
    combined_start = datetime.strptime(combined_dates[0], "%Y-%m-%d")
    combined_end = datetime.strptime(combined_dates[-1], "%Y-%m-%d")
    combined_daily = []
    d = combined_start
    while d <= combined_end:
        iso = d.strftime("%Y-%m-%d")
        if iso in mp_by_date:
            row = mp_by_date[iso]
            combined_daily.append({"date": iso, "downloads": row["downloads"], "reach": row["reach"], "source": "megaphone"})
        elif iso in art19_by_date:
            row = art19_by_date[iso]
            combined_daily.append({"date": iso, "downloads": row["downloads"], "reach": row["listeners"], "source": "art19"})
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

    spotify_age_gender, age_gender_range = build_spotify_age_gender()
    spotify_episode_plays, episode_plays_asof = build_spotify_episode_plays()
    spotify_geo, spotify_geo_asof = build_spotify_geo()
    youtube_device, youtube_device_asof = build_youtube_device_type()
    youtube_geo, youtube_geo_asof = build_youtube_geography()
    youtube_clips, youtube_clips_range = build_youtube_clips()
    stephen_miller_title = "BREAKING: Stephen Miller interrupts our podcast!"
    stephen_miller_daily = build_youtube_clip_daily(stephen_miller_title)

    periods = {
        "spotifyApple": "All-time",
        # Traffic Sources/Videos/Overview refreshed Sep 2026 (no exact day-count
        # given by the export, unlike the original screenshot); the Funnel tab
        # is not part of this refresh and still shows its own Jan 1-Jul 14 note.
        "youtube": "Jan 1 – Sep 2026",
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
            # Megaphone's own per-app download report is the authoritative
            # source for this row (see the "Other platforms" caveat below) —
            # override whatever was last hand-transcribed into the workbook
            # so this row always matches the fresh CSVs.
            row["spotify"] = spotify_downloads
            row["apple"] = apple_downloads
            row["other"] = other_downloads
        if row["metric"] == "Reporting period":
            # Same staleness problem as the row above, just undiscovered until
            # now: this row is hand-transcribed free text, so refreshing
            # Megaphone/YouTube data (which updates periods{} above and the
            # section headers) silently left this table row on its original
            # "Jan 1 - Jul 14, 2026" text. Derive it from periods{} instead so
            # it can't drift from the section headers again.
            row["spotify"] = periods["spotifyApple"]
            row["apple"] = periods["spotifyApple"]
            row["youtube"] = periods["youtube"]
            row["notes"] = (
                f"Spotify and Apple columns are all-time; YouTube is {periods['youtube']}; "
                f"the Megaphone daily-downloads report covers {periods['megaphoneDaily']}. "
                "Different periods — cross-platform totals are not strictly comparable."
            )
        if row["metric"] == "Plays / Streams / Views":
            # Same staleness class again: this Spotify cell is meant to be the
            # exact sum of the per-episode "Plays since published" export, not
            # a hand-copied figure — so a refresh of that export (new episodes,
            # accumulated plays) silently left this cell behind until now.
            row["spotify"] = sum(e["plays"] for e in spotify_episode_plays)

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
    # Episode Comparison tool (see ART19_EPISODE_30DAY above). "confirmed"
    # mirrors the Megaphone first30d/first7d distinction: True means the
    # 30-day window has fully elapsed (a real final figure); False means
    # it's a downloads-to-date lower bound for a window still in progress.
    all_episode_values = [
        (e["first30d"] if e["first30d"] is not None else e["first7d"], e["first30d"] is not None)
        for e in megaphone_episodes
    ] + [(e["downloads"], e["confirmed"]) for e in ART19_EPISODE_30DAY]
    total_episode_count = len(all_episode_values)
    exact_30d_count = sum(1 for _, exact in all_episode_values if exact)
    # Rounded to a whole number: the fractional digit implied false precision
    # for a figure that's already a blend of exact and lower-bound episodes.
    our_blended_avg = round(sum(v for v, _ in all_episode_values) / total_episode_count)
    # Lifetime-to-date average is Megaphone-only: Art19 episodes are all
    # recent enough (and Art19 hasn't exposed a lifetime figure the way
    # Megaphone's per-episode report does) that a "lifetime" number for them
    # would just be their current 30-day-or-less figure, not a real lifetime
    # total — mixing the two would understate it misleadingly.
    our_lifetime_avg = round(sum(e["downloadsToDate"] for e in megaphone_episodes) / len(megaphone_episodes), 1)
    episode_chart = sorted(
        [
            {"title": e["title"], "value": e["first30d"], "source": "megaphone", "confirmed": True}
            for e in megaphone_episodes
            if e["first30d"] is not None
        ]
        + [
            {"title": e["title"], "value": e["downloads"], "source": "art19", "confirmed": e["confirmed"]}
            for e in ART19_EPISODE_30DAY
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
                "total": app_report_total,
                "apple": apple_downloads,
                "spotify": spotify_downloads,
                "other": other_downloads,
            },
        },
        "art19": {
            "asOfDate": ART19_GEO_SCREENSHOT_DATE,
            "geo": ART19_GEO,
            "topEpisodesSinceMigration": ART19_TOP_EPISODES_SINCE_MIGRATION,
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
            "overview": {
                "views": metric("Plays / Streams / Views", "youtube"),
                "watchTimeHours": metric("Listen / Watch time (hours)", "youtube"),
                "subscribersNet": metric("Followers / Subscribers", "youtube"),
                "impressions": metric("Impressions", "youtube"),
                # Not available as their own Platform Summary cells — only
                # mentioned in that row's prose Notes column, so these two
                # stay hand-maintained constants rather than a fragile
                # regex over free text. Updated from the Sep 2026 Traffic
                # Sources export's Total row (blended across all sources).
                "ctr": 0.0372,
                "avgViewDuration": "4:52",
            },
            **build_youtube_videos(wb),
            "trafficSources": build_youtube_traffic(wb),
            "funnel": build_youtube_funnel(wb),
            "deviceType": youtube_device,
            "geo": youtube_geo,
            "audienceSegments": build_youtube_audience_segments(),
            "subscribersDaily": build_youtube_subscribers_daily(),
            "clips": youtube_clips,
            "clipHighlightTitle": stephen_miller_title,
            "clipHighlightDaily": stephen_miller_daily,
        },
        "caveats": [
            "Different units, not interchangeable: Megaphone counts downloads, Apple/Spotify count plays, Spotify separately reports streams, YouTube counts views. These are not summed or directly compared as if identical.",
            f"Different time windows: Spotify & Apple figures are all-time; YouTube is {periods['youtube']}; Megaphone's daily-downloads tab covers {periods['megaphoneDaily']} and its app (Technology) report covers {periods['megaphoneAppReport']}. Periods are labeled throughout — no shared window is implied.",
            f"Plays ≠ downloads: Apple plays (~{fmt(apple_plays_header)}) exceed Megaphone's total downloads ({fmt(min(daily_tab_total, app_report_total))}–{fmt(max(daily_tab_total, app_report_total))}) because a play is a playback event (repeats included) while a download is one de-duplicated file request.",
            f"“Other platforms” residual uses Megaphone's own per-app DOWNLOAD figures, not native play counts: Megaphone total ({fmt(app_report_total)}) − Apple downloads ({fmt(apple_downloads)}) − Spotify downloads ({fmt(spotify_downloads)}) = {fmt(other_downloads)}. This is the only valid basis for that residual.",
            "Snapshots vs. trends: only Megaphone daily downloads, Spotify daily streams/engagement, and Spotify WoW retention are true time series. Apple and YouTube figures are single point-in-time snapshots, shown as KPI cards / rank bars — not fabricated trend lines.",
            (
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
                + " The app-report total is used for the “Other” residual, which reflects the app report's period, not the daily tab's."
            ),
            f"YouTube's country breakdown only lists {len(youtube_geo['countries'])} countries totaling {fmt(sum(c['views'] for c in youtube_geo['countries']))} views, against a channel total of {fmt(youtube_geo['total']['views'] if youtube_geo['total'] else 0)} — YouTube Studio doesn't break out the long tail below its per-country reporting threshold, so the list is shown as-is rather than padded to 100%.",
            (
                f"The show migrated hosting platforms from Megaphone to Art19 on {fmt_month_day(art19_migration_date)}, 2026. "
                f"The daily downloads chart above combines both ({fmt(daily_tab_total)} Megaphone + {fmt(art19_total)} Art19 = {fmt(daily_tab_total + art19_total)} total), "
                "tagged by source — not asserted to be identical methodology across the switch, just two platforms both reporting downloads. "
                "The by-app breakdown, country/time-of-day breakdowns, and per-episode downloads panels are still Megaphone-only, frozen at their last export "
                f"({fmt_month_day(tech_end)}), since no Art19 equivalents exist yet — the \"Other platforms\" residual above reflects only that frozen Megaphone window, not downloads since the migration."
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
