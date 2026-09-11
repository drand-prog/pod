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
    for fmt in ("%Y-%m-%d", "%m/%d/%Y"):
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


def build_megaphone_technology():
    path = find_one("Technology_Performance*.csv")
    rows = read_csv_rows(path)
    out = []
    for r in rows:
        app = r["APPLICATION"].strip()
        downloads = int(r["DOWNLOADS"].replace(",", ""))
        pct = float(r["% OF TOTAL"].strip().rstrip("%")) / 100.0
        out.append({"app": app, "downloads": downloads, "pct": pct})
    return out, path


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
    megaphone_tech, megaphone_tech_path = build_megaphone_technology()

    daily_start, daily_end = megaphone_daily[0]["date"], megaphone_daily[-1]["date"]
    tech_range = extract_date_range_from_filename(megaphone_tech_path)
    tech_start, tech_end = tech_range if tech_range else (daily_start, daily_end)

    daily_tab_total = sum(d["downloads"] for d in megaphone_daily if d["downloads"] is not None)
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

    periods = {
        "spotifyApple": "All-time",
        # Traffic Sources/Videos/Overview refreshed Sep 2026 (no exact day-count
        # given by the export, unlike the original screenshot); the Funnel tab
        # is not part of this refresh and still shows its own Jan 1-Jul 14 note.
        "youtube": "Jan 1 – Sep 2026",
        "megaphoneDaily": fmt_period(daily_start, daily_end),
        "megaphoneAppReport": fmt_period(tech_start, tech_end),
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

    data = {
        "generatedAt": datetime.utcnow().strftime("%Y-%m-%d"),
        "periods": periods,
        "platformSummary": platform_summary,
        "megaphone": {
            "daily": megaphone_daily,
            "byApp": megaphone_tech,
            "appReportTotal": app_report_total,
            "dailyTabTotal": daily_tab_total,
            "otherResidual": {
                "total": app_report_total,
                "apple": apple_downloads,
                "spotify": spotify_downloads,
                "other": other_downloads,
            },
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
            **build_apple_episodes(wb),
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
                    else f", because the app report hasn't been refreshed as recently as the daily downloads."
                )
                + " The app-report total is used for the “Other” residual, which reflects the app report's period, not the daily tab's."
            ),
            f"YouTube's country breakdown only lists {len(youtube_geo['countries'])} countries totaling {fmt(sum(c['views'] for c in youtube_geo['countries']))} views, against a channel total of {fmt(youtube_geo['total']['views'] if youtube_geo['total'] else 0)} — YouTube Studio doesn't break out the long tail below its per-country reporting threshold, so the list is shown as-is rather than padded to 100%.",
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
