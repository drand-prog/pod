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
from datetime import datetime
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


def read_csv_rows(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def build_megaphone_daily():
    # Megaphone's export is sorted by Downloads descending, not by date — sort
    # chronologically or the daily trend chart plots dates out of order.
    rows = read_csv_rows(RAW / "Megaphone_podcast-downloads-performance-2026-01-01-2026-07-15.csv")
    out = [
        {
            "date": parse_date(r["Date"]),
            "downloads": int(r["Downloads"]),
            "reach": int(r["Download reach"]),
        }
        for r in rows
    ]
    out.sort(key=lambda d: d["date"])
    return out


def build_megaphone_technology():
    rows = read_csv_rows(RAW / "Technology_Performance__2026-01-01_-_2026-07-16_.csv")
    out = []
    for r in rows:
        app = r["APPLICATION"].strip()
        downloads = int(r["DOWNLOADS"].replace(",", ""))
        pct = float(r["% OF TOTAL"].strip().rstrip("%")) / 100.0
        out.append({"app": app, "downloads": downloads, "pct": pct})
    return out


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
            "completionPct": float(r["Completion rate (%)"]),
            "publishDate": parse_date(r["Publish date"]),
        }
        for r in rows
    ]


def build_spotify_retention():
    rows = read_csv_rows(RAW / "Spotify_TheMeltingPod_WeekOverWeekRetention_1-1-2026--7-15-2026.csv")
    out = [
        {"weekStart": parse_date(r["Week starting"]), "retentionPct": round(float(r["Retention rate (%)"]), 2)}
        for r in rows
    ]
    out.sort(key=lambda d: d["weekStart"])
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
    sum_plays = None
    for r in rows[header_idx + 1 :]:
        if str(r[0]).startswith("Sum of listed"):
            sum_plays = r[6]
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
    return {"episodes": episodes, "sumOfListedPlays": sum_plays}


def build_apple_top_cities(wb):
    rows = sheet_rows(wb, "Apple Top Cities")
    header_idx = next(i for i, r in enumerate(rows) if r[0] == "City")
    all_cities_total = None
    cities = []
    truncated = False
    for r in rows[header_idx + 1 :]:
        if r[0] == "All Cities":
            all_cities_total = r[1]
            continue
        if str(r[0]).startswith("…") or str(r[0]).startswith("..."):
            truncated = True
            continue
        cities.append({"city": r[0], "listeners": r[1]})
    return {"allCitiesTotal": all_cities_total, "cities": cities, "listTruncated": truncated}


def build_youtube_videos(wb):
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
                "subscribers": r[6],
                "subscribersPct": r[7],
                "impressions": r[8],
                "ctr": r[9],
            }
        )
    return {
        "videos": videos,
        "sumOfRows": {"views": sum_row[2], "watchTimeHours": sum_row[4], "subscribers": sum_row[6], "impressions": sum_row[8]} if sum_row else None,
        "channelTotal": {"views": total_row[2], "watchTimeHours": total_row[4], "subscribers": total_row[6], "impressions": total_row[8], "ctr": total_row[9]} if total_row else None,
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

    data = {
        "generatedAt": datetime.utcnow().strftime("%Y-%m-%d"),
        "periods": {
            "spotifyApple": "All-time",
            "youtube": "Jan 1 – Jul 14, 2026 (195 days)",
            "megaphoneDaily": "Jan 1 – Jul 15, 2026",
            "megaphoneAppReport": "Jan 1 – Jul 16, 2026",
        },
        "platformSummary": build_platform_summary(wb),
        "megaphone": {
            "daily": build_megaphone_daily(),
            "byApp": build_megaphone_technology(),
            "appReportTotal": 9632,
            "dailyTabTotal": 9585,
            "otherResidual": {
                "total": 9632,
                "apple": 6200,
                "spotify": 1003,
                "other": 2429,
            },
        },
        "spotify": {
            "streams": build_spotify_streams(),
            "engagement": build_spotify_engagement(),
            "completion": build_spotify_completion(),
            "wowRetention": build_spotify_retention(),
        },
        "apple": {
            "overview": {
                "followers": 487,
                "listeners": 587,
                "engagedListeners": 409,
                "playsHeader": 10400,
                "timeListenedHours": 1217,
                "timeListenedFollowing": 930,
                "timeListenedNotFollowing": 287,
            },
            **build_apple_episodes(wb),
            "topCities": build_apple_top_cities(wb),
        },
        "youtube": {
            "overview": {
                "views": 7543,
                "watchTimeHours": 241.5,
                "subscribersNet": 178,
                "impressions": 17427,
                "ctr": 0.042,
                "avgViewDuration": "4:05",
            },
            **build_youtube_videos(wb),
            "trafficSources": build_youtube_traffic(wb),
            "funnel": build_youtube_funnel(wb),
        },
        "caveats": [
            "Different units, not interchangeable: Megaphone counts downloads, Apple/Spotify count plays, Spotify separately reports streams, YouTube counts views. These are not summed or directly compared as if identical.",
            "Different time windows: Spotify & Apple figures are all-time; YouTube is Jan 1–Jul 14, 2026; Megaphone covers Jan 1–Jul 15/16, 2026. Periods are labeled throughout — no shared window is implied.",
            "Plays ≠ downloads: Apple plays (~10,400) exceed Megaphone's total downloads (9,585–9,632) because a play is a playback event (repeats included) while a download is one de-duplicated file request.",
            "“Other platforms” residual uses Megaphone's own per-app DOWNLOAD figures, not native play counts: Megaphone total (9,632) − Apple downloads (6,200) − Spotify downloads (1,003) = 2,429. This is the only valid basis for that residual.",
            "Snapshots vs. trends: only Megaphone daily downloads, Spotify daily streams/engagement, and Spotify WoW retention are true time series. Apple and YouTube figures are single point-in-time snapshots, shown as KPI cards / rank bars — not fabricated trend lines.",
            "Two Megaphone totals differ slightly: the app (Technology) report totals 9,632 through Jul 16, while the daily-downloads tab totals 9,585 through Jul 15 — one extra day plus rounding. The app-report total is used for the “Other” residual so it ties out internally.",
        ],
        "sources": [
            "Apple Podcasts Connect — Listener analytics: podcasters.apple.com/support/5392-listener-analytics",
            "Spotify — “A New Standard for Podcast Plays” (11 Jun 2026): newsroom.spotify.com",
            "YouTube Help — Impressions & CTR / key metrics: support.google.com/youtube",
            "IAB Tech Lab Podcast Measurement — Megaphone certification (v2.2): compliance.iabtechnologylab.com",
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
