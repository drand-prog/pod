#!/usr/bin/env python3
"""
Pulls the latest Megaphone downloads and technology-performance reports and
writes them into data/raw/, in the format scripts/build_data.py expects.

--- What's confirmed vs. not (read before running this in production) ---

CONFIRMED (against Megaphone's own open-source client,
github.com/theatlantic/megaphone, and its published API auth format):
  - Base URL for Megaphone's content API: https://cms.megaphone.fm/api
  - Auth header: Authorization: Token token="<token>"
  - GET /networks/{network_id}/podcasts -> list of {id, title, ...}

NOT YET CONFIRMED: the reporting/analytics endpoints that back the
"Podcast Performance" / "Technology Performance" dashboards (i.e. the ones
that produced the CSVs this script is meant to replace) are documented at
https://developers.megaphone.fm/, which was unreachable from the environment
this script was written in. fetch_downloads_report() and
fetch_technology_report() below are stubs for exactly that reason — filling
them in requires either:
  1. Checking developers.megaphone.fm's Reporting section for the exact
     endpoint path, query params (date range, podcast id), and response
     shape, then implementing the two functions below to match; or
  2. Asking Megaphone support to enable the Metrics Export Service (writes
     hourly report files to an S3 bucket you control — see
     https://support.megaphone.fm/en/articles/2678649-metrics-export-service)
     and rewriting this script to read from S3 instead of calling a REST
     endpoint.

Until one of those is done, running this script will fail loudly with
NotImplementedError rather than silently write wrong data.

Usage:
    MEGAPHONE_API_TOKEN=... MEGAPHONE_NETWORK_ID=... python3 scripts/fetch_megaphone.py
    (MEGAPHONE_PODCAST_ID is optional — only needed if the network has more
    than one podcast.)
"""
import csv
import os
import re
from datetime import date, datetime
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "raw"
BASE_URL = "https://cms.megaphone.fm/api"


def client(token):
    session = requests.Session()
    session.headers["Authorization"] = f'Token token="{token}"'
    return session


def get_podcast_id(session, network_id, podcast_id=None):
    if podcast_id:
        return podcast_id
    resp = session.get(f"{BASE_URL}/networks/{network_id}/podcasts")
    resp.raise_for_status()
    podcasts = resp.json()
    if len(podcasts) != 1:
        raise SystemExit(
            f"Network {network_id} has {len(podcasts)} podcasts; set MEGAPHONE_PODCAST_ID "
            f"explicitly. Found: {[(p['id'], p['title']) for p in podcasts]}"
        )
    return podcasts[0]["id"]


def fetch_downloads_report(session, podcast_id, start, end):
    """Daily downloads + download reach for [start, end].

    Expected return: list of {"date": "YYYY-MM-DD", "downloads": int, "reach": int}.
    See the module docstring — the endpoint here is not yet confirmed.
    """
    raise NotImplementedError(
        "Confirm the downloads-report endpoint at https://developers.megaphone.fm/ "
        "(or switch to the Metrics Export Service), then implement this call. "
        "See the module docstring for what's already confirmed."
    )


def fetch_technology_report(session, podcast_id, start, end):
    """Downloads broken out by listening app, for [start, end].

    Expected return: list of {"app": str, "downloads": int, "pct": float in [0,1]}.
    See the module docstring — the endpoint here is not yet confirmed.
    """
    raise NotImplementedError(
        "Confirm the technology-performance endpoint at https://developers.megaphone.fm/ "
        "(or switch to the Metrics Export Service), then implement this call. "
        "See the module docstring for what's already confirmed."
    )


def clear_old_reports(prefix_pattern):
    """Remove previously fetched files matching prefix_pattern so build_data.py's
    glob (which errors on multiple matches) doesn't choke on stale date ranges."""
    for old in RAW.glob(prefix_pattern):
        old.unlink()


def write_downloads_csv(rows, start, end):
    clear_old_reports("Megaphone_podcast-downloads-performance-*.csv")
    path = RAW / f"Megaphone_podcast-downloads-performance-{start}-{end}.csv"
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Date", "Downloads", "Download reach"])
        for r in rows:
            w.writerow([r["date"], r["downloads"], r["reach"]])
    return path


def write_technology_csv(rows, start, end):
    clear_old_reports("Technology_Performance*.csv")
    path = RAW / f"Technology_Performance__{start}_-_{end}_.csv"
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["APPLICATION", "DOWNLOADS", "% OF TOTAL"])
        for r in rows:
            w.writerow([r["app"], r["downloads"], f'{round(r["pct"] * 100)}%'])
    return path


def main():
    token = os.environ["MEGAPHONE_API_TOKEN"]
    network_id = os.environ["MEGAPHONE_NETWORK_ID"]
    podcast_id = os.environ.get("MEGAPHONE_PODCAST_ID")
    # Keep the existing history's start date so the refreshed report still
    # covers the show's full run, not just the most recent window.
    start = os.environ.get("MEGAPHONE_START_DATE", "2026-01-01")
    end = date.today().isoformat()

    session = client(token)
    podcast_id = get_podcast_id(session, network_id, podcast_id)

    downloads = fetch_downloads_report(session, podcast_id, start, end)
    downloads_path = write_downloads_csv(downloads, start, end)

    technology = fetch_technology_report(session, podcast_id, start, end)
    technology_path = write_technology_csv(technology, start, end)

    print(f"Wrote {len(downloads)} daily rows to {downloads_path.name}")
    print(f"Wrote {len(technology)} app rows to {technology_path.name}")


if __name__ == "__main__":
    main()
