#!/usr/bin/env python3
"""
Pulls new Megaphone "Delivery Export" files from S3 and folds them into the
historical CSVs in data/raw/ that scripts/build_data.py reads.

--- Background ---

Megaphone has no REST endpoint for downloads/technology-performance reports.
The only supported way to pull this data programmatically is the Delivery &
Impressions Export service: Megaphone pushes gzipped, newline-delimited JSON
files to an S3 bucket you own, one per day, named
    delivery-v2-day-YYYY-MM-DD.json.gz
updated hourly through the day and finalized (won't change again) once a
sibling marker key
    _finalized_delivery_YYYY-MM-DD_<timestamp>
appears — confirmed in Megaphone's "Delivery & Impressions Export V2" help
article (May 2026 revision). This script only reads the delivery file, not
the impressions file (that one's ad-verification data, not needed here).

Per that article, each row's relevant fields are:
    delivery_type   "download" | "play"  (download = all platforms, matches
                     what the existing CSVs measure; play = Spotify only)
    normalized_user_agent   e.g. "Spotify", "Apple Podcasts", "Chrome", ...
    created_at       ISO 8601 timestamp (UTC) of the event
    ip               a HASHED per-listener identifier (not a raw IP)
Megaphone's own guidance: "Downloads: count all rows where delivery_type ==
'download'" — and blacklist/bot filtering is already applied server-side, so
no de-dup logic is needed on our end for the *download* count itself.

The one figure the article does NOT define is "download reach" (a column in
the existing Megaphone Downloads CSV, distinct from raw downloads). This
script approximates it as the count of distinct hashed `ip` values among a
day's download rows — a reasonable read of "reach" as unique listeners
reached, consistent with reach being <= downloads in the existing data — but
it's an inference, not something Megaphone's docs state outright. Flagging
this rather than presenting it as confirmed.

--- Why this appends instead of replacing ---

The export has NO BACKFILL: Megaphone only pushes data from the date the S3
bucket was wired up, forward. So this script can never regenerate the full
Jan-onward history from S3 — it can only extend what's already in
data/raw/. Each run:
  1. Reads the existing daily-downloads CSV and technology-performance CSV
     (whatever's currently in data/raw/, from the very first manual export
     onward) to find the last date already covered.
  2. Lists finalized delivery-export files in S3 for dates after that.
  3. Aggregates each new day's downloads/reach/by-app counts and merges them
     into the existing data (new dates appended; an existing date, if ever
     reprocessed, is overwritten rather than double-counted).
  4. Rewrites both CSVs under a new date-ranged filename, matching the
     existing naming convention, and removes the old ones.

Required environment variables (set as GitHub Actions secrets):
    AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_REGION  - read-only
        credentials scoped to the export bucket (see README's Automation
        section for the IAM setup)
    MEGAPHONE_S3_BUCKET   - the bucket Megaphone was configured to write to
Optional:
    MEGAPHONE_S3_PREFIX   - key prefix, if the exports aren't at the bucket root
"""
import csv
import gzip
import io
import json
import os
import re
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from pathlib import Path

import boto3

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "raw"

DAILY_GLOB = "Megaphone_podcast-downloads-performance-*.csv"
TECH_GLOB = "Technology_Performance*.csv"


def find_one(pattern):
    matches = sorted(RAW.glob(pattern))
    if len(matches) != 1:
        raise RuntimeError(f"Expected exactly one file matching {pattern!r} in {RAW}, found {matches}")
    return matches[0]


def extract_date_range_from_filename(path):
    found = re.findall(r"\d{4}-\d{2}-\d{2}", path.stem)
    if len(found) != 2:
        raise RuntimeError(f"Couldn't find a start/end date in filename {path.name!r}")
    return found  # (start, end)


def read_existing_daily():
    path = find_one(DAILY_GLOB)
    _, end = extract_date_range_from_filename(path)
    by_date = {}
    with open(path, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            by_date[row["Date"]] = {"downloads": int(row["Downloads"]), "reach": int(row["Download reach"])}
    return by_date, end, path


def read_existing_technology():
    path = find_one(TECH_GLOB)
    start, end = extract_date_range_from_filename(path)
    by_app = {}
    with open(path, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            by_app[row["APPLICATION"].strip()] = int(row["DOWNLOADS"].replace(",", ""))
    return by_app, start, end, path


def s3_client():
    return boto3.client("s3", region_name=os.environ.get("AWS_REGION"))


def list_finalized_dates(s3, bucket, prefix, since_date):
    """Dates (as 'YYYY-MM-DD' strings, > since_date) for which Megaphone has
    written the finalization marker, meaning that day's delivery file is done
    changing and safe to aggregate."""
    marker_prefix = f"{prefix}_finalized_delivery_"
    paginator = s3.get_paginator("list_objects_v2")
    dates = set()
    for page in paginator.paginate(Bucket=bucket, Prefix=marker_prefix):
        for obj in page.get("Contents", []):
            m = re.search(r"_finalized_delivery_(\d{4}-\d{2}-\d{2})_", obj["Key"])
            if m and m.group(1) > since_date:
                dates.add(m.group(1))
    return sorted(dates)


def fetch_delivery_rows(s3, bucket, prefix, date_str):
    key = f"{prefix}delivery-v2-day-{date_str}.json.gz"
    obj = s3.get_object(Bucket=bucket, Key=key)
    raw = gzip.decompress(obj["Body"].read())
    text = raw.decode("utf-8").strip()
    if not text:
        return []
    if text[0] == "[":
        return json.loads(text)
    # newline-delimited JSON — one event per line
    return [json.loads(line) for line in text.splitlines() if line.strip()]


def aggregate_day(rows):
    downloads = 0
    reach_ips = set()
    by_app = Counter()
    for r in rows:
        if r.get("delivery_type") != "download":
            continue
        downloads += 1
        ip = r.get("ip")
        if ip:
            reach_ips.add(ip)
        app = r.get("normalized_user_agent") or "Unknown"
        by_app[app] += 1
    return downloads, len(reach_ips), by_app


def write_daily_csv(by_date, start, end):
    for old in RAW.glob(DAILY_GLOB):
        old.unlink()
    path = RAW / f"Megaphone_podcast-downloads-performance-{start}-{end}.csv"
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Date", "Downloads", "Download reach"])
        for date in sorted(by_date):
            w.writerow([date, by_date[date]["downloads"], by_date[date]["reach"]])
    return path


def write_technology_csv(by_app, start, end):
    for old in RAW.glob(TECH_GLOB):
        old.unlink()
    path = RAW / f"Technology_Performance__{start}_-_{end}_.csv"
    total = sum(by_app.values())
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["APPLICATION", "DOWNLOADS", "% OF TOTAL"])
        for app, downloads in sorted(by_app.items(), key=lambda kv: -kv[1]):
            pct = round(downloads / total * 100) if total else 0
            w.writerow([app, downloads, f"{pct}%"])
    return path


def main():
    bucket = os.environ["MEGAPHONE_S3_BUCKET"]
    prefix = os.environ.get("MEGAPHONE_S3_PREFIX", "")
    if prefix and not prefix.endswith("/"):
        prefix += "/"

    daily_by_date, daily_end, daily_path = read_existing_daily()
    tech_by_app, tech_start, tech_end, tech_path = read_existing_technology()
    # Both reports are folded from the same underlying event log, so advance
    # them together from whichever is further behind.
    since_date = min(daily_end, tech_end)

    s3 = s3_client()
    new_dates = list_finalized_dates(s3, bucket, prefix, since_date)
    if not new_dates:
        print(f"No finalized delivery-export files newer than {since_date}. Nothing to do.")
        return

    for date_str in new_dates:
        rows = fetch_delivery_rows(s3, bucket, prefix, date_str)
        downloads, reach, by_app = aggregate_day(rows)
        daily_by_date[date_str] = {"downloads": downloads, "reach": reach}
        for app, count in by_app.items():
            tech_by_app[app] = tech_by_app.get(app, 0) + count
        print(f"{date_str}: +{downloads} downloads, {reach} unique (reach), {len(by_app)} apps")

    new_end = max(new_dates)
    daily_start = min(daily_by_date)
    written_daily = write_daily_csv(daily_by_date, daily_start, new_end)
    written_tech = write_technology_csv(tech_by_app, tech_start, new_end)

    print(f"Merged {len(new_dates)} new day(s) ({new_dates[0]}..{new_dates[-1]})")
    print(f"Wrote {written_daily.name}")
    print(f"Wrote {written_tech.name}")


if __name__ == "__main__":
    main()
