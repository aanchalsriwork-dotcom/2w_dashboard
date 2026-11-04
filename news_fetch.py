"""
2W industry news fetcher, brand-wise, using newsapi.ai (Event Registry) Python SDK.

Setup:
    pip install eventregistry pandas python-dotenv
    .env file (same folder):  NEWSAPI_AI_KEY=your_key

Run:
    python twowheeler_news_agent.py --days 7 --max-per-brand 20
"""

import argparse
import json
import os
import re
import time
from datetime import datetime, timedelta, timezone

import pandas as pd
from dotenv import load_dotenv
from eventregistry import (
    BaseQuery,
    ComplexArticleQuery,
    EventRegistry,
    QueryArticlesIter,
    QueryItems,
)

load_dotenv()

# Brand search terms sent to the API (keep each brand to ~4-6 words total:
# the API counts individual words against your plan's keyword limit).
BRANDS = {
    "Hero MotoCorp":    ["Hero MotoCorp", "Hero Electric"],
    "Honda 2W":         ["Honda Motorcycle", "HMSI"],
    "TVS Motor":        ["TVS Motor", "TVS Apache"],
    "Bajaj Auto":       ["Bajaj Auto", "Bajaj Chetak"],
    "Royal Enfield":    ["Royal Enfield", "Eicher Motors"],
    "Yamaha Motor":     ["Yamaha Motor", "Yamaha motorcycle"],
    "Suzuki 2W":        ["Suzuki Motorcycle", "Suzuki Access"],
    "Harley-Davidson":  ["Harley-Davidson", "LiveWire"],
    "Ola Electric":     ["Ola Electric"],
    "Ather Energy":     ["Ather Energy"],
}

# 2W context is checked locally in Python, NOT sent to the API.
CONTEXT_RE = re.compile(
    r"two[- ]?wheeler|\b2W\b|\be2W\b|motorcycle|motorbike|scooter|moped", re.I
)

LANG = "eng"


def build_query(search_terms, date_start, date_end):
    q = BaseQuery(
        keyword=QueryItems.OR(search_terms),
        keywordLoc="title,body",
        dateStart=date_start,
        dateEnd=date_end,
        lang=LANG,
    )
    complex_q = ComplexArticleQuery(
        query=q,
        dataType=["news", "blog"],
        isDuplicateFilter="skipDuplicates",
        startSourceRankPercentile=0,
        endSourceRankPercentile=50,
    )
    return QueryArticlesIter.initWithComplexQuery(complex_q)


def fetch_brand_articles(er, brand, terms, days, max_items):
    end = datetime.now(timezone.utc).date()
    start = end - timedelta(days=days)
    q = build_query(terms, start.isoformat(), end.isoformat())

    rows = []
    # fetch extra, since some articles are dropped by the 2W context filter
    for art in q.execQuery(er, sortBy="date", maxItems=max_items * 3):
        text = f"{art.get('title') or ''} {art.get('body') or ''}"
        if not CONTEXT_RE.search(text):
            continue
        rows.append({
            "brand": brand,
            "date": art.get("dateTime") or art.get("date"),
            "title": art.get("title"),
            "source": (art.get("source") or {}).get("title"),
            "source_uri": (art.get("source") or {}).get("uri"),
            "url": art.get("url"),
            "lang": art.get("lang"),
            "sentiment": art.get("sentiment"),
            "body_snippet": (art.get("body") or "")[:500],
            "article_uri": art.get("uri"),
        })
        if len(rows) >= max_items:
            break
    return rows


def run(days, max_per_brand, out_prefix):
    api_key = os.environ.get("NEWSAPI_AI_KEY")
    if not api_key:
        raise SystemExit("NEWSAPI_AI_KEY not found. Check your .env file.")

    er = EventRegistry(apiKey=api_key, allowUseOfArchive=False)

    all_rows = []
    for brand, terms in BRANDS.items():
        try:
            rows = fetch_brand_articles(er, brand, terms, days, max_per_brand)
            print(f"{brand:<18} -> {len(rows)} articles")
            all_rows.extend(rows)
        except Exception as e:
            print(f"{brand:<18} !! failed: {e}")
        time.sleep(0.5)

    if not all_rows:
        print("No articles fetched.")
        return

    df = pd.DataFrame(all_rows)
    df = df.drop_duplicates(subset=["brand", "url"]).sort_values(
        ["brand", "date"], ascending=[True, False]
    )

    ts = datetime.now().strftime("%Y%m%d_%H%M")
    csv_path, json_path = f"{out_prefix}_{ts}.csv", f"{out_prefix}_{ts}.json"
    df.to_csv(csv_path, index=False)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(df.to_dict(orient="records"), f, ensure_ascii=False, indent=2)
    print(f"\nSaved {len(df)} rows -> {csv_path}, {json_path}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=7)
    ap.add_argument("--max-per-brand", type=int, default=20)
    ap.add_argument("--out", default="news_2w")
    args = ap.parse_args()
    run(args.days, args.max_per_brand, args.out)