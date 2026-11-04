"""
2W news pipeline (CSV storage): fetch by COUNTRY, match BRANDS locally, export to Excel.

    pip install eventregistry python-dotenv openpyxl
    .env: NEWSAPI_AI_KEY=your_key

    python pipeline.py fetch  [--countries India,Japan] [--days 7] [--max-per-country 300]
    python pipeline.py match  [--all]      # no API cost; re-run whenever brands.csv changes
    python pipeline.py export [--brand "TVS Motor"] [--out news_2w.xlsx]
    python pipeline.py stats

Files:
    brands.csv         brand dictionary (you edit this)
    data/articles.csv  every fetched article + matched brands (written by the script)
    data/runs.json     last fetch date per country (for incremental fetching)
"""
import argparse, csv, json, os, re, time
from datetime import datetime, timedelta, timezone

BRANDS_CSV = "brands.csv"
DATA_DIR = "data"
ARTICLES_CSV = os.path.join(DATA_DIR, "articles.csv")
RUNS_JSON = os.path.join(DATA_DIR, "runs.json")

FIELDS = ["uri", "date", "title", "source", "source_uri", "url", "lang", "country",
          "sentiment", "body", "matched", "brands", "aliases"]

# ---------------------------------------------------------------------------
# Country config: where to search and in which language(s).
# `terms` = 2W topic words sent to the API. Keep each list SHORT
# (your plan caps ~15 keyword-words per query).
# ---------------------------------------------------------------------------
COUNTRIES = {
    "India":     {"lang": "eng", "terms": ["two-wheeler", "motorcycle", "scooter", "electric scooter"]},
    "Indonesia": {"lang": "ind", "terms": ["sepeda motor", "motor listrik", "skuter"]},
    "Vietnam":   {"lang": "vie", "terms": ["xe máy", "xe mô tô", "xe máy điện"]},
    "Japan":     {"lang": "jpn", "terms": ["オートバイ", "二輪車", "スクーター"]},
    "Brazil":    {"lang": "por", "terms": ["motocicleta", "moto elétrica", "scooter"]},
    "USA":       {"lang": "eng", "terms": ["motorcycle", "scooter", "electric motorcycle"]},
}

# Used locally to decide whether an ambiguous brand (Honda, Suzuki...) is talking 2W.
CONTEXT_RE = re.compile(
    r"two[- ]?wheeler|\b2W\b|\be2W\b|motorcycle|motorbike|scooter|moped|"
    r"sepeda motor|xe máy|オートバイ|二輪|motocicleta|\bmoto\b", re.I)
PROXIMITY = 200   # chars around an ambiguous alias that must contain a 2W word


# ------------------------------ CSV storage --------------------------------
def load_articles():
    """-> dict uri -> row (insertion ordered)."""
    if not os.path.exists(ARTICLES_CSV):
        return {}
    with open(ARTICLES_CSV, newline="", encoding="utf-8-sig") as f:
        return {r["uri"]: {k: r.get(k, "") for k in FIELDS} for r in csv.DictReader(f)}


def save_articles(arts):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(ARTICLES_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(arts.values())


def load_runs():
    if os.path.exists(RUNS_JSON):
        with open(RUNS_JSON, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_runs(runs):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(RUNS_JSON, "w", encoding="utf-8") as f:
        json.dump(runs, f, indent=2)


# ------------------------------- FETCH -------------------------------------
def fetch(countries, days, max_per_country):
    from dotenv import load_dotenv
    from eventregistry import (EventRegistry, QueryArticlesIter, QueryItems,
                               BaseQuery, ComplexArticleQuery)
    load_dotenv()
    key = os.environ.get("NEWSAPI_AI_KEY")
    if not key:
        raise SystemExit("NEWSAPI_AI_KEY not found (check .env).")
    er = EventRegistry(apiKey=key, allowUseOfArchive=False)

    arts, runs = load_articles(), load_runs()
    today = datetime.now(timezone.utc).date()

    for country in countries:
        cfg = COUNTRIES[country]
        # incremental: start from last successful fetch (minus 1 day overlap)
        if country in runs:
            start = datetime.fromisoformat(runs[country]).date() - timedelta(days=1)
        else:
            start = today - timedelta(days=days)
        try:
            q = BaseQuery(
                keyword=QueryItems.OR(cfg["terms"]),
                keywordLoc="title,body",
                sourceLocationUri=er.getLocationUri(country),
                lang=cfg["lang"],
                dateStart=start.isoformat(), dateEnd=today.isoformat())
            it = QueryArticlesIter.initWithComplexQuery(ComplexArticleQuery(
                query=q, dataType=["news", "blog"], isDuplicateFilter="skipDuplicates"))
            n = new = 0
            for a in it.execQuery(er, sortBy="date", maxItems=max_per_country):
                n += 1
                uri = a.get("uri")
                if not uri or uri in arts:
                    continue
                arts[uri] = {
                    "uri": uri, "date": a.get("dateTime") or a.get("date"),
                    "title": a.get("title"),
                    "source": (a.get("source") or {}).get("title"),
                    "source_uri": (a.get("source") or {}).get("uri"),
                    "url": a.get("url"), "lang": a.get("lang"), "country": country,
                    "sentiment": a.get("sentiment"), "body": (a.get("body") or "")[:4000],
                    "matched": "", "brands": "", "aliases": ""}
                new += 1
            runs[country] = today.isoformat()
            save_articles(arts); save_runs(runs)   # save after every country
            print(f"{country:<12} fetched {n:>4}  new {new:>4}  (since {start})")
        except Exception as e:
            print(f"{country:<12} !! failed: {e}")
        time.sleep(0.5)


# ------------------------------- MATCH -------------------------------------
def load_brands(path=BRANDS_CSV):
    """Returns (regex over ALL aliases, alias_lower -> (brand, ambiguous))."""
    alias_map = {}
    with open(path, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            for al in filter(None, (x.strip() for x in r["aliases"].split("|"))):
                alias_map[al.lower()] = (r["brand"], r["ambiguous"].strip() == "1")
    parts = sorted(alias_map, key=len, reverse=True)   # longest alias first
    rx = re.compile(r"(?<!\w)(" + "|".join(re.escape(p) for p in parts) + r")(?!\w)", re.I)
    return rx, alias_map


def match_text(text, rx, alias_map):
    """-> {brand: alias} for brands genuinely mentioned in text."""
    found = {}
    for m in rx.finditer(text):
        brand, ambiguous = alias_map[m.group(1).lower()]
        if brand in found:
            continue
        if ambiguous:
            lo, hi = max(0, m.start() - PROXIMITY), m.end() + PROXIMITY
            if not CONTEXT_RE.search(text[lo:hi]):
                continue          # "Honda" near no 2W words -> probably a car story
        found[brand] = m.group(1)
    return found


def match(rematch_all=False):
    rx, alias_map = load_brands()
    arts = load_articles()
    todo = [r for r in arts.values() if rematch_all or r["matched"] != "1"]
    hits = 0
    for r in todo:
        found = match_text(f"{r['title'] or ''}\n{r['body'] or ''}", rx, alias_map)
        r["brands"] = "|".join(sorted(found))
        r["aliases"] = "|".join(found[b] for b in sorted(found))
        r["matched"] = "1"
        hits += len(found)
    save_articles(arts)
    print(f"Matched {len(todo)} articles -> {hits} brand links")


# ------------------------------- EXPORT / STATS ----------------------------
def export(brand, out):
    arts = load_articles()
    rows = []
    for r in arts.values():                      # one output row per (article, brand)
        for b in filter(None, r["brands"].split("|")):
            if brand and b != brand:
                continue
            rows.append([b, r["date"], r["country"], r["title"], r["source"], r["url"], r["sentiment"]])
    rows.sort(key=lambda x: (x[0], str(x[1] or "")), reverse=False)
    # newest first within each brand
    rows.sort(key=lambda x: str(x[1] or ""), reverse=True)
    rows.sort(key=lambda x: x[0])
    header = ["brand", "date", "country", "title", "source", "url", "sentiment"]

    if out.lower().endswith(".xlsx"):
        from openpyxl import Workbook
        from openpyxl.styles import Font
        wb = Workbook()
        ws = wb.active
        ws.title = "News"
        ws.append(header)
        for r in rows:
            ws.append(r)
        for c in ws[1]:
            c.font = Font(bold=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col, width in zip("ABCDEFG", [20, 20, 12, 70, 25, 50, 10]):
            ws.column_dimensions[col].width = width
        wb.save(out)
    else:
        with open(out, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(header)
            w.writerows(rows)
    print(f"Wrote {len(rows)} rows -> {out}")


def stats():
    arts = load_articles()
    print("articles:", len(arts))
    by_brand, by_country = {}, {}
    for r in arts.values():
        by_country[r["country"]] = by_country.get(r["country"], 0) + 1
        for b in filter(None, r["brands"].split("|")):
            by_brand[b] = by_brand.get(b, 0) + 1
    print("\nbrand mentions:")
    for b, n in sorted(by_brand.items(), key=lambda x: -x[1]):
        print(f"  {b:<20}{n}")
    print("\narticles by country:")
    for c, n in sorted(by_country.items(), key=lambda x: -x[1]):
        print(f"  {c:<20}{n}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    f = sub.add_parser("fetch")
    f.add_argument("--countries", default=",".join(COUNTRIES))
    f.add_argument("--days", type=int, default=7)
    f.add_argument("--max-per-country", type=int, default=300)
    m = sub.add_parser("match"); m.add_argument("--all", action="store_true")
    e = sub.add_parser("export")
    e.add_argument("--brand"); e.add_argument("--out", default="news_2w.xlsx")
    sub.add_parser("stats")
    a = ap.parse_args()
    if a.cmd == "fetch":
        fetch([c.strip() for c in a.countries.split(",")], a.days, a.max_per_country)
    elif a.cmd == "match": match(a.all)
    elif a.cmd == "export": export(a.brand, a.out)
    else: stats()