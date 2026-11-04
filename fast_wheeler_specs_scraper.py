#!/usr/bin/env python3
"""
Fasterwheeler detail-page scraper
Reads the URL column of 
nd adds:
    Brand, Model, Release Date, Status, Features, full Specs (one column per spec)

Why it is fast: the same model appears in several countries (same ID at the end of
the URL, e.g. ...-10232). Specs do not change by country, so each model ID is fetched
ONCE and the result is merged back onto every country row.

pip install requests beautifulsoup4 pandas openpyxl

Usage
  python fasterwheeler_specs_scraper.py                         # full run
  python fasterwheeler_specs_scraper.py --limit 20              # first 20 model IDs (test)
  python fasterwheeler_specs_scraper.py --test <detail-page-url>  # print one parsed page
The run is resumable: results are cached in specs_cache.jsonl, so re-running skips
models that were already fetched.
"""
import argparse
import json
import re
import threading
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup, NavigableString, Tag
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

INPUT_FILE = "2W_pricing_Data.xlsx"
OUTPUT_FILE = "2W_pricing_Data_with_specs.xlsx"
CACHE_FILE = Path("specs_cache.jsonl")
WORKERS = 3
DELAY = 0.5                       # seconds each worker waits after a request
GENERIC_OG = "og-site-imgs"       # site-wide placeholder image

session = requests.Session()
session.headers.update({"User-Agent": "Mozilla/5.0 (price-research script)",
                        "Accept-Language": "en-US,en;q=0.9"})
session.mount("https://", HTTPAdapter(pool_maxsize=WORKERS * 2, max_retries=Retry(
    total=3, backoff_factor=1.5, status_forcelist=[429, 500, 502, 503, 504])))

CORE_MAP = {"brand": "Brand", "model": "Model",
            "released date": "Release Date", "release date": "Release Date",
            "launch date": "Release Date", "launched": "Release Date",
            "status": "Status"}


# ───────────────────────── helpers ─────────────────────────
def norm(s):
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().lower()


def clean(s):
    s = re.sub(r"\s+", " ", s or "").strip()
    return re.sub(r"\s+([,.;:!?])", r"\1", s)


def model_id(url):
    m = re.search(r"-(\d+)/?$", str(url))
    return m.group(1) if m else None


def meta(soup, prop):
    tag = soup.find("meta", attrs={"property": prop}) or soup.find("meta", attrs={"name": prop})
    return tag.get("content") if tag and tag.get("content") else None


# ───────────────────────── parsing one detail page ─────────────────────────
def find_spec_table(soup):
    best = None
    for t in soup.find_all("table"):
        rows = t.find_all("tr")
        labels = set()
        for r in rows:
            c = r.find(["td", "th"])
            if c:
                labels.add(norm(c.get_text(" ", strip=True)))
        if "brand" in labels and "model" in labels:
            return t
        if best is None or len(rows) > len(best.find_all("tr")):
            best = t
    return best


def parse_table(table):
    """-> (core dict, [(section, label, value), ...])"""
    core, specs, section = {}, [], "General"
    if table is None:
        return core, specs
    for tr in table.find_all("tr"):
        cells = [clean(c.get_text(" ", strip=True)) for c in tr.find_all(["td", "th"])]
        if not cells:
            continue
        label, value = cells[0], " ".join(cells[1:]).strip()
        if not label and not value:
            continue
        if not value:                              # row with no value = section header
            section = label.title()
            continue
        key = norm(label)
        if section == "General" and key in CORE_MAP:
            core[CORE_MAP[key]] = value
            continue
        specs.append((section, label, value))
    return core, specs


def parse_description(soup, table):
    """Prose between the 'Specifications' heading and the table, plus bold lead-ins
    such as 'Digital Instrument Cluster:' which name the individual features."""
    h = soup.find(lambda t: t.name in ("h1", "h2", "h3", "h4")
                  and "specification" in t.get_text(" ", strip=True).lower())
    if h is None or table is None:
        return "", []
    skip = {id(x) for x in h.descendants}
    lines, buf, names = [], [], []

    def flush():
        t = clean(" ".join(buf))
        if t:
            lines.append(t)
        buf.clear()

    for el in h.next_elements:
        if el is table:
            break
        if id(el) in skip:
            continue
        if isinstance(el, NavigableString):
            if el.parent is not None and el.parent.name in ("script", "style"):
                continue
            if str(el).strip():
                buf.append(str(el))
        elif isinstance(el, Tag):
            if el.name in ("p", "div", "li", "ul", "ol", "br", "h4", "h5", "h6"):
                flush()
            if el.name in ("strong", "b"):
                t = clean(el.get_text(" ", strip=True))
                n = norm(t)
                if (t.endswith(":") and len(t) < 70 and "features" not in n
                        and not n.startswith(("what", "note", "key"))):
                    names.append(t.rstrip(":").strip())
    flush()
    return "\n".join(lines), names


def parse_detail(html):
    soup = BeautifulSoup(html, "html.parser")
    table = find_spec_table(soup)
    core, specs = parse_table(table)
    description, feature_names = parse_description(soup, table)
    text = soup.get_text(" ", strip=True)

    og = meta(soup, "og:image")
    year = re.search(r"\b(19|20)\d{2}\b", core.get("Release Date", ""))
    return {
        "Brand_Site": core.get("Brand"),
        "Model_Site": core.get("Model"),
        "Release_Date": core.get("Release Date"),
        "Release_Year": int(year.group(0)) if year else None,
        "Status": core.get("Status"),
        "Features": "; ".join(dict.fromkeys(feature_names)),
        "Description": description,
        "Photo_Large": og if og and GENERIC_OG not in og else None,
        "Page_Added": (re.search(r"Added on (\d{4}-\d{2}-\d{2})", text) or [None, None])[1],
        "Page_Updated": (re.search(r"Update on (\d{4}-\d{2}-\d{2})", text) or [None, None])[1],
        "specs": [list(s) for s in specs],
    }


# ───────────────────────── fetching (with cache) ─────────────────────────
cache_lock = threading.Lock()


def load_cache():
    cache = {}
    if CACHE_FILE.exists():
        for line in CACHE_FILE.read_text(encoding="utf-8").splitlines():
            try:
                rec = json.loads(line)
                cache[rec["model_id"]] = rec
            except (ValueError, KeyError):
                continue
    return cache


def fetch_model(mid, urls):
    """Try each country URL for this model until one works."""
    last_err = "no url"
    for url in urls:
        try:
            r = session.get(url, timeout=30)
            time.sleep(DELAY)
            if r.status_code != 200:
                last_err = f"HTTP {r.status_code}"
                continue
            rec = parse_detail(r.text)
            rec.update({"model_id": mid, "source_url": url})
            if rec["specs"] or rec["Brand_Site"]:
                return rec, None
            last_err = "page parsed but no spec table found"
        except requests.RequestException as e:
            last_err = e.__class__.__name__
    return None, last_err


# ───────────────────────── build output ─────────────────────────
def col_name(section, label):
    return label if section == "General" else f"{section} | {label}"


def build_outputs(df, cache):
    wide_rows, long_rows = {}, []
    for mid, rec in cache.items():
        row, seen = {}, {}
        for section, label, value in rec["specs"]:
            c = col_name(section, label)
            seen[c] = seen.get(c, 0) + 1
            if seen[c] > 1:
                c = f"{c} ({seen[c]})"
            row[c] = value
            long_rows.append((mid, rec.get("Brand_Site"), rec.get("Model_Site"), section, label, value))
        wide_rows[mid] = row

    spec_df = pd.DataFrame.from_dict(wide_rows, orient="index")
    freq = spec_df.notna().sum().sort_values(ascending=False)
    spec_df = spec_df[freq.index]                      # most common specs first
    spec_df.index.name = "ModelID"

    info = pd.DataFrame([{
        "ModelID": mid,
        "Brand (site)": r.get("Brand_Site"), "Model (site)": r.get("Model_Site"),
        "Release Date": r.get("Release_Date"), "Release Year": r.get("Release_Year"),
        "Status": r.get("Status"), "Features": r.get("Features"),
        "Description": r.get("Description"), "Photo (large)": r.get("Photo_Large"),
        "Spec count": len(r["specs"]),
        "Page added": r.get("Page_Added"), "Page updated": r.get("Page_Updated"),
    } for mid, r in cache.items()]).set_index("ModelID")

    out = df.copy()
    out["ModelID"] = out["URL"].map(model_id)
    out = out.merge(info, left_on="ModelID", right_index=True, how="left")
    out = out.merge(spec_df, left_on="ModelID", right_index=True, how="left")
    long_df = pd.DataFrame(long_rows, columns=["ModelID", "Brand", "Model", "Section", "Spec", "Value"])
    return out, long_df, spec_df.shape[1]


def export(out, long_df, errors, path):
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        out.to_excel(xw, sheet_name="Data_with_Specs", index=False)
        long_df.to_excel(xw, sheet_name="Specs_Long", index=False)
        pd.DataFrame(errors, columns=["ModelID", "Tried_URL", "Error"]).to_excel(
            xw, sheet_name="Errors", index=False)
        for name, ws in xw.sheets.items():
            ws.freeze_panes = "B2"
            ws.auto_filter.ref = ws.dimensions
            if name == "Data_with_Specs":              # too many cells to auto-measure
                for col in ws.columns:
                    ws.column_dimensions[col[0].column_letter].width = 18
            else:
                for col in ws.columns:
                    w = max((len(str(c.value)) for c in list(col)[:300] if c.value is not None), default=8)
                    ws.column_dimensions[col[0].column_letter].width = min(max(w + 2, 10), 60)
    print(f"[export] {path}")


# ───────────────────────── main ─────────────────────────
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", default=INPUT_FILE)
    ap.add_argument("--output", default=OUTPUT_FILE)
    ap.add_argument("--limit", type=int, help="only the first N model IDs")
    ap.add_argument("--workers", type=int, default=WORKERS)
    ap.add_argument("--test", help="parse one detail-page URL and print the result")
    args = ap.parse_args()

    if args.test:
        r = session.get(args.test, timeout=30)
        rec = parse_detail(r.text)
        specs = rec.pop("specs")
        print(json.dumps(rec, indent=2, ensure_ascii=False))
        print(f"\n{len(specs)} specs")
        for s, l, v in specs:
            print(f"  [{s}] {l}: {v}")
        return

    df = pd.read_excel(args.input)
    df["_mid"] = df["URL"].map(model_id)
    df = df.dropna(subset=["_mid"])

    # one entry per model ID, with every country URL as a fallback
    groups = df.groupby("_mid")["URL"].apply(list).to_dict()
    ids = list(groups)[: args.limit] if args.limit else list(groups)
    print(f"{len(df):,} rows -> {len(groups):,} unique models "
          f"(fetching {len(ids):,}, {len(groups) - len(ids):,} skipped)")

    cache = load_cache()
    todo = [m for m in ids if m not in cache]
    print(f"{len(ids) - len(todo):,} already cached, {len(todo):,} to fetch")

    errors, done = [], 0
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futs = {ex.submit(fetch_model, m, groups[m]): m for m in todo}
        for fut in as_completed(futs):
            mid = futs[fut]
            rec, err = fut.result()
            done += 1
            if rec:
                with cache_lock:
                    cache[mid] = rec
                    with CACHE_FILE.open("a", encoding="utf-8") as f:
                        f.write(json.dumps(rec, ensure_ascii=False) + "\n")
            else:
                errors.append((mid, groups[mid][0], err))
            if done % 50 == 0 or done == len(todo):
                print(f"  {done}/{len(todo)} fetched  ({len(errors)} errors)")

    use = {m: cache[m] for m in ids if m in cache}
    out, long_df, n_spec_cols = build_outputs(df.drop(columns="_mid"), use)
    print(f"{len(use):,} models with data, {n_spec_cols} distinct spec columns")
    export(out, long_df, errors, args.output)


if __name__ == "__main__":
    main()