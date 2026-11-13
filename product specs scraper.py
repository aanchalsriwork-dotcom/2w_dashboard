"""
Scrape product description + specification table for every URL in pricing.csv
and save the result to Excel.

Usage:
    pip install requests beautifulsoup4 lxml pandas openpyxl
    python scrape_specs.py                       # all rows
    python scrape_specs.py --limit 20            # quick test on first 20 rows
    python scrape_specs.py --workers 8 --delay 0.3

Features:
  * Parallel fetching with retries + polite delay
  * Resumable: every result is appended to specs_cache.jsonl, so if the run is
    interrupted, re-running only fetches URLs that are still missing
  * Excel output with 3 sheets:
        Products   - 1 row per product (original CSV columns + description + one
                     column per spec, named "<Section> | <Spec>")
        Specs_Long - tidy format: URL, Section, Spec, Value (good for pivoting)
        Errors     - URLs that failed / had no spec table
"""
import argparse
import json
import os
import random
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime

ts = datetime.now().strftime("%Y%m%d_%H%M%S")
INPUT_CSV = r"C:\Users\91931\2w_dashboard\argentina_prices.csv"
OUTPUT_XLSX = f"product_specsts_{ts}.xlsx"
CACHE_FILE = "specs_cache.jsonl"
URL_COL = "URL"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

_write_lock = threading.Lock()


# --------------------------------------------------------------------------- #
# Parsing
# --------------------------------------------------------------------------- #
def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def find_spec_table(soup):
    """Return the <table> that holds the specs (the one with a 'Brand' / 'Model' row)."""
    best = None
    for table in soup.find_all("table"):
        first_cells = [
            clean(tr.find(["td", "th"]).get_text())
            for tr in table.find_all("tr")
            if tr.find(["td", "th"])
        ]
        if any(c.lower() in ("brand", "model") for c in first_cells):
            return table
        # fallback: the biggest table on the page
        if best is None or len(table.find_all("tr")) > len(best.find_all("tr")):
            best = table
    return best


def extract_description(soup, table):
    """
    Description = paragraph(s) between the '<X> Specifications' heading and the table.
    Falls back to the closest <p> above the table.
    """
    # 1) Heading that contains 'Specifications'
    heading = None
    for h in soup.find_all(re.compile(r"^h[1-6]$")):
        if "specification" in h.get_text().lower():
            heading = h
            break

    if heading is not None:
        parts = []
        for sib in heading.find_next_siblings():
            if sib is table or (table is not None and table in sib.find_all("table")):
                break
            if sib.name in ("p", "div", "span") and clean(sib.get_text()):
                parts.append(clean(sib.get_text()))
        if parts:
            return " ".join(parts)

    # 2) Closest paragraph(s) before the table
    if table is not None:
        parts = []
        for prev in table.find_all_previous("p", limit=3):
            t = clean(prev.get_text())
            if len(t) > 40:
                parts.append(t)
                break
        if parts:
            return parts[0]
    return ""


def extract_specs(table):
    """
    Walk table rows.
      - Row with ONE cell  -> section header (e.g. 'Engine', 'Dimensions and Weights')
      - Row with 2+ cells  -> spec name / value
    Returns list of (section, key, value).
    """
    rows = []
    section = "General"
    if table is None:
        return rows

    for tr in table.find_all("tr"):
        cells = tr.find_all(["td", "th"])
        if not cells:
            continue
        texts = [clean(c.get_text(" ")) for c in cells]

        if len(cells) == 1 or (len(texts) >= 2 and texts[0] and not any(texts[1:])):
            section = texts[0] or section
            continue

        key, value = texts[0], texts[1]
        if key:
            rows.append((section, key, value))
    return rows


def parse_page(html: str):
    soup = BeautifulSoup(html, "lxml")
    table = find_spec_table(soup)
    description = extract_description(soup, table)
    specs = extract_specs(table)
    return description, specs


# --------------------------------------------------------------------------- #
# Fetching
# --------------------------------------------------------------------------- #
def fetch(session, url, retries=4, timeout=30):
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            r = session.get(url, headers=HEADERS, timeout=timeout)
            if r.status_code == 200:
                return r.text
            if r.status_code in (429, 500, 502, 503, 504):
                last_err = f"HTTP {r.status_code}"
                time.sleep(2 ** attempt + random.random())
                continue
            return None if r.status_code == 404 else r.raise_for_status()
        except requests.RequestException as e:
            last_err = str(e)
            time.sleep(2 ** attempt + random.random())
    raise RuntimeError(last_err or "unknown error")


def scrape_one(url, delay):
    session = requests.Session()
    time.sleep(delay * random.uniform(0.5, 1.5))
    try:
        html = fetch(session, url)
        if html is None:
            return {"URL": url, "status": "404", "description": "", "specs": []}
        desc, specs = parse_page(html)
        status = "ok" if specs else "no_spec_table"
        return {"URL": url, "status": status, "description": desc, "specs": specs}
    except Exception as e:  # noqa: BLE001
        return {"URL": url, "status": f"error: {e}", "description": "", "specs": []}


# --------------------------------------------------------------------------- #
# Cache (resume support)
# --------------------------------------------------------------------------- #
def load_cache():
    cache = {}
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, encoding="utf-8") as f:
            for line in f:
                try:
                    rec = json.loads(line)
                    cache[rec["URL"]] = rec
                except Exception:  # noqa: BLE001
                    continue
    return cache


def append_cache(rec):
    with _write_lock, open(CACHE_FILE, "a", encoding="utf-8") as f:
        f.write(json.dumps(rec, ensure_ascii=False) + "\n")


# --------------------------------------------------------------------------- #
# Excel output
# --------------------------------------------------------------------------- #
def style_sheet(ws, wrap_cols=(), width_cap=60):
    arial = Font(name="Arial", size=10)
    head_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", start_color="1F3864")

    for cell in ws[1]:
        cell.font, cell.fill = head_font, head_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = arial

    for idx, col in enumerate(ws.columns, start=1):
        header = str(col[0].value or "")
        longest = max((len(str(c.value)) for c in col[:200] if c.value is not None), default=10)
        ws.column_dimensions[get_column_letter(idx)].width = min(max(12, longest + 2), width_cap)
        if header in wrap_cols:
            for c in col[1:]:
                c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def build_excel(df_in, results):
    wide_rows, long_rows, error_rows = [], [], []

    for _, src in df_in.iterrows():
        url = src[URL_COL]
        rec = results.get(url)
        base = src.to_dict()
        if rec is None:
            error_rows.append({URL_COL: url, "status": "not_fetched"})
            wide_rows.append(base)
            continue

        row = dict(base)
        row["Description"] = rec["description"]
        seen = {}
        for section, key, value in rec["specs"]:
            col = f"{section} | {key}"
            seen[col] = seen.get(col, 0) + 1
            if seen[col] > 1:                      # duplicate spec name inside a section
                col = f"{col} ({seen[col]})"
            row[col] = value
            long_rows.append({URL_COL: url, "Section": section, "Spec": key, "Value": value})
        wide_rows.append(row)

        if rec["status"] != "ok":
            error_rows.append({URL_COL: url, "status": rec["status"]})

    df_wide = pd.DataFrame(wide_rows)
    # keep original columns first, then Description, then specs in first-seen order
    orig = list(df_in.columns)
    rest = [c for c in df_wide.columns if c not in orig]
    df_wide = df_wide[orig + rest]
    df_long = pd.DataFrame(long_rows, columns=[URL_COL, "Section", "Spec", "Value"])
    df_err = pd.DataFrame(error_rows, columns=[URL_COL, "status"])

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as xw:
        df_wide.to_excel(xw, sheet_name="Products", index=False)
        df_long.to_excel(xw, sheet_name="Specs_Long", index=False)
        
        
        df_err.to_excel(xw, sheet_name="Errors", index=False)
        style_sheet(xw.sheets["Products"], wrap_cols=("Description",), width_cap=60)
        style_sheet(xw.sheets["Specs_Long"], width_cap=50)
        style_sheet(xw.sheets["Errors"], width_cap=80)

    print(f"Saved {OUTPUT_XLSX}: {len(df_wide)} products, "
          f"{len(df_long)} spec rows, {len(df_err)} problem URLs")


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=None, help="only first N rows (testing)")
    ap.add_argument("--workers", type=int, default=6)
    ap.add_argument("--delay", type=float, default=0.5, help="base delay (s) before each request")
    args = ap.parse_args()

    df = pd.read_csv(INPUT_CSV, encoding="utf-8-sig")
    df = df.dropna(subset=[URL_COL]).drop_duplicates(subset=[URL_COL]).reset_index(drop=True)
    if args.limit:
        df = df.head(args.limit)

    cache = load_cache()
    todo = [u for u in df[URL_COL] if u not in cache or cache[u]["status"].startswith("error")]
    print(f"{len(df)} URLs total | {len(df) - len(todo)} cached | {len(todo)} to fetch")

    if todo:
        with ThreadPoolExecutor(max_workers=args.workers) as ex:
            futures = {ex.submit(scrape_one, u, args.delay): u for u in todo}
            for i, fut in enumerate(as_completed(futures), 1):
                rec = fut.result()
                cache[rec["URL"]] = rec
                append_cache(rec)
                if i % 25 == 0 or i == len(todo):
                    print(f"  {i}/{len(todo)} done")

    build_excel(df, cache)


if __name__ == "__main__":
    main()