#!/usr/bin/env python3
"""
Fasterwheeler multi-country model scraper  ->  Country, ModelName, Price, Photo

pip install requests beautifulsoup4 pandas openpyxl

python fasterwheeler_scraper.py                       # all 5 countries
python fasterwheeler_scraper.py --countries ar bo     # subset
python fasterwheeler_scraper.py --debug-url https://fasterwheeler.com/ar/home
"""
import argparse
import datetime as dt
import re
import time
from collections import deque
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE = "https://fasterwheeler.com"
COUNTRIES = {"bz": "Brazil", "hn": "Honduras", "pe": "Peru", "uy": "Uruguay"}
CATEGORIES = ["bikes", "electric-bike", "bicycle"]
DELAY = 0.6
MAX_PAGES_PER_COUNTRY = 600          # safety cap
OUT_FILE = f"fasterwheeler_models_{dt.date.today():%Y%m%d}.xlsx"
GENERIC_OG = "og-site-imgs"          # site-wide placeholder image, not a model photo

session = requests.Session()
session.headers.update({"User-Agent": "Mozilla/5.0 (price-research script)",
                        "Accept-Language": "en-US,en;q=0.9"})
session.mount("https://", HTTPAdapter(max_retries=Retry(
    total=3, backoff_factor=1.5, status_forcelist=[429, 500, 502, 503, 504])))

CAT_ALT = "|".join(CATEGORIES)


def patterns(cc):
    return {
        # /ar/bikes/suzuki-gixxer-10232
        "model": re.compile(rf"^/{cc}/({CAT_ALT})/([a-z0-9\-_.]+-\d+)/?$", re.I),
        # /ar/bikes/ktm/   (slug does not end in -digits, not a .php page)
        "brand": re.compile(rf"^/{cc}/({CAT_ALT})/([a-z0-9\-_]+)/?$", re.I),
        # /ar/bike-price-ars-2000001-to-ars-3000000   (all-brand price ranges)
        "range": re.compile(rf"^/{cc}/(bike|electric-bike|bicycle)-price-[a-z]{{3}}-\d+-to-[a-z]{{3}}-\d+/?$", re.I),
    }


# ───────────────────────── helpers ─────────────────────────
def get(url):
    try:
        r = session.get(url, timeout=30)
        if r.status_code != 200:
            print(f"   ! HTTP {r.status_code}  {url}")
            return None
        return r.text
    except requests.RequestException as e:
        print(f"   ! {e.__class__.__name__}  {url}")
        return None


def clean_url(u):
    p = urlparse(u)
    q = f"?{p.query}" if re.search(r"page=\d+", p.query or "") else ""
    return f"{p.scheme}://{p.netloc}{p.path.rstrip('/')}{q}"


def parse_price(text):
    """'ARS 2,162,900' / 'ARS 153940' / 'COP 1.234.567' -> ('ARS', 2162900)"""
    m = re.search(r"\b([A-Z]{3})\s*([\d][\d.,]*)", text or "")
    if not m:
        return None, None
    cur, raw = m.groups()
    raw = raw.rstrip(".,")
    raw = raw.replace(",", "")
    if raw.count(".") > 1:
        raw = raw.replace(".", "")
    elif raw.count(".") == 1:
        left, right = raw.split(".")
        if len(right) == 3 and len(left) <= 3:      # 1.234 -> thousands separator
            raw = left + right
    try:
        val = float(raw)
    except ValueError:
        return cur, None
    return cur, int(val) if val == int(val) else val


def photo_from(card, page_url):
    for img in card.find_all("img"):
        for attr in ("data-src", "data-original", "data-lazy-src", "data-lazy", "src"):
            v = img.get(attr)
            if v and not v.startswith("data:"):
                return urljoin(page_url, v)
        srcset = img.get("srcset") or img.get("data-srcset")
        if srcset:
            return urljoin(page_url, srcset.split()[0])
    return None


# ───────────────────────── card extraction ─────────────────────────
def card_container(a, pat):
    """Climb from the link to the smallest element that holds exactly one model."""
    def model_paths(node):
        return {urlparse(x["href"]).path.rstrip("/")
                for x in node.find_all("a", href=True) if pat["model"].match(urlparse(x["href"]).path)}
    node = a
    for _ in range(5):
        p = node.parent
        if p is None or p.name in ("body", "html", "main") or len(model_paths(p)) > 1:
            break
        node = p
    return node


def extract_cards(soup, page_url, cc, brand_ctx=None):
    pat = patterns(cc)
    cards = {}
    for a in soup.find_all("a", href=True):
        path = urlparse(urljoin(page_url, a["href"])).path
        m = pat["model"].match(path)
        if not m:
            continue
        url = BASE + path.rstrip("/")
        if url in cards:
            continue
        card = card_container(a, pat)
        text = card.get_text(" ", strip=True)
        cur, price = parse_price(text)

        img = card.find("img")
        name = (img.get("alt") or "").strip() if img else ""
        if not name:
            lines = [l for l in card.get_text("\n", strip=True).split("\n") if not re.search(r"\b[A-Z]{3}\s*\d", l)]
            name = lines[0] if lines else m.group(2)

        year = re.search(r"\b(19|20)\d{2}\b", name)
        cards[url] = {
            "Country": COUNTRIES[cc],
            "ModelName": re.sub(r"\s+", " ", name),
            "Year": int(year.group(0)) if year else None,
            "Price": price,
            "Currency": cur,
            "Photo": photo_from(card, page_url),
            "Category": m.group(1),
            "Brand": brand_ctx,
            "URL": url,
        }
    return cards


# ───────────────────────── crawl one country ─────────────────────────
def crawl_country(cc, max_pages=MAX_PAGES_PER_COUNTRY):
    pat = patterns(cc)
    home = f"{BASE}/{cc}/home"
    queue = deque([(home, None)] + [(f"{BASE}/{cc}/{c}/", None) for c in CATEGORIES])
    seen, models, pages = set(), {}, 0

    while queue and pages < max_pages:
        url, brand_ctx = queue.popleft()
        key = clean_url(url)
        if key in seen:
            continue
        seen.add(key)

        html = get(url)
        time.sleep(DELAY)
        if html is None:
            continue
        pages += 1
        soup = BeautifulSoup(html, "html.parser")

        found = extract_cards(soup, url, cc, brand_ctx)
        for u, rec in found.items():
            if u in models:
                models[u]["Brand"] = models[u]["Brand"] or rec["Brand"]
            else:
                models[u] = rec
        print(f"[{cc}] {url.replace(BASE, '')}  -> {len(found)} cards  (total {len(models)})")

        # discover more listing pages
        for a in soup.find_all("a", href=True):
            href = urljoin(url, a["href"])
            p = urlparse(href)
            if p.netloc and p.netloc != urlparse(BASE).netloc:
                continue
            path = p.path
            if path.endswith(".php") or "/features/" in path:
                continue
            if pat["model"].match(path):
                continue
            if pat["range"].match(path):
                queue.append((f"{BASE}{path}", None))
            elif pat["brand"].match(path):
                slug = pat["brand"].match(path).group(2)
                if slug not in ("home",):
                    queue.append((f"{BASE}{path.rstrip('/')}/", slug))
            elif re.search(r"[?&]page=\d+", p.query or "") and path.startswith(f"/{cc}/"):
                queue.append((href, brand_ctx))

    return models


# ───────────────────────── fill missing photos ─────────────────────────
def fill_missing_photos(models):
    missing = [r for r in models if not r["Photo"] or GENERIC_OG in r["Photo"]]
    if not missing:
        return
    print(f"[photos] {len(missing)} models without a listing photo - checking detail pages")
    for r in missing:
        html = get(r["URL"])
        time.sleep(DELAY)
        if not html:
            continue
        soup = BeautifulSoup(html, "html.parser")
        tag = (soup.find("meta", attrs={"property": "og:image"})
               or soup.find("meta", attrs={"name": "twitter:image"}))
        src = tag.get("content") if tag else None
        if src and GENERIC_OG not in src:
            r["Photo"] = urljoin(r["URL"], src)


# ───────────────────────── export ─────────────────────────
def export(df, path):
    summary = (df.groupby(["Country", "Category"]).agg(
        Models=("URL", "count"),
        With_price=("Price", lambda s: s.notna().sum()),
        With_photo=("Photo", lambda s: s.notna().sum())).reset_index())

    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="Models", index=False)
        summary.to_excel(xw, sheet_name="Summary", index=False)
        for ws in xw.sheets.values():
            for col in ws.columns:
                w = max((len(str(c.value)) for c in col if c.value is not None), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(max(w + 2, 10), 55)
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
        ws = xw.sheets["Models"]
        pc = {c.value: c.column_letter for c in ws[1]}["Price"]
        for cell in ws[pc][1:]:
            cell.number_format = "#,##0"
    df.to_csv(path.replace(".xlsx", ".csv"), index=False, encoding="utf-8-sig")
    print(f"[export] {path}  (+ .csv)")


# ───────────────────────── main ─────────────────────────
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--countries", nargs="+", default=list(COUNTRIES), choices=list(COUNTRIES))
    ap.add_argument("--debug-url", help="print the cards parsed from one page")
    ap.add_argument("--max-pages", type=int, default=MAX_PAGES_PER_COUNTRY)
    ap.add_argument("--no-photo-fill", action="store_true")
    args = ap.parse_args()

    if args.debug_url:
        cc = urlparse(args.debug_url).path.split("/")[1]
        soup = BeautifulSoup(get(args.debug_url) or "", "html.parser")
        for rec in extract_cards(soup, args.debug_url, cc).values():
            print(rec)
        return

    all_rows = []
    for cc in args.countries:
        print(f"\n=== {COUNTRIES[cc]} ===")
        models = crawl_country(cc, args.max_pages)
        rows = list(models.values())
        if not args.no_photo_fill:
            fill_missing_photos(rows)
        all_rows += rows
        print(f"=== {COUNTRIES[cc]}: {len(rows)} models ===")

    df = pd.DataFrame(all_rows)
    cols = ["Country", "ModelName", "Price", "Currency", "Photo", "Year", "Category", "Brand", "URL"]
    df = df[cols].sort_values(["Country", "Category", "Brand", "ModelName"], na_position="last")
    export(df, OUT_FILE)


if __name__ == "__main__":
    main()