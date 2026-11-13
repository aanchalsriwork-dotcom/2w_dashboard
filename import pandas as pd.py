"""
Scrape the product catalogue of a fasterwheeler.com country site (default: Brazil)
into Excel with the columns:

    ModelName | Price | Currency | Photo | Year | Category | Brand | URL

Usage:
    pip install requests beautifulsoup4 lxml pandas openpyxl
    python scrape_listing.py                          # Brazil
    python scrape_listing.py --country in             # another country site
    python scrape_listing.py --enrich                 # also open product pages to fill blanks
    python scrape_listing.py --max-pages 300 --delay 0.5

How it works:
  1. Starts at /<country>/home and the category roots (bikes, electric-bike, bicycle).
  2. Follows listing / brand / pagination links inside those categories (BFS).
  3. On every page, finds product cards: <a href=".../<slug>-<id>"> + image + name + "BRL 9666".
  4. De-duplicates by product URL and writes product_listing.xlsx (+ crawl_log.csv).
"""
import argparse
import re
import time
from collections import deque
from urllib.parse import urljoin, urlparse, urldefrag

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = "https://fasterwheeler.com"
CATEGORIES = ["bikes", "electric-bike", "bicycle"]
OUTPUT_XLSX = "product_listing.xlsx"
COLUMNS = ["ModelName", "Price", "Currency", "Photo", "Year", "Category", "Brand", "URL"]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

PRICE_RE = re.compile(r"\b([A-Z]{3})\s*([\d][\d,\.]*)")
YEAR_RE = re.compile(r"\b(20\d{2})\b")

# Brands whose names contain more than one word (slug form). Extend as needed.
MULTI_WORD_BRANDS = [
    "harley-davidson", "royal-enfield", "mv-agusta", "moto-guzzi", "moto-morini",
    "can-am", "big-boy", "land-rover", "tvs-motor", "hero-motocorp", "super-soco",
    "ultra-violette", "ather-energy", "ola-electric", "revolt-motors", "bmw-motorrad",
    "zero-motorcycles", "gas-gas", "sym-motors", "keeway-motors", "cf-moto",
    "ampere-vehicles", "okinawa-autotech", "pure-ev", "simple-energy", "bike-bd",
]


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def clean(t):
    return re.sub(r"\s+", " ", t or "").strip()


def get(session, url, retries=4, timeout=30):
    for attempt in range(1, retries + 1):
        try:
            r = session.get(url, headers=HEADERS, timeout=timeout)
            if r.status_code == 200:
                return r.text
            if r.status_code in (404, 403, 410):
                return None
        except requests.RequestException:
            pass
        time.sleep(2 ** attempt)
    return None


def product_pattern(country):
    cats = "|".join(CATEGORIES)
    return re.compile(rf"^/{country}/({cats})/([a-z0-9\-_.]+?)-(\d+)/?$", re.I)


def split_price(text):
    """'BRL 9,666' -> ('BRL', 9666.0)"""
    m = PRICE_RE.search(text or "")
    if not m:
        return None, None
    cur, num = m.group(1), m.group(2).rstrip(".,").replace(",", "")
    try:
        val = float(num)
        return cur, int(val) if val.is_integer() else val
    except ValueError:
        return cur, None


def split_year(name):
    """'Suzuki Gixxer 2026' -> ('Suzuki Gixxer', 2026)"""
    m = YEAR_RE.search(name)
    if not m:
        return name, None
    year = int(m.group(1))
    base = clean((name[: m.start()] + " " + name[m.end():]))
    return base, year


def guess_brand(slug):
    s = slug.lower()
    for b in MULTI_WORD_BRANDS:
        if s.startswith(b + "-") or s == b:
            return b.replace("-", " ")
    return s.split("-")[0]


def img_url(img, page_url):
    for attr in ("data-src", "data-original", "data-lazy-src", "src"):
        v = img.get(attr)
        if v and not v.startswith("data:"):
            return urljoin(page_url, v)
    return ""


# --------------------------------------------------------------------------- #
# Card parsing
# --------------------------------------------------------------------------- #
def parse_cards(html, page_url, country):
    soup = BeautifulSoup(html, "lxml")
    pat = product_pattern(country)

    def prod_path(a):
        href = a.get("href")
        if not href:
            return None
        p = urlparse(urljoin(page_url, href)).path
        return p if pat.match(p) else None

    anchors = [(a, prod_path(a)) for a in soup.find_all("a", href=True)]
    anchors = [(a, p) for a, p in anchors if p]

    cards = {}
    for a, path in anchors:
        url = urljoin(BASE, path)
        m = pat.match(path)
        category, slug = m.group(1).lower(), m.group(2)

        # --- name: bold text > anchor text > image alt
        strong = a.find(["strong", "b"])
        img = a.find("img")
        name = clean(strong.get_text()) if strong else ""
        if not name:
            name = clean(a.get_text(" "))
        if not name and img is not None:
            name = clean(img.get("alt"))
        if not name:
            continue

        # --- price: look in anchor, then the smallest ancestor that holds only this product
        currency = price = None
        node = a
        for _ in range(4):
            if node is None:
                break
            own_links = {
                prod_path(x) for x in node.find_all("a", href=True) if prod_path(x)
            } if node is not a else {path}
            if own_links - {path}:
                break  # ancestor contains other products -> stop climbing
            txt = clean(node.get_text(" "))
            txt = txt.replace(name, " ")
            currency, price = split_price(txt)
            if price is not None:
                break
            node = node.parent
        if price is None:  # plain text right after the anchor: '</a> BRL 9666'
            sib = a.next_sibling
            hops = 0
            while sib is not None and hops < 3:
                t = sib if isinstance(sib, str) else sib.get_text(" ")
                currency, price = split_price(t)
                if price is not None:
                    break
                sib, hops = sib.next_sibling, hops + 1

        photo = img_url(img, page_url) if img is not None else ""

        base_name, year = split_year(name)
        rec = {
            "ModelName": base_name,
            "Price": price,
            "Currency": currency,
            "Photo": photo,
            "Year": year,
            "Category": category,
            "Brand": guess_brand(slug),
            "URL": url,
        }
        # keep the most complete version if the same product appears twice on a page
        old = cards.get(url)
        if old is None or sum(v not in (None, "") for v in rec.values()) > sum(
            v not in (None, "") for v in old.values()
        ):
            cards[url] = rec
    return soup, cards


def listing_links(soup, page_url, country, pat):
    """Links to follow: listing/brand/pagination pages inside the same country+category."""
    out = set()
    cat_re = re.compile(rf"^/{country}/({'|'.join(CATEGORIES)})(/|$)", re.I)
    for a in soup.find_all("a", href=True):
        full, _ = urldefrag(urljoin(page_url, a["href"]))
        u = urlparse(full)
        if u.netloc and u.netloc.replace("www.", "") != "fasterwheeler.com":
            continue
        if not cat_re.match(u.path) or pat.match(u.path):
            continue
        if "top-ten" in u.path:          # duplicates of the main listing
            continue
        out.add(full)
    return out


# --------------------------------------------------------------------------- #
# Optional: fill blanks from product page
# --------------------------------------------------------------------------- #
def enrich(session, rec, delay):
    time.sleep(delay)
    html = get(session, rec["URL"])
    if not html:
        return
    soup = BeautifulSoup(html, "lxml")
    if rec["Price"] is None:
        m = re.search(r"Price in [A-Za-z .'-]+?\s+([A-Z]{3})\s*([\d][\d,\.]*)", soup.get_text(" "))
        if m:
            rec["Currency"], rec["Price"] = split_price(f"{m.group(1)} {m.group(2)}")
    if not rec["Photo"]:
        og = soup.find("meta", property="og:image")
        if og and og.get("content"):
            rec["Photo"] = og["content"]
    if rec["Year"] is None:
        h1 = soup.find(["h1", "title"])
        if h1:
            _, rec["Year"] = split_year(clean(h1.get_text()))


# --------------------------------------------------------------------------- #
# Excel
# --------------------------------------------------------------------------- #
def write_excel(df):
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="Products", index=False)
        ws = xw.sheets["Products"]
        arial = Font(name="Arial", size=10)
        for c in ws[1]:
            c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", start_color="1F3864")
            c.alignment = Alignment(vertical="center")
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font = arial
        widths = {"ModelName": 38, "Price": 12, "Currency": 10, "Photo": 50,
                  "Year": 8, "Category": 14, "Brand": 18, "URL": 60}
        for i, col in enumerate(df.columns, 1):
            ws.column_dimensions[get_column_letter(i)].width = widths.get(col, 15)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    print(f"Saved {OUTPUT_XLSX} ({len(df)} products)")


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--country", default="br", help="site code, e.g. br, in, uk, de")
    ap.add_argument("--delay", type=float, default=0.5)
    ap.add_argument("--max-pages", type=int, default=400, help="safety cap on listing pages")
    ap.add_argument("--enrich", action="store_true", help="open product pages to fill blanks")
    args = ap.parse_args()

    country = args.country.lower()
    pat = product_pattern(country)
    session = requests.Session()

    seeds = [f"{BASE}/{country}/home"] + [f"{BASE}/{country}/{c}/" for c in CATEGORIES]
    queue, seen = deque(seeds), set(seeds)
    products, log = {}, []

    pages = 0
    while queue and pages < args.max_pages:
        url = queue.popleft()
        pages += 1
        time.sleep(args.delay)
        html = get(session, url)
        if html is None:
            log.append({"page": url, "status": "failed", "cards": 0})
            print(f"[{pages}] FAILED  {url}")
            continue
        soup, cards = parse_cards(html, url, country)
        new = 0
        for u, rec in cards.items():
            if u not in products:
                products[u] = rec
                new += 1
        for link in listing_links(soup, url, country, pat):
            if link not in seen:
                seen.add(link)
                queue.append(link)
        log.append({"page": url, "status": "ok", "cards": len(cards)})
        print(f"[{pages}] {len(cards):3d} cards ({new} new) | total {len(products)} | queue {len(queue)} | {url}")

    if args.enrich:
        todo = [r for r in products.values() if r["Price"] is None or not r["Photo"] or r["Year"] is None]
        print(f"Enriching {len(todo)} products from their own pages ...")
        for i, rec in enumerate(todo, 1):
            enrich(session, rec, args.delay)
            if i % 25 == 0:
                print(f"  {i}/{len(todo)}")

    pd.DataFrame(log).to_csv("crawl_log.csv", index=False)
    df = pd.DataFrame(list(products.values()), columns=COLUMNS)
    df = df.sort_values(["Category", "Brand", "ModelName"]).reset_index(drop=True)
    write_excel(df)


if __name__ == "__main__":
    main()