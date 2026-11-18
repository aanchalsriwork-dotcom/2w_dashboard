#!/usr/bin/env python3
"""
Fasterwheeler country scraper  ->  Country, ModelName, Price, Currency, Photo (+ extras)

Run ONE country at a time (any code listed on https://fasterwheeler.com/countries.php):

    python Final_Price_Fetch.py --list-countries
    python Final_Price_Fetch.py --country in
    python Final_Price_Fetch.py --country pk bd            # one file per country, one after another
    python Final_Price_Fetch.py --country in --categories bikes electric-bike bicycle
    python Final_Price_Fetch.py --debug-url https://fasterwheeler.com/in/home    # check how a page parses

pip install requests beautifulsoup4 pandas openpyxl
"""
import argparse
import datetime as dt
import re
import sys
import time
from collections import deque
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE = "https://fasterwheeler.com"
COUNTRIES_URL = f"{BASE}/countries.php"
DEFAULT_CATEGORIES = ["bikes", "electric-bike"]
ALL_CATEGORIES = ["bikes", "electric-bike", "bicycle"]
DELAY = 0.6
MAX_PAGES = 600                      # safety cap per country
# Optional: country code(s) to run when you start the script WITHOUT --country
# (handy in an IDE where passing options is awkward), e.g. RUN_COUNTRIES = ["in"]
RUN_COUNTRIES = []
GENERIC_OG = "og-site-imgs"          # site-wide placeholder image, not a model photo

session = requests.Session()
session.headers.update({"User-Agent": "Mozilla/5.0 (price-research script)",
                        "Accept-Language": "en-US,en;q=0.9"})
session.mount("https://", HTTPAdapter(max_retries=Retry(
    total=3, backoff_factor=1.5, status_forcelist=[429, 500, 502, 503, 504])))


# ───────────────────────── http ─────────────────────────
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


# ───────────────────────── country list ─────────────────────────
def parse_countries(html):
    """{'in': 'India', 'pk': 'Pakistan', ...} from the links on countries.php"""
    soup = BeautifulSoup(html or "", "html.parser")
    out = {}
    for a in soup.find_all("a", href=True):
        m = re.match(r"^/([a-z]{2})/home/?$", urlparse(urljoin(COUNTRIES_URL, a["href"])).path)
        if not m:
            continue
        name = a.get_text(" ", strip=True) or re.sub(r"^Bike Price in\s*", "", a.get("title", "")).strip()
        out.setdefault(m.group(1), name or m.group(1).upper())
    return out


def load_countries():
    return parse_countries(get(COUNTRIES_URL))


# ───────────────────────── url patterns / prefix ─────────────────────────
def detect_prefix(soup, cc):
    """Country sites link to /xx/bikes/...; the global fallback site (e.g. /bz/home) links to /bikes/..."""
    paths = [urlparse(a["href"]).path for a in soup.find_all("a", href=True)]
    if any(re.match(rf"^/{cc}/(bikes|electric-bike|bicycle)/", p) for p in paths):
        return f"/{cc}"
    if any(re.match(r"^/(bikes|electric-bike|bicycle)/[^/]+-\d+/?$", p) for p in paths):
        return ""
    return f"/{cc}"


def patterns(prefix, categories):
    pre = re.escape(prefix)
    cats = "|".join(categories)
    return {
        "model": re.compile(rf"^{pre}/({cats})/([a-z0-9\-_.]+-\d+)/?$", re.I),     # /in/bikes/suzuki-gixxer-10232
        "brand": re.compile(rf"^{pre}/({cats})/([a-z0-9\-_]+)/?$", re.I),          # /in/bikes/ktm/
        "range": re.compile(rf"^{pre}/(bike|electric-bike|bicycle)-price-[a-z]{{3}}-\d+-to-[a-z]{{3}}-\d+/?$", re.I),
    }


def site_code(soup):
    tag = soup.find("meta", attrs={"name": "sitecode"}) or soup.find("meta", attrs={"property": "sitecode"})
    return (tag.get("content") or "").strip().lower() if tag else None


# ───────────────────────── price parsing ─────────────────────────
# Formats seen: "$ 1790" · "CAN$ 2291" · "Euro 1611" · "GBP 1307" · "ARS 153940" · "BOB 12,351" · "$1,790"
CURRENCY_MAP = {"$": "USD", "US$": "USD", "USD": "USD", "CAN$": "CAD", "CA$": "CAD", "C$": "CAD",
                "CAD": "CAD", "A$": "AUD", "AU$": "AUD", "AUD": "AUD", "NZ$": "NZD",
                "€": "EUR", "EURO": "EUR", "EUR": "EUR", "£": "GBP", "GBP": "GBP"}
SYMBOLS = set("$€£¥₹₽₩₺")
NUM = r"\d[\d.,]*(?:[\s\u00a0]\d{3})*"
PRICE_PREFIX = re.compile(rf"^\s*(?P<cur>[^\d\s,]{{1,8}})\s*(?P<num>{NUM})\s*$")     # "Rs. 1,50,000"
PRICE_SUFFIX = re.compile(rf"^\s*(?P<num>{NUM})\s*(?P<cur>[^\d\s,]{{1,8}})\s*$")     # "1.234,56 €"


def is_known_currency(t):
    t = t.strip()
    return bool(t) and (t in SYMBOLS or t.upper() in CURRENCY_MAP or t.endswith("$")
                        or re.fullmatch(r"[A-Z]{3}", t) is not None)


def normalise_currency(tok):
    t = tok.strip().rstrip(".")
    return CURRENCY_MAP.get(t.upper(), t)


def to_number(raw):
    """'1790' '1,790' '2,162,900' '1.234.567' '1,50,000' '1.790,50' '1,790.50' '12 345' -> number"""
    raw = re.sub(r"[\s\u00a0]", "", raw).rstrip(".,")
    if "," in raw and "." in raw:                       # the last separator is the decimal one
        if raw.rfind(",") > raw.rfind("."):
            raw = raw.replace(".", "").replace(",", ".")
        else:
            raw = raw.replace(",", "")
    elif "," in raw:
        raw = raw.replace(",", "") if len(raw.split(",")[-1]) == 3 else raw.replace(",", ".")
    elif "." in raw:
        parts = raw.split(".")
        if len(parts) > 2 or (len(parts[1]) == 3 and len(parts[0]) <= 3):
            raw = raw.replace(".", "")
    try:
        val = float(raw)
    except ValueError:
        return None
    return int(val) if val == int(val) else val


def match_price(text, strict):
    """strict=True additionally requires a recognisable currency token (used for text inside the link)."""
    for rx in (PRICE_PREFIX, PRICE_SUFFIX):
        m = rx.match(text or "")
        if m and (not strict or is_known_currency(m.group("cur"))):
            return m
    return None


def price_from_card(card):
    """(price, currency, raw_text). Works on individual text nodes, never on the whole card text,
    so model names like 'KTM 790 ADVENTURE 2026' or 'Zero DSR 2026' can't be read as a price.
    The site puts the price outside the model link, so that text is checked first (any currency
    token accepted); text inside the link is a fallback and needs a known currency token."""
    outside, inside = [], []
    for s in card.find_all(string=True):
        if s.parent is None or s.parent.name in ("script", "style"):
            continue
        t = str(s).strip()
        if t:
            (inside if s.find_parent("a") is not None else outside).append(t)
    for pool, strict in ((outside, False), (inside, True)):
        for t in reversed(pool):
            m = match_price(t, strict)
            if m:
                val = to_number(m.group("num"))
                if val:
                    return val, normalise_currency(m.group("cur")), t
    return None, None, None


# ───────────────────────── card extraction ─────────────────────────
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


def extract_cards(soup, page_url, pat, country_name, cc, sitecode, brand_ctx=None):
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
        price, cur, price_text = price_from_card(card)

        img = card.find("img")
        name = (img.get("alt") or "").strip() if img else ""
        if not name:
            lines = [s.strip() for s in card.stripped_strings if not match_price(s, strict=True)]
            name = lines[0] if lines else m.group(2)

        year = re.search(r"\b(19|20)\d{2}\b", name)
        mid = re.search(r"-(\d+)$", m.group(2))
        cards[url] = {
            "Country": country_name,
            "CountryCode": cc,
            "ModelName": re.sub(r"\s+", " ", name),
            "Year": int(year.group(0)) if year else None,
            "Price": price,
            "Currency": cur,
            "PriceText": price_text,
            "Photo": photo_from(card, page_url),
            "Category": m.group(1),
            "Brand": brand_ctx,
            "ModelID": mid.group(1) if mid else None,
            "SiteCode": sitecode,
            "URL": url,
        }
    return cards


# ───────────────────────── crawl one country ─────────────────────────
def clean_url(u):
    p = urlparse(u)
    q = f"?{p.query}" if re.search(r"page=\d+", p.query or "") else ""
    return f"{p.scheme}://{p.netloc}{p.path.rstrip('/')}{q}"


def crawl_country(cc, country_name, categories, max_pages=MAX_PAGES, home_url=None):
    home = home_url or f"{BASE}/{cc}/home"
    html = get(home)
    if html is None:
        print(f"Could not open {home} - is '{cc}' a valid country code? (see --list-countries)")
        return {}
    soup = BeautifulSoup(html, "html.parser")
    sc = site_code(soup)
    prefix = detect_prefix(soup, cc)
    pat = patterns(prefix, categories)

    if sc and sc != cc:
        print(f"\n   WARNING: {home} is served by the '{sc}' site, not '{cc}'. "
              f"Prices on it are that site's prices (usually USD), not local {country_name} prices.\n")
    if prefix == "":
        print("   note: this country's links have no /xx/ prefix - using the global URL layout")

    models, seen, pages = {}, {clean_url(home)}, 0
    queue = deque([(f"{BASE}{prefix}/{c}/", None) for c in categories])

    def take(h, url, brand_ctx):
        s = BeautifulSoup(h, "html.parser")
        found = extract_cards(s, url, pat, country_name, cc, sc, brand_ctx)
        for u, rec in found.items():
            if u in models:
                models[u]["Brand"] = models[u]["Brand"] or rec["Brand"]
            else:
                models[u] = rec
        print(f"[{cc}] {url.replace(BASE, '')}  -> {len(found)} cards  (total {len(models)})")
        return s

    def discover(s, url, brand_ctx):
        for a in s.find_all("a", href=True):
            href = urljoin(url, a["href"])
            p = urlparse(href)
            if p.netloc and p.netloc != urlparse(BASE).netloc:
                continue
            path = p.path
            if path.endswith(".php") or "/features/" in path or pat["model"].match(path):
                continue
            # keep ?page=N on the link, otherwise page 2+ of a brand listing is treated as already crawled
            q = f"?{p.query}" if re.search(r"(^|&)page=\d+", p.query or "") else ""
            if pat["range"].match(path):
                queue.append((f"{BASE}{path}{q}", None))
            elif pat["brand"].match(path):
                queue.append((f"{BASE}{path.rstrip('/')}/{q}", pat["brand"].match(path).group(2)))
            elif q and path.startswith(prefix or "/"):
                queue.append((href, brand_ctx))

    try:
        discover(take(html, home, None), home, None)
        while queue and pages < max_pages:
            url, brand_ctx = queue.popleft()
            key = clean_url(url)
            if key in seen:
                continue
            seen.add(key)
            h = get(url)
            time.sleep(DELAY)
            if h is None:
                continue
            pages += 1
            discover(take(h, url, brand_ctx), url, brand_ctx)
    except KeyboardInterrupt:
        print("\nInterrupted - keeping what was collected so far.")
    return models


# ───────────────────────── photos ─────────────────────────
def fill_missing_photos(rows):
    missing = [r for r in rows if not r["Photo"] or GENERIC_OG in r["Photo"]]
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


# ───────────────────────── checks + export ─────────────────────────
def quality_report(df, country_name):
    print(f"\n── {country_name}: {len(df):,} models ──")
    print("  by category:", {k: int(v) for k, v in df["Category"].value_counts().items()})
    cur = df["Currency"].value_counts(dropna=False)
    print("  currency tokens:", {k: int(v) for k, v in cur.items()}, "  <- should be ONE value")
    print(f"  missing price: {df['Price'].isna().sum()}   missing photo: {df['Photo'].isna().sum()}")
    if df["Currency"].nunique() > 1:
        dominant = cur.index[0]
        odd = df[df["Currency"] != dominant]
        print(f"  WARNING: {len(odd)} rows have a different currency token than '{dominant}'. First few:")
        print(odd[["ModelName", "Price", "Currency", "PriceText"]].head(5).to_string(index=False))
    return cur


def export(df, path):
    summary = (df.groupby("Category").agg(
        Models=("URL", "count"),
        With_price=("Price", lambda s: s.notna().sum()),
        With_photo=("Photo", lambda s: s.notna().sum()),
        Currencies=("Currency", lambda s: ", ".join(sorted({str(x) for x in s.dropna()})))).reset_index())
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
            cell.number_format = "#,##0.##"
    df.to_csv(path.replace(".xlsx", ".csv"), index=False, encoding="utf-8-sig")
    print(f"[export] {path}  (+ .csv)")


COLUMNS = ["Country", "CountryCode", "ModelName", "Price", "Currency", "PriceText", "Photo",
           "Year", "Category", "Brand", "ModelID", "SiteCode", "URL"]


def run_country(cc, countries, categories, max_pages, fill_photos, home_url=None, country_name=None):
    name = country_name or countries.get(cc, cc.upper())
    print(f"\n==================== {name} ({cc}) ====================")
    models = crawl_country(cc, name, categories, max_pages, home_url=home_url)
    if not models:
        print("No models found - nothing to export.")
        return
    rows = list(models.values())
    if fill_photos:
        fill_missing_photos(rows)
    df = pd.DataFrame(rows)[COLUMNS].sort_values(["Category", "Brand", "ModelName"], na_position="last")
    quality_report(df, name)
    export(df, f"fasterwheeler_{cc}_{dt.date.today():%Y%m%d}.xlsx")


# ───────────────────────── URL list input ─────────────────────────
def load_url_list(path):
    """Load Country, Country Code and Final URL columns from CSV or Excel."""
    if path.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(path)
    else:
        df = pd.read_csv(path)

    # Ignore differences in column capitalisation and surrounding spaces.
    df.columns = [str(c).strip().lower() for c in df.columns]
    required = {"country", "country code", "final url"}
    if not required.issubset(df.columns):
        raise ValueError(
            "URL list must contain these columns: Country, Country Code, Final URL"
        )

    records = []
    for _, row in df.iterrows():
        country = str(row["country"]).strip()
        cc = str(row["country code"]).strip().lower()
        url = str(row["final url"]).strip()

        if not country or country.lower() == "nan" or not cc or cc == "nan":
            continue
        if not url or url.lower() == "nan":
            continue

        parsed = urlparse(url)
        if parsed.scheme not in ("http", "https") or parsed.netloc.lower() not in (
            "fasterwheeler.com", "www.fasterwheeler.com"
        ):
            print(f"Skipping invalid Fasterwheeler URL for {country}: {url}")
            continue

        records.append((country, cc, url))

    return records


# ───────────────────────── main ─────────────────────────
def prompt_for_countries():
    """Used when the script is started without --country."""
    print("No --country option was given.")
    while True:
        try:
            raw = input("Country code(s), e.g. 'in' or 'in pk bd'  "
                        "('list' shows all codes, Enter quits): ").strip().lower()
        except EOFError:
            return []
        if not raw:
            return []
        if raw == "list":
            for cc, name in load_countries().items():
                print(f"  {cc}  {name}")
            continue
        return raw.replace(",", " ").split()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--country", nargs="+", metavar="CC", help="country code(s), e.g. in pk bd")
    ap.add_argument("--url-list", help="CSV or Excel file containing Country, Country Code, Final URL")
    ap.add_argument("--list-countries", action="store_true")
    ap.add_argument("--categories", nargs="+", default=DEFAULT_CATEGORIES, choices=ALL_CATEGORIES)
    ap.add_argument("--debug-url", help="print the cards parsed from one page")
    ap.add_argument("--max-pages", type=int, default=MAX_PAGES)
    ap.add_argument("--no-photo-fill", action="store_true")
    args = ap.parse_args()

    if args.list_countries:
        for cc, name in load_countries().items():
            print(f"{cc}  {name}")
        return

    if args.debug_url:
        html = get(args.debug_url) or ""
        soup = BeautifulSoup(html, "html.parser")
        seg = urlparse(args.debug_url).path.split("/")[1]
        cc = seg if re.fullmatch(r"[a-z]{2}", seg) else "xx"
        pat = patterns(detect_prefix(soup, cc), ALL_CATEGORIES)
        print("site code:", site_code(soup), "| url prefix:", repr(detect_prefix(soup, cc)))
        for rec in extract_cards(soup, args.debug_url, pat, cc.upper(), cc, site_code(soup)).values():
            print({k: rec[k] for k in ("ModelName", "Price", "Currency", "PriceText", "Category", "Photo")})
        return

    if args.url_list:
        url_records = load_url_list(args.url_list)
        if not url_records:
            ap.error("No valid country URLs found in the supplied list.")

        for country_name, cc, home_url in url_records:
            run_country(
                cc, {}, args.categories, args.max_pages,
                not args.no_photo_fill, home_url=home_url,
                country_name=country_name
            )
        return

    if not args.country:
        args.country = RUN_COUNTRIES or prompt_for_countries()
    if not args.country:
        ap.error("no country given. Example:  python Final_Price_Fetch.py --country in   "
                 "(codes: python Final_Price_Fetch.py --list-countries)")

    countries = load_countries()
    for cc in args.country:
        cc = cc.lower()
        if countries and cc not in countries:
            print(f"'{cc}' is not on countries.php - skipping. Run --list-countries to see valid codes.")
            continue
        run_country(cc, countries, args.categories, args.max_pages, not args.no_photo_fill)


if __name__ == "__main__":
    main()