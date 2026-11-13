"""
Scrape product listings from https://fasterwheeler.com/br/home (and its category pages).

Output columns:
    Country, Model Name, Price, Currency, Photo, Year, Category, Brand, URL

Usage:
    pip install requests beautifulsoup4 lxml pandas
    python fasterwheeler_br_scraper.py
    python fasterwheeler_br_scraper.py --debug      # prints what each page yielded
"""
import argparse
import re
import time
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE = "https://fasterwheeler.com"
COUNTRY = "Brazil"
SEED_PAGES = [
    f"{BASE}/br/home",
    f"{BASE}/br/bikes/",
    f"{BASE}/br/electric-bike/",
    f"{BASE}/br/bicycle/",
    f"{BASE}/br/bikes/upcoming.php",
    f"{BASE}/br/electric-bike/upcoming.php",
    f"{BASE}/br/bicycle/upcoming.php",
    f"{BASE}/br/bikes/top-ten-bikes.php",
    f"{BASE}/br/electric-bike/top-ten-electric-bike.php",
    f"{BASE}/br/bicycle/top-ten-bicycle.php",
]
OUTPUT_CSV = "fasterwheeler_br_products.csv"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

# Product URLs look like /br/bikes/suzuki-gixxer-10232
PRODUCT_RE = re.compile(r"/br/(bikes|electric-bike|bicycle)/[a-z0-9\-]+-\d+/?$", re.I)
PRICE_RE = re.compile(r"\b([A-Z]{3})\s*([\d][\d.,]*)")
YEAR_RE = re.compile(r"\b(20\d{2})\b")
CATEGORY_MAP = {"bikes": "Bike", "electric-bike": "Electric Bike", "bicycle": "Bicycle"}

# Multi-word brands first matter; the list is matched longest-first.
KNOWN_BRANDS = [
    "Royal Enfield", "Harley Davidson", "MV Agusta", "Indian", "Suzuki", "Honda",
    "Kawasaki", "Yamaha", "Bajaj", "Benelli", "Yadea", "Husqvarna", "Piaggio",
    "Aprilia", "Zero", "Triumph", "Ducati", "KTM", "BMW", "TVS", "Hero", "Vespa",
    "CFMoto", "CF Moto", "Moto Guzzi", "Can-Am", "Energica", "Ola", "Ather",
    "Revolt", "Segway", "Niu", "Super Soco", "Gas Gas", "Beta", "Sherco",
]


def clean(t):
    return re.sub(r"\s+", " ", t or "").strip()


def get(session, url, retries=3):
    for i in range(retries):
        try:
            r = session.get(url, headers=HEADERS, timeout=30)
            if r.status_code == 200:
                return r.text
            if r.status_code == 404:
                return None
        except requests.RequestException:
            pass
        time.sleep(2 ** i)
    return None


def split_brand(name):
    low = name.lower()
    for b in sorted(KNOWN_BRANDS, key=len, reverse=True):
        if low.startswith(b.lower() + " "):
            return b, clean(name[len(b):])
    parts = name.split(" ", 1)
    return parts[0], (parts[1] if len(parts) > 1 else "")


def parse_products(html, page_url):
    soup = BeautifulSoup(html, "lxml")
    out = []
    for a in soup.find_all("a", href=True):
        href = urljoin(page_url, a["href"]).split("#")[0]
        if not PRODUCT_RE.search(urlparse(href).path):
            continue

        # Name: prefer <strong>/<b>, then image alt, then anchor text
        strong = a.find(["strong", "b"])
        img = a.find("img")
        name = clean(strong.get_text()) if strong else ""
        if not name and img is not None:
            name = clean(img.get("alt", ""))
        if not name:
            name = clean(a.get_text(" "))
        if not name:
            continue

        # Photo
        photo = ""
        if img is not None:
            src = img.get("data-src") or img.get("data-original") or img.get("src") or ""
            photo = urljoin(page_url, src) if src else ""

        # Price: look in the anchor, then its parent block
        price, currency = "", ""
        for node in (a, a.parent, a.parent.parent if a.parent else None):
            if node is None:
                continue
            m = PRICE_RE.search(clean(node.get_text(" ")))
            if m:
                currency = m.group(1)
                price = m.group(2).replace(",", "")
                break

        year_m = YEAR_RE.search(name)
        year = year_m.group(1) if year_m else ""
        name_no_year = clean(YEAR_RE.sub("", name))
        brand, model = split_brand(name_no_year)
        section = urlparse(href).path.split("/")[2]

        out.append({
            "Country": COUNTRY,
            "Model Name": model or name_no_year,
            "Price": price,
            "Currency": currency,
            "Photo": photo,
            "Year": year,
            "Category": CATEGORY_MAP.get(section, section),
            "Brand": brand,
            "URL": href,
        })
    return out, soup


def find_more_pages(soup, page_url):
    """Follow pagination / brand listing links that stay under /br/."""
    found = set()
    for a in soup.find_all("a", href=True):
        href = urljoin(page_url, a["href"]).split("#")[0]
        p = urlparse(href)
        if p.netloc.endswith("fasterwheeler.com") and p.path.startswith("/br/"):
            if re.search(r"[?&]page=\d+|/page/\d+|/p/\d+|-bikes/?$|/brand", href, re.I):
                found.add(href)
    return found


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--debug", action="store_true")
    ap.add_argument("--max-pages", type=int, default=200)
    args = ap.parse_args()

    session = requests.Session()
    queue, seen_pages, products = list(SEED_PAGES), set(), {}

    while queue and len(seen_pages) < args.max_pages:
        url = queue.pop(0)
        if url in seen_pages:
            continue
        seen_pages.add(url)
        html = get(session, url)
        if html is None:
            if args.debug:
                print(f"[skip] {url} (not available)")
            continue
        items, soup = parse_products(html, url)
        new = 0
        for it in items:
            if it["URL"] not in products:
                products[it["URL"]] = it
                new += 1
        if args.debug:
            print(f"[page] {url}: {len(items)} products, {new} new")
        queue.extend(u for u in find_more_pages(soup, url) if u not in seen_pages)
        time.sleep(0.5)

    cols = ["Country", "Model Name", "Price", "Currency", "Photo", "Year",
            "Category", "Brand", "URL"]
    df = pd.DataFrame(list(products.values()), columns=cols)
    df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"Saved {OUTPUT_CSV}: {len(df)} products")
    if len(df):
        print(df.head(10).to_string())


if __name__ == "__main__":
    main()