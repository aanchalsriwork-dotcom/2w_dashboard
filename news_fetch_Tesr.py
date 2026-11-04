"""
MotoPulse - Argentina two-wheeler news fetcher and categoriser (test version)

What it does
  1. Searches Google News RSS (Argentina, Spanish) for each brand and for general market topics
  2. Keeps only articles that clearly mention a brand AND two-wheelers
  3. Removes duplicate stories
  4. Tags each article with a brand and one of the 7 topics (simple keyword rules)
  5. Saves everything to an Excel file for the team to review

How to run
  pip install requests openpyxl
  python fetch_news.py                      (last 180 days, saves argentina_news.xlsx)
  python fetch_news.py --days 90 --out test.xlsx

To tune the results, edit the settings in the CONFIG section below.
No scheduling yet - this is run by hand for testing.
"""

import argparse
import difflib
import re
import time
import unicodedata
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from urllib.parse import quote_plus

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# =====================================================================
# CONFIG - edit this part to tune the results
# =====================================================================

COUNTRY = "Argentina"

# Google News settings for Argentina (Spanish). hl = language, gl = country, ceid = edition
GOOGLE_NEWS = "https://news.google.com/rss/search?q={query}&hl=es-419&gl=AR&ceid=AR:es-419"

# Brands to track. Left side = name shown in the Excel, right side = other ways it is written.
# This list is a starting guess - replace it with the brands you want to track.
BRANDS = {
    "Honda": ["Honda"],
    "Yamaha": ["Yamaha"],
    "Motomel": ["Motomel"],
    "Zanella": ["Zanella"],
    "Corven": ["Corven"],
    "Gilera": ["Gilera"],
    "Keeway": ["Keeway"],
    "Bajaj": ["Bajaj", "Rouser"],
    "Benelli": ["Benelli"],
    "Suzuki": ["Suzuki"],
    "TVS": ["TVS"],
    "Kawasaki": ["Kawasaki"],
    "KTM": ["KTM"],
    "Royal Enfield": ["Royal Enfield"],
    "Voge": ["Voge"],
    "CFMoto": ["CFMoto", "CF Moto"],
    "Mondial": ["Mondial"],
    "Beta": ["Beta Motor", "Beta Motos"],
    "BMW Motorrad": ["BMW Motorrad"],
    "Ducati": ["Ducati"],
    "Triumph": ["Triumph"],
}

# These names are also used for cars, pianos, appliances etc., so for them the article
# MUST also contain a two-wheeler word (see TWO_WHEELER_WORDS) to be kept.
AMBIGUOUS_BRANDS = {"Honda", "Yamaha", "Suzuki", "Bajaj", "TVS", "Kawasaki", "Beta",
                    "Triumph", "Mondial", "Benelli"}

TWO_WHEELER_WORDS = ["moto", "motos", "motocicleta", "motocicletas", "motoneta", "scooter",
                     "ciclomotor", "motovehiculo", "motorcycle", "motorbike", "bike",
                     "cuatriciclo", "two-wheeler", "dos ruedas", "motero", "motoqueros"]

# General (no brand) searches so that market-wide news is also caught
GENERAL_QUERIES = [
    "motos Argentina lanzamiento",
    "motos Argentina precios",
    "motos Argentina patentamientos",
    "motos eléctricas Argentina",
    "fábrica de motos Argentina producción",
    "concesionarias motos Argentina nueva sucursal",
    "exportación motos Argentina",
    "inversión motos Argentina",
]

# Topic rules. An article gets the topic with the most keyword hits.
# Keywords are written without accents and in lowercase (the text is cleaned the same way).
# If two topics tie, the one higher in this list wins.
TOPICS = {
    "New Launches": [
        "lanza", "lanzamiento", "lanzo", "presenta", "presento", "nuevo modelo", "nueva moto",
        "nuevas motos", "llega a la argentina", "llega al pais", "debut", "estreno",
        "novedad", "se suma a su gama", "launch", "unveil", "introduces", "nueva version",
    ],
    "Production & Capacity Changes": [
        "produccion", "fabrica", "planta", "capacidad", "ensamblado", "ensamble",
        "linea de montaje", "fabricacion", "produce", "turno", "paro de planta",
        "suspende produccion", "amplia planta", "manufactur", "factory", "plant",
    ],
    "Dealer Network Expansion": [
        "concesionaria", "concesionario", "sucursal", "red de venta", "red comercial",
        "nuevo local", "inaugura", "inauguracion", "abre", "apertura", "showroom",
        "puntos de venta", "dealer", "red de concesionarios", "distribuidor",
    ],
    "Price Changes": [
        "precio", "precios", "aumento", "suba", "sube", "baja de precio", "bajan",
        "descuento", "lista de precios", "cuanto cuesta", "cuesta", "valor", "cuotas",
        "financiacion", "tasa 0", "price", "oferta",
    ],
    "EV & Technology Announcements": [
        "electrica", "electricas", "electrico", "electricos", "bateria", "baterias",
        "autonomia", "carga rapida", "hibrida", "hibrido", "conectividad", "app",
        "tecnologia", "abs", "tft", "electric", "ev ", "litio", "kw",
    ],
    "Partnerships & Investments": [
        "alianza", "acuerdo", "convenio", "inversion", "invertira", "invierte", "socio",
        "joint venture", "asociacion", "adquiere", "adquisicion", "compra de", "fusion",
        "partnership", "investment", "millones de dolares", "financiamiento",
    ],
    "Exports & International Expansion": [
        "exporta", "exportacion", "exportaciones", "exportar", "mercado externo",
        "mercados regionales", "brasil", "chile", "uruguay", "paraguay", "latinoamerica",
        "region", "internacional", "importacion", "importaciones", "aranceles", "export",
    ],
}
OTHER_TOPIC = "Other / General"

DAYS_BACK_DEFAULT = 180      # ignore articles older than this
PAUSE_SECONDS = 1.5          # wait between requests so Google does not block us
SIMILAR_TITLE = 0.85         # titles this similar (0-1) are treated as the same story

# =====================================================================
# SMALL HELPER FUNCTIONS
# =====================================================================


def clean_text(text):
    """Lowercase, remove accents and punctuation, so keyword matching is simple."""
    text = unicodedata.normalize("NFKD", text or "")
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"<[^>]+>", " ", text)          # remove html tags
    text = re.sub(r"[^a-z0-9 ]+", " ", text.lower())
    return " " + re.sub(r"\s+", " ", text).strip() + " "   # spaces at both ends help whole-word matching


def has_word(cleaned_text, phrase):
    """True if the phrase appears as a whole word/phrase (so 'abs' does not match 'absoluto')."""
    return (" " + clean_text(phrase).strip() + " ") in cleaned_text


def has_keyword(cleaned_text, keyword):
    """Short keywords (4 letters or fewer, like 'abs' or 'app') must match a whole word.
    Longer ones match the start of a word, so 'fabrica' also finds 'fabricas'."""
    key = clean_text(keyword).strip()
    if len(key) <= 4:
        return (" " + key + " ") in cleaned_text
    return (" " + key) in cleaned_text


def fetch_feed(query, retries=2):
    """Download one Google News RSS search and return the raw XML text (or None if it failed)."""
    url = GOOGLE_NEWS.format(query=quote_plus(query))
    for attempt in range(retries + 1):
        try:
            response = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
            if response.status_code == 200:
                return response.text
            print(f"   ! HTTP {response.status_code} for '{query}' (try {attempt + 1})")
        except requests.RequestException as error:
            print(f"   ! Network problem for '{query}': {error}")
        time.sleep(PAUSE_SECONDS * 2)
    return None


def parse_feed(xml_text):
    """Turn RSS XML into a list of simple dictionaries (one per article)."""
    articles = []
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return articles
    for item in root.iter("item"):
        title = (item.findtext("title") or "").strip()
        source_el = item.find("source")
        source = source_el.text.strip() if source_el is not None and source_el.text else ""
        # Google adds " - Source name" at the end of the title, so remove it
        if source and title.endswith(" - " + source):
            title = title[: -len(source) - 3].strip()
        try:
            published = parsedate_to_datetime(item.findtext("pubDate"))
            if published.tzinfo is None:
                published = published.replace(tzinfo=timezone.utc)
        except (TypeError, ValueError):
            published = None
        articles.append({
            "title": title,
            "source": source,
            "link": (item.findtext("link") or "").strip(),
            "published": published,
            "snippet": re.sub(r"<[^>]+>", " ", item.findtext("description") or "").strip(),
        })
    return articles


def find_brands(cleaned_text):
    """Return the list of brands mentioned in the text."""
    found = []
    for brand, aliases in BRANDS.items():
        if any(has_word(cleaned_text, alias) for alias in aliases):
            found.append(brand)
    return found


def classify_topic(cleaned_text):
    """Return (main topic, all matching topics). Main topic = most keyword hits."""
    scores = {}
    for topic, keywords in TOPICS.items():
        hits = sum(1 for kw in keywords if has_keyword(cleaned_text, kw))
        if hits:
            scores[topic] = hits
    if not scores:
        return OTHER_TOPIC, ""
    best = max(scores.values())
    # TOPICS order decides ties because we loop in that order
    main = next(topic for topic in TOPICS if scores.get(topic) == best)
    others = [f"{t} ({n})" for t, n in sorted(scores.items(), key=lambda x: -x[1])]
    return main, ", ".join(others)


def judge_article(article):
    """Decide whether to keep an article. Returns (keep, reason, brands, cleaned_text)."""
    cleaned = clean_text(article["title"] + " " + article["snippet"])
    brands = find_brands(cleaned)
    has_two_wheeler = any(has_keyword(cleaned, w) for w in TWO_WHEELER_WORDS)

    if not brands:
        # General market news: keep only if it is clearly about two-wheelers
        if has_two_wheeler:
            return True, "", ["General / Market"], cleaned
        return False, "No brand and no two-wheeler word", [], cleaned

    # If every brand found is an ambiguous name, a two-wheeler word is required
    if all(b in AMBIGUOUS_BRANDS for b in brands) and not has_two_wheeler:
        return False, "Ambiguous brand without a two-wheeler word", brands, cleaned
    return True, "", brands, cleaned


def is_duplicate(cleaned_title, seen_titles):
    """True if a very similar title was already kept."""
    for old in seen_titles:
        if difflib.SequenceMatcher(None, cleaned_title, old).ratio() >= SIMILAR_TITLE:
            return True
    return False


# =====================================================================
# MAIN WORK
# =====================================================================


def build_queries():
    """List of (search text, label) to run. Each brand gets one search."""
    queries = []
    for brand, aliases in BRANDS.items():
        name = aliases[0]
        queries.append((f'"{name}" (moto OR motos OR motocicleta OR scooter) Argentina', brand))
    for text in GENERAL_QUERIES:
        queries.append((text, "General / Market"))
    return queries


def collect(days_back, from_file=None):
    cutoff = datetime.now(timezone.utc) - timedelta(days=days_back)
    kept, rejected = [], []
    seen_links, seen_titles = set(), []
    stats = Counter()

    # Test mode: read one saved XML file instead of calling Google
    if from_file:
        with open(from_file, encoding="utf-8") as f:
            jobs = [("(file)", "(file)", f.read())]
    else:
        jobs = []
        for query, label in build_queries():
            print(f"Searching: {label:<18} | {query}")
            jobs.append((query, label, fetch_feed(query)))
            time.sleep(PAUSE_SECONDS)

    for query, label, xml_text in jobs:
        if xml_text is None:
            stats["searches failed"] += 1
            continue
        for article in parse_feed(xml_text):
            stats["articles seen"] += 1
            row_base = {"Search": query, "Headline": article["title"], "Source": article["source"],
                        "Published": article["published"], "Link": article["link"]}

            if article["published"] and article["published"] < cutoff:
                stats["too old"] += 1
                continue
            if article["link"] in seen_links:
                stats["duplicate link"] += 1
                continue
            seen_links.add(article["link"])

            keep, reason, brands, cleaned = judge_article(article)
            if not keep:
                rejected.append({**row_base, "Reason rejected": reason})
                stats["rejected"] += 1
                continue

            cleaned_title = clean_text(article["title"])
            if is_duplicate(cleaned_title, seen_titles):
                rejected.append({**row_base, "Reason rejected": "Duplicate story (similar headline)"})
                stats["duplicate story"] += 1
                continue
            seen_titles.append(cleaned_title)

            topic, all_topics = classify_topic(cleaned)
            kept.append({**row_base, "Country": COUNTRY, "Brand": ", ".join(brands),
                         "Topic": topic, "All matching topics": all_topics})
            stats["kept"] += 1

    kept.sort(key=lambda r: r["Published"] or datetime.min.replace(tzinfo=timezone.utc), reverse=True)
    return kept, rejected, stats


def write_header(ws, headers, widths):
    ws.append(headers)
    for col, (head, width) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2563EB")
        cell.alignment = Alignment(vertical="center")
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"


def add_rows(ws, rows, columns):
    for r in rows:
        values = []
        for col in columns:
            v = r.get(col, "")
            if col == "Published" and v:
                v = v.astimezone(timezone.utc).replace(tzinfo=None)   # Excel cannot store time zones
            values.append(v)
        ws.append(values)
    # make the link column clickable and wrap long headlines
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if ws.cell(row=1, column=cell.column).value == "Link" and cell.value:
                cell.hyperlink = cell.value
                cell.font = Font(color="2563EB", underline="single")
            if ws.cell(row=1, column=cell.column).value == "Published" and cell.value:
                cell.number_format = "dd-mmm-yyyy"
    ws.auto_filter.ref = ws.dimensions


def save_excel(kept, rejected, stats, path):
    wb = Workbook()

    # Sheet 1: the clean news list
    ws = wb.active
    ws.title = "News"
    cols = ["Published", "Country", "Brand", "Topic", "Headline", "Source", "Link", "All matching topics", "Search"]
    write_header(ws, cols, [13, 11, 18, 30, 70, 22, 40, 45, 40])
    add_rows(ws, kept, cols)

    # Sheet 2: count of articles by brand and topic
    ws2 = wb.create_sheet("Summary")
    topics = list(TOPICS) + [OTHER_TOPIC]
    write_header(ws2, ["Brand"] + topics + ["Total"], [20] + [16] * len(topics) + [10])
    counts = Counter()
    for r in kept:
        for brand in r["Brand"].split(", "):
            counts[(brand, r["Topic"])] += 1
    for brand in sorted({b for b, _ in counts}):
        row = [counts[(brand, t)] for t in topics]
        ws2.append([brand] + row + [sum(row)])
    ws2.row_dimensions[1].height = 45
    for cell in ws2[1]:
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    # Sheet 3: what was thrown away and why (helps the team tune the rules)
    ws3 = wb.create_sheet("Rejected")
    cols3 = ["Reason rejected", "Headline", "Source", "Published", "Link", "Search"]
    write_header(ws3, cols3, [35, 70, 22, 13, 40, 40])
    add_rows(ws3, rejected, cols3)

    # Sheet 4: run details
    ws4 = wb.create_sheet("Run info")
    write_header(ws4, ["Item", "Value"], [28, 30])
    ws4.append(["Run time (UTC)", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M")])
    for key, value in stats.items():
        ws4.append([key, value])

    wb.save(path)


def main():
    parser = argparse.ArgumentParser(description="Fetch and categorise Argentina two-wheeler news")
    parser.add_argument("--days", type=int, default=DAYS_BACK_DEFAULT, help="how many days back to keep")
    parser.add_argument("--out", default="argentina_news.xlsx", help="Excel file to create")
    parser.add_argument("--from-file", help="test mode: read a saved RSS xml file instead of calling Google")
    args = parser.parse_args()

    kept, rejected, stats = collect(args.days, args.from_file)
    save_excel(kept, rejected, stats, args.out)

    print("\nDone.", dict(stats))
    print(f"Saved {len(kept)} articles to {args.out}")


if __name__ == "__main__":
    main()


                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            