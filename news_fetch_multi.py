"""
MotoPulse - multi-country two-wheeler news fetcher and categoriser (test version)

What it does (for each country you select)
  1. Searches Google News RSS (Spanish) for each of that country's brands and for general market topics
  2. Keeps only articles that clearly mention a brand AND two-wheelers
  3. Removes duplicate stories (inside each country)
  4. Tags each article with a brand and one of the 7 topics (simple keyword rules)
  5. Saves everything to one Excel file, with a Country column

How to run
  pip install requests openpyxl
  python news_fetch_multi.py --countries "Costa Rica,Peru" --days 90
  python news_fetch_multi.py --countries all --out latam_news.xlsx

Brands per country live in country_brands.csv (columns: Country, Brand, Aliases).
To add a brand, add a row. To add a country, add its rows AND one line in COUNTRY_SETTINGS below.
"""

import argparse
import csv
import os
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

# Per-country settings.
#   gl         = Google News country code
#   query_name = how the country is written inside the search text
COUNTRY_SETTINGS = {
    "Argentina":          {"gl": "AR", "query_name": "Argentina"},
    "Costa Rica":         {"gl": "CR", "query_name": "Costa Rica"},
    "Dominican Republic": {"gl": "DO", "query_name": "República Dominicana"},
    "El Salvador":        {"gl": "SV", "query_name": "El Salvador"},
    "Guatemala":          {"gl": "GT", "query_name": "Guatemala"},
    "Honduras":           {"gl": "HN", "query_name": "Honduras"},
    "Nicaragua":          {"gl": "NI", "query_name": "Nicaragua"},
    "Paraguay":           {"gl": "PY", "query_name": "Paraguay"},
    "Peru":               {"gl": "PE", "query_name": "Perú"},
    "Uruguay":            {"gl": "UY", "query_name": "Uruguay"},
}

# Google News edition. hl = language, gl = country, ceid = edition
GOOGLE_NEWS = "https://news.google.com/rss/search?q={query}&hl=es-419&gl={gl}&ceid={gl}:es-419"

BRANDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "country_brands.csv")

# Brand names written in capitals in the CSV are shown in Title Case, except these
KEEP_UPPER = {"TVS", "KTM", "AKT", "LML", "KMF", "ZMOTO"}

# Other ways a brand is written (used when the Aliases column of the CSV is empty)
DEFAULT_ALIASES = {
    "BAJAJ AUTO": ["Bajaj Auto", "Bajaj"],
    "HERO MOTOCORP": ["Hero MotoCorp"],
    "HARLEY DAVIDSON": ["Harley Davidson", "Harley-Davidson"],
}

# These names are also used for cars, football, films, common words etc., so for them the
# article MUST also contain a two-wheeler word (see TWO_WHEELER_WORDS) to be kept.
# Written in CAPITALS. Applies to every country.
AMBIGUOUS_BRANDS = {
    "HONDA", "YAMAHA", "SUZUKI", "KAWASAKI", "TVS", "BAJAJ", "BAJAJ AUTO", "PIAGGIO", "BETA",
    "TRIUMPH", "MONDIAL", "BENELLI", "MARUTI", "VENTO",
    # generic words used as brand names
    "VICTORY", "FORMULA", "FREEDOM", "STAR", "VITAL", "ALEX", "VINCE", "GENESIS", "NEXUS",
    "KATANA", "LEOPARD", "GACELA", "TAIGA", "KENTON", "RYDER", "YUMBO", "BLACK PANTHER",
    "SUPER TUCAN",
}

TWO_WHEELER_WORDS = ["moto", "motos", "motocicleta", "motocicletas", "motoneta", "scooter",
                     "ciclomotor", "motovehiculo", "motorcycle", "motorbike", "bike",
                     "cuatriciclo", "two-wheeler", "dos ruedas", "motero", "motoqueros"]

# General (no brand) searches so that market-wide news is also caught. {country} is filled in.
GENERAL_QUERY_TEMPLATES = [
    "motos {country} lanzamiento",
    "motos {country} precios",
    "motos {country} patentamientos",
    "motos eléctricas {country}",
    "fábrica de motos {country} producción",
    "concesionarias motos {country} nueva sucursal",
    "exportación motos {country}",
    "inversión motos {country}",
]

# Look-alike letters (e.g. a Cyrillic M inside "KYMCO") are replaced with the normal Latin letter
LOOKALIKES = str.maketrans({"\u041c": "M", "\u0410": "A", "\u0415": "E", "\u041e": "O", "\u0420": "P",
                            "\u0421": "C", "\u0422": "T", "\u041d": "H", "\u041a": "K", "\u0425": "X",
                            "\u0412": "B"})


def display_name(raw):
    """HONDA -> Honda, TVS -> TVS. Names that are not all capitals are kept as written."""
    if raw != raw.upper():
        return raw
    return raw if raw in KEEP_UPPER else raw.title()


def load_countries(path=BRANDS_FILE):
    """Read country_brands.csv -> {country: {gl, query_name, brands: {name: [aliases]}}}"""
    if not os.path.exists(path):
        raise SystemExit(f"Brand file not found: {path}")
    countries = {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            country = row["Country"].strip()
            raw = row["Brand"].strip().translate(LOOKALIKES)
            if not country or not raw:
                continue
            if country not in COUNTRY_SETTINGS:
                raise SystemExit(f"'{country}' is in {os.path.basename(path)} but not in COUNTRY_SETTINGS")
            name = display_name(raw)
            aliases = [a.strip() for a in (row.get("Aliases") or "").split("|") if a.strip()]
            aliases = aliases or DEFAULT_ALIASES.get(raw.upper(), [name])
            entry = countries.setdefault(country, {**COUNTRY_SETTINGS[country], "brands": {}})
            entry["brands"][name] = aliases
    return countries


COUNTRIES = load_countries()

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


def fetch_feed(query, gl, retries=2):
    """Download one Google News RSS search and return the raw XML text (or None if it failed)."""
    url = GOOGLE_NEWS.format(query=quote_plus(query), gl=gl)
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


def find_brands(cleaned_text, brands):
    """Return the list of brands mentioned in the text."""
    found = []
    for brand, aliases in brands.items():
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


def judge_article(article, brands):
    """Decide whether to keep an article. Returns (keep, reason, brands, cleaned_text)."""
    cleaned = clean_text(article["title"] + " " + article["snippet"])
    brands = find_brands(cleaned, brands)
    has_two_wheeler = any(has_keyword(cleaned, w) for w in TWO_WHEELER_WORDS)

    if not brands:
        # General market news: keep only if it is clearly about two-wheelers
        if has_two_wheeler:
            return True, "", ["General / Market"], cleaned
        return False, "No brand and no two-wheeler word", [], cleaned

    # If every brand found is an ambiguous name, a two-wheeler word is required
    if all(b.upper() in AMBIGUOUS_BRANDS for b in brands) and not has_two_wheeler:
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


def build_queries(country):
    """List of (search text, label) to run for one country. Each brand gets one search."""
    cfg = COUNTRIES[country]
    name = cfg["query_name"]
    queries = []
    for brand, aliases in cfg["brands"].items():
        queries.append((f'"{aliases[0]}" (moto OR motos OR motocicleta OR scooter) {name}', brand))
    for template in GENERAL_QUERY_TEMPLATES:
        queries.append((template.format(country=name), "General / Market"))
    return queries


def collect(countries, days_back, from_file=None):
    cutoff = datetime.now(timezone.utc) - timedelta(days=days_back)
    kept, rejected = [], []
    stats = Counter()
    country_stats = {}

    for country in countries:
        cfg = COUNTRIES[country]
        cs = country_stats[country] = Counter()
        seen_links, seen_titles = set(), []      # duplicates are checked inside one country

        def bump(key):
            stats[key] += 1
            cs[key] += 1

        # Test mode: read one saved XML file instead of calling Google
        if from_file:
            with open(from_file, encoding="utf-8") as f:
                jobs = [("(file)", "(file)", f.read())]
        else:
            jobs = []
            for query, label in build_queries(country):
                print(f"[{country}] Searching: {label:<18} | {query}")
                jobs.append((query, label, fetch_feed(query, cfg["gl"])))
                time.sleep(PAUSE_SECONDS)

        for query, label, xml_text in jobs:
            if xml_text is None:
                bump("searches failed")
                continue
            for article in parse_feed(xml_text):
                bump("articles seen")
                row_base = {"Search": query, "Headline": article["title"], "Source": article["source"],
                            "Published": article["published"], "Link": article["link"]}

                if article["published"] and article["published"] < cutoff:
                    bump("too old")
                    continue
                if article["link"] in seen_links:
                    bump("duplicate link")
                    continue
                seen_links.add(article["link"])

                keep, reason, brands, cleaned = judge_article(article, cfg["brands"])
                if not keep:
                    rejected.append({**row_base, "Country": country, "Reason rejected": reason})
                    bump("rejected")
                    continue

                cleaned_title = clean_text(article["title"])
                if is_duplicate(cleaned_title, seen_titles):
                    rejected.append({**row_base, "Country": country,
                                     "Reason rejected": "Duplicate story (similar headline)"})
                    bump("duplicate story")
                    continue
                seen_titles.append(cleaned_title)

                topic, all_topics = classify_topic(cleaned)
                kept.append({**row_base, "Country": country, "Brand": ", ".join(brands),
                             "Topic": topic, "All matching topics": all_topics})
                bump("kept")

        print(f"[{country}] done: seen {cs['articles seen']}, kept {cs['kept']}, "
              f"failed searches {cs['searches failed']}")

    kept.sort(key=lambda r: r["Published"] or datetime.min.replace(tzinfo=timezone.utc), reverse=True)
    return kept, rejected, stats, country_stats


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


def save_excel(kept, rejected, stats, country_stats, path):
    wb = Workbook()

    # Sheet 1: the clean news list
    ws = wb.active
    ws.title = "News"
    cols = ["Published", "Country", "Brand", "Topic", "Headline", "Source", "Link", "All matching topics", "Search"]
    write_header(ws, cols, [13, 18, 18, 30, 70, 22, 40, 45, 40])
    add_rows(ws, kept, cols)

    # Sheet 2: count of articles by country, brand and topic
    ws2 = wb.create_sheet("Summary")
    topics = list(TOPICS) + [OTHER_TOPIC]
    write_header(ws2, ["Country", "Brand"] + topics + ["Total"], [18, 20] + [16] * len(topics) + [10])
    counts = Counter()
    for r in kept:
        for brand in r["Brand"].split(", "):
            counts[(r["Country"], brand, r["Topic"])] += 1
    for country, brand in sorted({(c, b) for c, b, _ in counts}):
        row = [counts[(country, brand, t)] for t in topics]
        ws2.append([country, brand] + row + [sum(row)])
    ws2.row_dimensions[1].height = 45
    for cell in ws2[1]:
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    # Sheet 3: what was thrown away and why (helps the team tune the rules)
    ws3 = wb.create_sheet("Rejected")
    cols3 = ["Country", "Reason rejected", "Headline", "Source", "Published", "Link", "Search"]
    write_header(ws3, cols3, [18, 35, 70, 22, 13, 40, 40])
    add_rows(ws3, rejected, cols3)

    # Sheet 4: run details, overall and per country
    ws4 = wb.create_sheet("Run info")
    write_header(ws4, ["Item", "Value"], [34, 30])
    ws4.append(["Run time (UTC)", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M")])
    for key, value in stats.items():
        ws4.append([key, value])
    for country, cs in country_stats.items():
        ws4.append([])
        ws4.append([f"--- {country} ---", ""])
        for key, value in cs.items():
            ws4.append([f"{country}: {key}", value])

    wb.save(path)


def main():
    parser = argparse.ArgumentParser(description="Fetch and categorise two-wheeler news for several countries")
    parser.add_argument("--countries", default="all",
                        help='comma-separated, e.g. "Costa Rica,Peru" (default: all countries in country_brands.csv)')
    parser.add_argument("--days", type=int, default=DAYS_BACK_DEFAULT, help="how many days back to keep")
    parser.add_argument("--out", default="latam_news.xlsx", help="Excel file to create")
    parser.add_argument("--from-file", help="test mode: read a saved RSS xml file instead of calling Google")
    args = parser.parse_args()

    if args.countries.strip().lower() == "all":
        selected = list(COUNTRIES)
    else:
        lookup = {c.lower(): c for c in COUNTRIES}
        selected = []
        for name in args.countries.split(","):
            key = name.strip().lower()
            if key not in lookup:
                raise SystemExit(f"Unknown country '{name.strip()}'. Available: {', '.join(COUNTRIES)}")
            selected.append(lookup[key])

    searches = sum(len(build_queries(c)) for c in selected)
    print(f"Countries: {', '.join(selected)}  |  about {searches} searches"
          f" (~{searches * PAUSE_SECONDS / 60:.0f} min of pauses)")

    kept, rejected, stats, country_stats = collect(selected, args.days, args.from_file)
    save_excel(kept, rejected, stats, country_stats, args.out)

    print("\nDone.", dict(stats))
    print(f"Saved {len(kept)} articles to {args.out}")


if __name__ == "__main__":
    main()
