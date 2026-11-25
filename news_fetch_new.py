"""
MotoPulse - GLOBAL two-wheeler news fetcher and categoriser (86 countries, 28 languages)

What it does (for each country you select)
  1. Searches Google News RSS in the country's own language(s) for its brands and for general market topics
  2. Keeps only articles that clearly mention a brand AND two-wheelers (words checked in the right language)
  3. Removes duplicate stories (inside each country)
  4. Tags each article with a brand and one of 7 topics (keyword rules per language)
  5. Saves everything to ONE Excel file (News, Summary, Rejected, Countries, Run info)

How to run
  pip install requests openpyxl deep-translator      (deep-translator only needed for --translate)
  python news_fetch_multi.py --probe --countries "Thailand,Germany,Egypt"     # quick check, no Excel
  python news_fetch_multi.py --countries "Philippines,Thailand" --days 90
  python news_fetch_multi.py --countries all --workers 3 --resume --out global_news.xlsx
  python news_fetch_multi.py --countries all --resume --translate     # adds a "Headline (English)" column

  --resume   re-uses countries already finished in news_cache/ (safe to stop with Ctrl+C and continue later)
  --translate  translates the headline of every KEPT non-English article to English (free Google Translate
               endpoint via deep-translator), caches translations in news_cache/translations.json, and uses
               the English text for topic tagging too
  --workers  countries fetched in parallel (default 3; use 1 if Google starts answering HTTP 429)

Brands: country_brands.csv (Country, Brand, Aliases) is OPTIONAL. If a country has rows in it, those
brands are used. Otherwise the built-in regional brand lists below are used - edit BRAND_SETS /
COUNTRY_EXTRA_BRANDS to tune them. To add a country, add one line in COUNTRY_SETTINGS.
"""

import argparse
import csv
import difflib
import json
import os
import re
import threading
import time
import unicodedata
import xml.etree.ElementTree as ET
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from urllib.parse import quote_plus

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# =====================================================================
# COUNTRIES
# =====================================================================
# C(gl, region, languages, **names)
#   gl        = Google News country code
#   languages = news languages to search in (first = main language)
#   names     = how the country is written in each language (English defaults to the country key)


def C(gl, region, langs=("en",), **names):
    return {"gl": gl, "region": region, "langs": list(langs), "names": names}


COUNTRY_SETTINGS = {
    # ---- ASEAN
    "Philippines": C("PH", "ASEAN"),
    "Malaysia": C("MY", "ASEAN"),
    "Thailand": C("TH", "ASEAN", ["th"], th="ประเทศไทย"),
    "Vietnam": C("VN", "ASEAN", ["vi"], vi="Việt Nam"),
    "Indonesia": C("ID", "ASEAN", ["id"], id="Indonesia"),
    "Cambodia": C("KH", "ASEAN"),
    "Myanmar": C("MM", "ASEAN"),
    "Singapore": C("SG", "ASEAN"),
    # ---- Other Asia / Oceania
    "Australia": C("AU", "Other Asia"),
    "New Zealand": C("NZ", "Other Asia"),
    "Bangladesh": C("BD", "Other Asia"),
    "Nepal": C("NP", "Other Asia"),
    "Sri Lanka": C("LK", "Other Asia"),
    # ---- Europe
    "Austria": C("AT", "Europe", ["de"], de="Österreich"),
    "United Kingdom": C("GB", "Europe", en="UK"),
    "Belgium": C("BE", "Europe", ["nl", "fr"], nl="België", fr="Belgique"),
    "Russia": C("RU", "Europe", ["ru"], ru="Россия"),
    "Switzerland": C("CH", "Europe", ["de", "fr"], de="Schweiz", fr="Suisse"),
    "Poland": C("PL", "Europe", ["pl"], pl="Polska"),
    "Estonia": C("EE", "Europe", ["et"], et="Eesti"),
    "Sweden": C("SE", "Europe", ["sv"], sv="Sverige"),
    "Finland": C("FI", "Europe", ["fi"], fi="Suomi"),
    "Norway": C("NO", "Europe", ["no"], no="Norge"),
    "Denmark": C("DK", "Europe", ["da"], da="Danmark"),
    "Romania": C("RO", "Europe", ["ro"], ro="România"),
    "Czech Republic": C("CZ", "Europe", ["cs"], cs="Česko"),
    "France": C("FR", "Europe", ["fr"], fr="France"),
    "Germany": C("DE", "Europe", ["de"], de="Deutschland"),
    "Greece": C("GR", "Europe", ["el"], el="Ελλάδα"),
    "Hungary": C("HU", "Europe", ["hu"], hu="Magyarország"),
    "Ukraine": C("UA", "Europe", ["uk"], uk="Україна"),
    "Ireland": C("IE", "Europe"),
    "Bulgaria": C("BG", "Europe", ["bg"], bg="България"),
    "Croatia": C("HR", "Europe", ["hr"], hr="Hrvatska"),
    "Italy": C("IT", "Europe", ["it"], it="Italia"),
    "Belarus": C("BY", "Europe", ["ru"], ru="Беларусь"),
    "Latvia": C("LV", "Europe", ["lv"], lv="Latvija"),
    "Luxembourg": C("LU", "Europe", ["fr", "de"], fr="Luxembourg", de="Luxemburg"),
    "Netherlands": C("NL", "Europe", ["nl"], nl="Nederland"),
    "Portugal": C("PT", "Europe", ["pt"], pt="Portugal"),
    "Slovakia": C("SK", "Europe", ["sk"], sk="Slovensko"),
    "Slovenia": C("SI", "Europe", ["sl"], sl="Slovenija"),
    "Spain": C("ES", "Europe", ["es"], es="España"),
    # ---- Africa
    "Cameroon": C("CM", "Africa", ["fr", "en"], fr="Cameroun"),
    "DR Congo": C("CD", "Africa", ["fr"], fr="RD Congo", en="DR Congo"),
    "Ivory Coast": C("CI", "Africa", ["fr"], fr="Côte d'Ivoire", en="Ivory Coast"),
    "Nigeria": C("NG", "Africa"),
    "Kenya": C("KE", "Africa"),
    "Tanzania": C("TZ", "Africa"),
    "Ghana": C("GH", "Africa"),
    "Uganda": C("UG", "Africa"),
    "Angola": C("AO", "Africa", ["pt"], pt="Angola"),
    "Egypt": C("EG", "Africa", ["ar", "en"], ar="مصر"),
    "Ethiopia": C("ET", "Africa"),
    "Niger": C("NE", "Africa", ["fr"], fr="Niger"),
    "Liberia": C("LR", "Africa"),
    "Morocco": C("MA", "Africa", ["fr", "ar"], fr="Maroc", ar="المغرب"),
    "Mozambique": C("MZ", "Africa", ["pt"], pt="Moçambique"),
    "Namibia": C("NA", "Africa"),
    "Rwanda": C("RW", "Africa", ["en", "fr"], fr="Rwanda"),
    "South Sudan": C("SS", "Africa"),
    "Zimbabwe": C("ZW", "Africa"),
    # ---- Middle East
    "KSA": C("SA", "Middle East", ["en", "ar"], en="Saudi Arabia", ar="السعودية"),
    "UAE": C("AE", "Middle East", ["en", "ar"], ar="الإمارات"),
    "Kuwait": C("KW", "Middle East", ["en", "ar"], ar="الكويت"),
    "Bahrain": C("BH", "Middle East", ["en", "ar"], ar="البحرين"),
    "Qatar": C("QA", "Middle East", ["en", "ar"], ar="قطر"),
    "Oman": C("OM", "Middle East", ["en", "ar"], ar="عمان"),
    # ---- LATAM
    "Ecuador": C("EC", "LATAM", ["es"], es="Ecuador"),
    "Argentina": C("AR", "LATAM", ["es"], es="Argentina"),
    "Bolivia": C("BO", "LATAM", ["es"], es="Bolivia"),
    "Brazil": C("BR", "LATAM", ["pt"], pt="Brasil"),
    "Chile": C("CL", "LATAM", ["es"], es="Chile"),
    "Colombia": C("CO", "LATAM", ["es"], es="Colombia"),
    "Costa Rica": C("CR", "LATAM", ["es"], es="Costa Rica"),
    "Dominican Republic": C("DO", "LATAM", ["es"], es="República Dominicana"),
    "El Salvador": C("SV", "LATAM", ["es"], es="El Salvador"),
    "Guatemala": C("GT", "LATAM", ["es"], es="Guatemala"),
    "Honduras": C("HN", "LATAM", ["es"], es="Honduras"),
    "Nicaragua": C("NI", "LATAM", ["es"], es="Nicaragua"),
    "Paraguay": C("PY", "LATAM", ["es"], es="Paraguay"),
    "Peru": C("PE", "LATAM", ["es"], es="Perú"),
    "Uruguay": C("UY", "LATAM", ["es"], es="Uruguay"),
    "Mexico": C("MX", "LATAM", ["es"], es="México"),
    # ---- North America
    "USA": C("US", "North America", en="USA"),
    "Canada": C("CA", "North America", ["en", "fr"], fr="Canada"),
}

# Other spellings accepted by --countries
COUNTRY_ALIASES = {
    "saudi arabia": "KSA", "united states": "USA", "us": "USA", "uk": "United Kingdom",
    "drc": "DR Congo", "congo": "DR Congo", "cote d'ivoire": "Ivory Coast", "czechia": "Czech Republic",
    "united arab emirates": "UAE", "britain": "United Kingdom",
}

# =====================================================================
# BRANDS (used for any country that has no rows in country_brands.csv)
# =====================================================================
BRANDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "country_brands.csv")

BRAND_SETS = {
    "ASEAN": ["Honda", "Yamaha", "Suzuki", "Kawasaki", "Vespa", "Piaggio", "SYM", "Kymco", "Benelli",
              "CFMoto", "Royal Enfield", "VinFast", "Yadea", "TVS", "KTM", "Harley-Davidson"],
    "Other Asia": ["Honda", "Yamaha", "Suzuki", "Kawasaki", "Bajaj", "TVS", "Hero MotoCorp", "Royal Enfield",
                   "KTM", "Harley-Davidson", "Triumph", "BMW Motorrad", "Ducati", "Vespa", "Yadea", "CFMoto"],
    "Europe": ["Honda", "Yamaha", "Suzuki", "Kawasaki", "BMW Motorrad", "Ducati", "KTM", "Triumph", "Piaggio",
               "Vespa", "Harley-Davidson", "Royal Enfield", "CFMoto", "Aprilia", "Husqvarna", "Zero Motorcycles"],
    "Africa": ["Honda", "Yamaha", "Suzuki", "Bajaj", "TVS", "Hero MotoCorp", "Haojue", "Kawasaki", "KTM",
               "Royal Enfield", "Spiro", "Ampersand", "Roam"],
    "Middle East": ["Honda", "Yamaha", "Suzuki", "Kawasaki", "Harley-Davidson", "BMW Motorrad", "Ducati", "KTM",
                    "Triumph", "Royal Enfield", "Bajaj", "TVS", "CFMoto", "Benelli", "Vespa"],
    "LATAM": ["Honda", "Yamaha", "Suzuki", "Bajaj", "TVS", "Hero MotoCorp", "KTM", "Royal Enfield", "Kawasaki",
              "Benelli", "CFMoto", "Harley-Davidson", "Triumph", "BMW Motorrad"],
    "North America": ["Harley-Davidson", "Indian Motorcycle", "Honda", "Yamaha", "Suzuki", "Kawasaki",
                      "BMW Motorrad", "Ducati", "KTM", "Triumph", "Royal Enfield", "Zero Motorcycles",
                      "LiveWire", "Can-Am", "Vespa", "Segway"],
}

# Local brands searched FIRST for these countries (added in front of the regional list)
COUNTRY_EXTRA_BRANDS = {
    "Philippines": ["Rusi", "Motorstar", "Skygo"], "Malaysia": ["Modenas"], "Thailand": ["GPX"],
    "Vietnam": ["Dat Bike"], "Indonesia": ["Viar", "Polytron"], "Bangladesh": ["Runner", "Walton"],
    "Nepal": ["Yatri"], "Russia": ["Motoland", "Avantis", "Irbis"], "Poland": ["Junak", "Romet"],
    "Italy": ["Moto Guzzi", "MV Agusta"], "Spain": ["Silence"], "Mexico": ["Italika", "Vento"],
    "Brazil": ["Haojue", "Shineray", "Dafra", "Voltz"], "Colombia": ["AKT", "Auteco"], "Chile": ["Voge", "Loncin"],
}

# Other ways a brand is written
BRAND_ALIASES = {
    "Harley-Davidson": ["Harley-Davidson", "Harley Davidson"],
    "BMW Motorrad": ["BMW Motorrad", "BMW"],
    "Can-Am": ["Can-Am", "Can Am"],
    "Bajaj": ["Bajaj Auto", "Bajaj"],
    "Hero MotoCorp": ["Hero MotoCorp"],
}
DEFAULT_ALIASES = {      # used for CSV rows with an empty Aliases column
    "BAJAJ AUTO": ["Bajaj Auto", "Bajaj"],
    "HERO MOTOCORP": ["Hero MotoCorp"],
    "HARLEY DAVIDSON": ["Harley Davidson", "Harley-Davidson"],
}

# Brand names in other alphabets (Russian/Ukrainian/Bulgarian, Thai, Arabic, Greek)
NATIVE_ALIASES = {
    "Honda": ["Хонда", "ฮอนด้า", "هوندا", "Χόντα"],
    "Yamaha": ["Ямаха", "ยามาฮ่า", "ياماها", "Γιαμάχα"],
    "Suzuki": ["Сузуки", "ซูซูกิ", "سوزوكي", "Σουζούκι"],
    "Kawasaki": ["Кавасаки", "คาวาซากิ", "كاواساكي", "Καουασάκι"],
    "Harley-Davidson": ["Харли-Дэвидсон", "ฮาร์ลีย์-เดวิดสัน", "هارلي ديفيدسون"],
    "Ducati": ["Дукати", "ดูคาติ", "دوكاتي"],
    "Triumph": ["Триумф", "ไทรอัมพ์", "تريومف"],
    "Vespa": ["Веспа", "เวสป้า", "فيسبا"],
    "Royal Enfield": ["Роял Энфилд", "รอยัล เอนฟิลด์", "رويال إنفيلد"],
    "Piaggio": ["Пиаджо", "เปียจิโอ", "بياجيو"],
}

KEEP_UPPER = {"TVS", "KTM", "AKT", "LML", "KMF", "ZMOTO"}

# These names are also used for cars, football, tools, common words etc., so for them the article
# MUST also contain a two-wheeler word to be kept. Written in CAPITALS. Applies to every country.
AMBIGUOUS_BRANDS = {
    "HONDA", "YAMAHA", "SUZUKI", "KAWASAKI", "TVS", "BAJAJ", "BAJAJ AUTO", "PIAGGIO", "BETA",
    "TRIUMPH", "MONDIAL", "BENELLI", "MARUTI", "VENTO",
    "VESPA", "SYM", "BMW MOTORRAD", "CAN-AM", "SEGWAY", "ROAM", "SPIRO", "AMPERSAND", "RUSI", "RACER",
    "HUSQVARNA", "VINFAST", "POLYTRON", "WALTON", "RUNNER", "ROMET", "LIVEWIRE", "VOLTZ",
    # generic words used as brand names
    "VICTORY", "FORMULA", "FREEDOM", "STAR", "VITAL", "ALEX", "VINCE", "GENESIS", "NEXUS",
    "KATANA", "LEOPARD", "GACELA", "TAIGA", "KENTON", "RYDER", "YUMBO", "BLACK PANTHER",
    "SUPER TUCAN",
}

# =====================================================================
# LANGUAGE PACKS
#   wheel  = words meaning motorcycle / scooter / moped (used to filter articles; first 3 are single
#            words and also go into the brand searches)
#   q      = general market searches ({c} = country name in that language)
#   topics = keywords per topic (optional; English keywords are always added on top)
# Keywords are matched without accents, lowercase. Words of 4 letters or fewer must match a whole word;
# longer ones match the start of a word. Thai and Arabic are matched anywhere inside the text.
# =====================================================================
TOPIC_NAMES = ["New Launches", "Production & Capacity Changes", "Dealer Network Expansion", "Price Changes",
               "EV & Technology Announcements", "Partnerships & Investments", "Exports & International Expansion"]
OTHER_TOPIC = "Other / General"


def T(*groups):
    return dict(zip(TOPIC_NAMES, groups))


LANGS = {
    "en": {
        "wheel": ["motorcycle", "motorbike", "scooter", "motorcycles", "motorbikes", "moped", "scooters",
                  "two-wheeler", "two-wheelers", "2-wheeler", "e-motorcycle", "boda boda", "bodaboda", "okada",
                  "dirt bike", "superbike"],
        "q": ["motorcycle launch {c}", "motorcycle sales {c}", "motorcycle prices {c}", "electric motorcycle {c}",
              "motorcycle factory production {c}", "motorcycle dealership new showroom {c}",
              "motorcycle exports {c}", "motorcycle investment {c}", "scooter {c} launch"],
        "topics": T(
            ["launch", "launches", "launched", "unveil", "unveils", "unveiled", "debut", "introduces", "introduced",
             "new model", "all-new", "revealed", "goes on sale", "arrives in", "new version"],
            ["production", "factory", "plant", "assembly", "manufacturing", "manufacture", "capacity",
             "production line", "assembles", "localisation", "localization"],
            ["dealer", "dealers", "dealership", "showroom", "outlet", "new store", "opens", "network expansion",
             "service centre", "distributor", "inaugurates"],
            ["price", "prices", "price hike", "discount", "price cut", "starting at", "cost", "offer",
             "financing", "emi", "on-road"],
            ["electric", "ev", "battery", "range", "fast charging", "hybrid", "connectivity", "swap", "charging",
             "tft", "abs", "kw", "lithium"],
            ["partnership", "tie-up", "agreement", "investment", "invests", "joint venture", "acquires",
             "acquisition", "stake", "funding", "collaboration", "merger", "mou"],
            ["export", "exports", "exporting", "import", "imports", "tariff", "tariffs", "international expansion",
             "overseas", "enters market", "new market", "global expansion"]),
    },
    "es": {
        "wheel": ["moto", "motos", "motocicleta", "motocicletas", "motoneta", "scooter", "ciclomotor",
                  "motovehiculo", "cuatriciclo", "dos ruedas", "motero", "motoqueros"],
        "q": ["motos {c} lanzamiento", "motos {c} precios", "motos {c} patentamientos", "motos eléctricas {c}",
              "fábrica de motos {c} producción", "concesionarias motos {c} nueva sucursal",
              "exportación motos {c}", "inversión motos {c}"],
        "topics": T(
            ["lanza", "lanzamiento", "lanzo", "presenta", "presento", "nuevo modelo", "nueva moto", "nuevas motos",
             "llega a la argentina", "llega al pais", "debut", "estreno", "novedad", "se suma a su gama",
             "nueva version"],
            ["produccion", "fabrica", "planta", "capacidad", "ensamblado", "ensamble", "linea de montaje",
             "fabricacion", "produce", "turno", "paro de planta", "suspende produccion", "amplia planta"],
            ["concesionaria", "concesionario", "sucursal", "red de venta", "red comercial", "nuevo local",
             "inaugura", "inauguracion", "abre", "apertura", "showroom", "puntos de venta",
             "red de concesionarios", "distribuidor"],
            ["precio", "precios", "aumento", "suba", "sube", "baja de precio", "bajan", "descuento",
             "lista de precios", "cuanto cuesta", "cuesta", "valor", "cuotas", "financiacion", "tasa 0", "oferta"],
            ["electrica", "electricas", "electrico", "electricos", "bateria", "baterias", "autonomia",
             "carga rapida", "hibrida", "hibrido", "conectividad", "app", "tecnologia", "abs", "tft", "litio"],
            ["alianza", "acuerdo", "convenio", "inversion", "invertira", "invierte", "socio", "joint venture",
             "asociacion", "adquiere", "adquisicion", "compra de", "fusion", "millones de dolares",
             "financiamiento"],
            ["exporta", "exportacion", "exportaciones", "exportar", "mercado externo", "mercados regionales",
             "latinoamerica", "region", "internacional", "importacion", "importaciones", "aranceles"]),
    },
    "fr": {
        "wheel": ["moto", "motos", "scooter", "motocyclette", "cyclomoteur", "deux-roues", "2 roues", "scooters"],
        "q": ["motos {c} lancement", "marché moto {c} ventes", "immatriculations motos {c}", "moto électrique {c}",
              "prix motos {c}", "usine motos production {c}", "concessionnaire moto {c}"],
        "topics": T(
            ["lancement", "lance", "presente", "presentation", "nouveaute", "nouveau modele", "nouvelle moto",
             "devoile", "debut"],
            ["production", "usine", "fabrication", "assemblage", "capacite", "ligne de production", "fabrique"],
            ["concessionnaire", "concession", "reseau", "point de vente", "showroom", "inaugure", "ouverture",
             "distributeur"],
            ["prix", "tarif", "tarifs", "hausse", "baisse", "remise", "promotion", "cout", "financement"],
            ["electrique", "electriques", "batterie", "autonomie", "recharge", "hybride", "technologie", "abs", "kw"],
            ["partenariat", "accord", "investissement", "investit", "co-entreprise", "acquisition", "rachat",
             "fusion", "alliance"],
            ["exportation", "exporte", "exportations", "importation", "droits de douane", "marche international",
             "expansion internationale"]),
    },
    "pt": {
        "wheel": ["moto", "motos", "motocicleta", "motocicletas", "scooter", "ciclomotor", "duas rodas"],
        "q": ["motos {c} lançamento", "mercado de motos {c} vendas", "emplacamentos motos {c}",
              "motos elétricas {c}", "preços motos {c}", "fábrica de motos {c}", "concessionária motos {c}"],
        "topics": T(
            ["lancamento", "lanca", "lancou", "apresenta", "apresentou", "nova moto", "novidade", "estreia",
             "nova versao"],
            ["producao", "fabrica", "planta", "capacidade", "montagem", "fabricacao", "linha de montagem"],
            ["concessionaria", "concessionarias", "revenda", "rede de vendas", "inaugura", "inauguracao",
             "abertura", "loja", "showroom", "distribuidor"],
            ["preco", "precos", "aumento", "reajuste", "desconto", "tabela de precos", "quanto custa",
             "financiamento", "consorcio", "promocao"],
            ["eletrica", "eletricas", "eletrico", "bateria", "autonomia", "recarga", "hibrida", "conectividade",
             "tecnologia", "abs"],
            ["parceria", "acordo", "investimento", "investe", "joint venture", "aquisicao", "adquire", "fusao",
             "alianca"],
            ["exporta", "exportacao", "exportacoes", "importacao", "importacoes", "tarifas", "mercado externo",
             "internacional"]),
    },
    "de": {
        "wheel": ["motorrad", "motorräder", "motorroller", "roller", "moped", "zweirad", "elektromotorrad"],
        "q": ["Motorrad {c}", "Motorrad Neuzulassungen {c}", "neues Motorrad {c}", "Elektromotorrad {c}",
              "Motorrad Preise {c}", "Motorradhändler {c}", "Motorrad Produktion {c}"],
        "topics": T(
            ["vorgestellt", "neuheit", "neues modell", "markteinführung", "enthüllt", "premiere", "neuvorstellung"],
            ["produktion", "werk", "fabrik", "fertigung", "kapazität", "montage", "produktionsstopp"],
            ["händler", "händlernetz", "motorradhändler", "filiale", "eröffnet", "eröffnung", "neuer standort",
             "vertrieb"],
            ["preis", "preise", "preiserhöhung", "rabatt", "preissenkung", "kosten", "angebot", "finanzierung"],
            ["elektro", "elektrisch", "elektromotorrad", "akku", "batterie", "reichweite", "schnellladen",
             "hybrid", "technologie", "abs"],
            ["partnerschaft", "kooperation", "investition", "investiert", "übernahme", "joint venture",
             "beteiligung", "fusion"],
            ["export", "exporte", "import", "zölle", "auslandsmarkt", "internationalisierung", "expansion"]),
    },
    "it": {
        "wheel": ["moto", "motocicletta", "motociclette", "scooter", "motorino", "ciclomotore"],
        "q": ["moto {c} novità", "mercato moto {c}", "immatricolazioni moto {c}", "moto elettriche {c}",
              "prezzi moto {c}", "produzione moto {c}"],
        "topics": T(
            ["lancio", "lancia", "presenta", "presentata", "nuovo modello", "nuova moto", "novita", "debutto",
             "arriva in", "svelata"],
            ["produzione", "fabbrica", "stabilimento", "capacita", "assemblaggio", "linea di montaggio"],
            ["concessionario", "concessionaria", "rete vendita", "nuova sede", "inaugura", "apertura", "showroom",
             "distributore"],
            ["prezzo", "prezzi", "aumento", "rincaro", "sconto", "listino", "costa", "finanziamento", "offerta",
             "promozione"],
            ["elettrica", "elettriche", "elettrico", "batteria", "autonomia", "ricarica", "ibrida", "connettivita",
             "tecnologia", "abs"],
            ["partnership", "accordo", "investimento", "investe", "joint venture", "acquisizione", "acquista",
             "fusione", "alleanza"],
            ["esporta", "esportazione", "esportazioni", "importazione", "dazi", "mercati esteri", "internazionale"]),
    },
    "nl": {
        "wheel": ["motorfiets", "motorfietsen", "scooter", "scooters", "bromfiets", "brommer"],
        "q": ["motoren {c}", "motorfietsen {c} verkoop", "nieuwe motor {c}", "elektrische motorfiets {c}",
              "motorfiets registraties {c}"],
        "topics": T(
            ["lancering", "introductie", "onthult", "presenteert", "nieuw model", "nieuwe motor"],
            ["productie", "fabriek", "assemblage", "capaciteit", "fabricage"],
            ["dealer", "dealers", "showroom", "vestiging", "opent", "opening", "dealernetwerk"],
            ["prijs", "prijzen", "prijsverhoging", "korting", "kosten", "aanbieding", "financiering"],
            ["elektrisch", "elektrische", "accu", "batterij", "actieradius", "laden", "hybride", "technologie"],
            ["samenwerking", "overname", "investering", "investeert", "joint venture", "partnerschap", "fusie"],
            ["export", "exporteert", "import", "invoerrechten", "tarieven", "internationale uitbreiding"]),
    },
    "ru": {
        "wheel": ["мотоцикл", "мотоциклы", "скутер", "мопед", "мотороллер", "мототехника", "электромотоцикл"],
        "q": ["мотоциклы {c}", "продажи мотоциклов {c}", "новые мотоциклы {c}", "электромотоциклы {c}",
              "производство мотоциклов {c}", "цены на мотоциклы {c}"],
        "topics": T(
            ["презентовал", "представил", "новинка", "новая модель", "дебют", "запуск", "выходит на рынок", "анонс"],
            ["производство", "завод", "мощности", "сборка", "локализация"],
            ["дилер", "дилерский", "дилерская сеть", "салон", "открыл", "открытие"],
            ["цена", "цены", "подорожание", "снижение цен", "скидка", "стоимость", "кредит", "прайс"],
            ["электрический", "электро", "электромотоцикл", "батарея", "аккумулятор", "запас хода", "гибрид",
             "зарядка"],
            ["партнерство", "соглашение", "инвестиции", "инвестирует", "совместное предприятие", "покупка",
             "слияние"],
            ["экспорт", "экспортирует", "импорт", "пошлины", "тарифы", "внешние рынки", "международный"]),
    },
    "pl": {
        "wheel": ["motocykl", "motocykle", "skuter", "skutery", "motorower", "motorowery", "motorynka"],
        "q": ["motocykle {c}", "sprzedaż motocykli {c}", "nowe motocykle {c}", "motocykle elektryczne {c}",
              "rejestracje motocykli {c}"],
        "topics": T(
            ["premiera", "debiut", "nowy model", "nowy motocykl", "prezentuje", "wchodzi na rynek", "nowosc"],
            ["produkcja", "fabryka", "zaklad", "moce produkcyjne", "montaz", "linia produkcyjna"],
            ["dealer", "salon", "siec dealerska", "otwiera", "otwarcie", "dystrybutor", "nowy salon"],
            ["cena", "ceny", "podwyzka", "obnizka", "rabat", "promocja", "koszt", "cennik", "finansowanie"],
            ["elektryczny", "elektryczne", "akumulator", "bateria", "zasieg", "ladowanie", "hybryda", "technologia"],
            ["partnerstwo", "umowa", "inwestycja", "inwestuje", "joint venture", "przejecie", "fuzja",
             "wspolpraca"],
            ["eksport", "eksportuje", "import", "cla", "rynki zagraniczne", "ekspansja zagraniczna",
             "miedzynarodowy"]),
    },
    "id": {
        "wheel": ["motor", "sepeda motor", "skuter", "motor listrik", "motor matic"],
        "q": ["sepeda motor {c}", "penjualan sepeda motor {c}", "motor baru {c}", "motor listrik {c}",
              "harga motor {c}", "pabrik motor {c}"],
        "topics": T(
            ["peluncuran", "meluncurkan", "luncurkan", "resmi meluncur", "model baru", "motor baru", "hadir",
             "debut"],
            ["produksi", "pabrik", "kapasitas", "perakitan", "manufaktur", "lini produksi"],
            ["dealer", "diler", "jaringan dealer", "showroom", "buka cabang", "pembukaan", "diresmikan",
             "distributor"],
            ["harga", "kenaikan harga", "turun harga", "diskon", "promo", "cicilan", "kredit", "daftar harga"],
            ["listrik", "baterai", "jarak tempuh", "pengisian", "hybrid", "teknologi", "konektivitas", "abs",
             "swap baterai"],
            ["kemitraan", "kerja sama", "investasi", "berinvestasi", "joint venture", "akuisisi", "kolaborasi",
             "merger"],
            ["ekspor", "mengekspor", "impor", "tarif", "pasar ekspor", "ekspansi", "internasional"]),
    },
    "vi": {
        "wheel": ["xe máy", "mô tô", "xe tay ga", "xe máy điện", "xe côn tay", "xe số"],
        "q": ["xe máy {c}", "doanh số xe máy {c}", "xe máy mới ra mắt {c}", "xe máy điện {c}", "giá xe máy {c}",
              "nhà máy xe máy {c}"],
        "topics": T(
            ["ra mắt", "giới thiệu", "mẫu mới", "xe mới", "trình làng", "công bố"],
            ["sản xuất", "nhà máy", "công suất", "lắp ráp", "dây chuyền"],
            ["đại lý", "showroom", "cửa hàng", "khai trương", "mở rộng mạng lưới", "nhà phân phối"],
            ["giá", "giá bán", "tăng giá", "giảm giá", "khuyến mãi", "ưu đãi", "trả góp", "bảng giá"],
            ["điện", "xe máy điện", "pin", "quãng đường", "sạc", "hybrid", "công nghệ", "đổi pin"],
            ["hợp tác", "đầu tư", "liên doanh", "mua lại", "thỏa thuận", "đối tác", "sáp nhập"],
            ["xuất khẩu", "nhập khẩu", "thuế", "thị trường quốc tế", "mở rộng ra nước ngoài"]),
    },
    "th": {
        "wheel": ["มอเตอร์ไซค์", "รถจักรยานยนต์", "สกู๊ตเตอร์", "จักรยานยนต์ไฟฟ้า", "บิ๊กไบค์"],
        "q": ["มอเตอร์ไซค์ {c}", "ยอดขายมอเตอร์ไซค์ {c}", "มอเตอร์ไซค์ใหม่ {c}", "มอเตอร์ไซค์ไฟฟ้า {c}",
              "ราคามอเตอร์ไซค์ {c}"],
        "topics": T(
            ["เปิดตัว", "รุ่นใหม่", "ตัวใหม่"],
            ["ผลิต", "โรงงาน", "กำลังการผลิต", "ประกอบ"],
            ["ตัวแทนจำหน่าย", "ดีลเลอร์", "โชว์รูม", "เปิดสาขา", "ศูนย์บริการ"],
            ["ราคา", "ลดราคา", "ขึ้นราคา", "โปรโมชั่น", "ผ่อน", "ดาวน์"],
            ["ไฟฟ้า", "แบตเตอรี่", "ระยะทาง", "ชาร์จ", "ไฮบริด", "เทคโนโลยี"],
            ["ร่วมมือ", "ลงทุน", "พันธมิตร", "ข้อตกลง", "เข้าซื้อกิจการ", "ร่วมทุน"],
            ["ส่งออก", "นำเข้า", "ภาษี", "ตลาดต่างประเทศ"]),
    },
    "ar": {
        "wheel": ["دراجة نارية", "دراجات نارية", "سكوتر", "موتوسيكل", "الدراجات النارية", "الدراجة النارية",
                  "دراجات بخارية"],
        "q": ["دراجات نارية {c}", "مبيعات الدراجات النارية {c}", "دراجات نارية جديدة {c}",
              "دراجات نارية كهربائية {c}", "أسعار الدراجات النارية {c}"],
        "topics": T(
            ["إطلاق", "تطلق", "تكشف", "طراز جديد", "دراجة جديدة", "تدشين"],
            ["إنتاج", "مصنع", "تصنيع", "طاقة إنتاجية", "تجميع"],
            ["وكيل", "وكلاء", "معرض", "افتتاح", "فرع جديد"],
            ["سعر", "أسعار", "خصم", "ارتفاع الأسعار", "تخفيض", "تمويل", "تقسيط"],
            ["كهربائية", "كهربائي", "بطارية", "مدى", "شحن", "هجينة", "تقنية"],
            ["شراكة", "اتفاقية", "استثمار", "استحواذ", "تعاون", "مشروع مشترك", "اندماج"],
            ["تصدير", "استيراد", "رسوم جمركية", "أسواق خارجية", "توسع دولي"]),
    },
    # ---- Smaller languages: search + two-wheeler filter only. Topics fall back to the English keywords
    # (the AI classification step of the pipeline can refine them after translation).
    "uk": {"wheel": ["мотоцикл", "мотоцикли", "скутер", "мопед", "електромотоцикл"],
           "q": ["мотоцикли {c}", "електромотоцикли {c}", "ціни на мотоцикли {c}", "виробництво мотоциклів {c}",
                 "продаж мотоциклів {c}"]},
    "cs": {"wheel": ["motocykl", "motocykly", "motorka", "motorky", "skútr", "skútry"],
           "q": ["motocykly {c}", "prodej motocyklů {c}", "nové motocykly {c}", "elektromotocykly {c}",
                 "registrace motocyklů {c}"]},
    "sk": {"wheel": ["motocykel", "motocykle", "motorka", "skúter"],
           "q": ["motocykle {c}", "predaj motocyklov {c}", "nové motocykle {c}", "elektrické motocykle {c}"]},
    "ro": {"wheel": ["motocicletă", "motociclete", "scuter", "scutere", "moped", "mopede"],
           "q": ["motociclete {c}", "vânzări motociclete {c}", "motociclete noi {c}", "motociclete electrice {c}",
                 "înmatriculări motociclete {c}"]},
    "hu": {"wheel": ["motorkerékpár", "motorkerékpárok", "robogó", "robogók", "motorbicikli"],
           "q": ["motorkerékpár {c}", "motorkerékpár eladások {c}", "új motorkerékpár {c}",
                 "elektromos motorkerékpár {c}", "robogó {c}"]},
    "el": {"wheel": ["μοτοσυκλέτα", "μοτοσυκλέτες", "σκούτερ", "μηχανάκι", "μηχανάκια"],
           "q": ["μοτοσυκλέτες {c}", "πωλήσεις μοτοσυκλετών {c}", "νέες μοτοσυκλέτες {c}",
                 "ηλεκτρικές μοτοσυκλέτες {c}", "ταξινομήσεις μοτοσυκλετών {c}"]},
    "bg": {"wheel": ["мотоциклет", "мотоциклети", "скутер", "мотопед"],
           "q": ["мотоциклети {c}", "продажби на мотоциклети {c}", "нови мотоциклети {c}",
                 "електрически мотоциклети {c}", "регистрации мотоциклети {c}"]},
    "hr": {"wheel": ["motocikl", "motocikli", "skuter", "skuteri", "moped"],
           "q": ["motocikli {c}", "prodaja motocikala {c}", "novi motocikli {c}", "električni motocikli {c}",
                 "registracije motocikala {c}"]},
    "sl": {"wheel": ["motocikel", "motocikli", "motorno kolo", "motorna kolesa", "skuter", "skuterji"],
           "q": ["motorna kolesa {c}", "prodaja motorjev {c}", "električni motorji {c}", "registracije motorjev {c}"]},
    "sv": {"wheel": ["motorcykel", "motorcyklar", "skoter", "moped", "mopeder"],
           "q": ["motorcyklar {c}", "försäljning motorcyklar {c}", "ny motorcykel {c}", "elmotorcykel {c}",
                 "registreringar motorcyklar {c}"]},
    "da": {"wheel": ["motorcykel", "motorcykler", "scooter", "knallert"],
           "q": ["motorcykler {c}", "salg af motorcykler {c}", "ny motorcykel {c}", "elmotorcykel {c}",
                 "indregistreringer motorcykler {c}"]},
    "no": {"wheel": ["motorsykkel", "motorsykler", "scooter", "moped"],
           "q": ["motorsykler {c}", "salg av motorsykler {c}", "ny motorsykkel {c}", "elmotorsykkel {c}",
                 "registreringer motorsykler {c}"]},
    "fi": {"wheel": ["moottoripyörä", "moottoripyörät", "skootteri", "mopo"],
           "q": ["moottoripyörät {c}", "moottoripyörien myynti {c}", "uusi moottoripyörä {c}",
                 "sähkömoottoripyörä {c}", "moottoripyörien rekisteröinnit {c}"]},
    "et": {"wheel": ["mootorratas", "mootorrattad", "mopeed", "skuuter"],
           "q": ["mootorrattad {c}", "mootorrataste müük {c}", "uus mootorratas {c}", "elektrimootorratas {c}"]},
    "lv": {"wheel": ["motocikls", "motocikli", "mopēds", "skuteris"],
           "q": ["motocikli {c}", "motociklu pārdošana {c}", "jauns motocikls {c}", "elektriskie motocikli {c}"]},
}

# Extra local-slang searches
COUNTRY_EXTRA_QUERIES = {
    "Kenya": ["boda boda {c}"], "Uganda": ["boda boda {c}"], "Tanzania": ["boda boda {c}"],
    "Rwanda": ["moto taxi {c}"], "Nigeria": ["okada {c}"], "Ghana": ["okada {c}"],
}

# =====================================================================
# SETTINGS
# =====================================================================
GOOGLE_NEWS = "https://news.google.com/rss/search?q={query}&hl={hl}&gl={gl}&ceid={ceid}"
DAYS_BACK_DEFAULT = 180      # ignore articles older than this
PAUSE_SECONDS = 1.5          # wait between requests (per worker) so Google does not block us
SIMILAR_TITLE = 0.85         # titles this similar (0-1) are treated as the same story
CACHE_DIR = "news_cache"

STOP = threading.Event()     # set by Ctrl+C so every worker stops cleanly
TRANSLATE_PAUSE = 0.2        # pause between translation calls
_GT = None                   # deep_translator.GoogleTranslator class (set when --translate is used)
EN_COL = "Headline (English)"

# Look-alike letters (e.g. a Cyrillic M inside "KYMCO") are replaced with the normal Latin letter
LOOKALIKES = str.maketrans({"\u041c": "M", "\u0410": "A", "\u0415": "E", "\u041e": "O", "\u0420": "P",
                            "\u0421": "C", "\u0422": "T", "\u041d": "H", "\u041a": "K", "\u0425": "X",
                            "\u0412": "B"})
# Thai and Arabic are written without spaces between words / with attached prefixes -> match anywhere
SUBSTRING_SCRIPTS = re.compile("[\u0e00-\u0e7f\u0600-\u06ff]")
ILLEGAL_XLSX = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


# =====================================================================
# BRAND TABLE
# =====================================================================
def display_name(raw):
    """HONDA -> Honda, TVS -> TVS. Names that are not all capitals are kept as written."""
    if raw != raw.upper():
        return raw
    return raw if raw in KEEP_UPPER else raw.title()


def load_brand_csv(path=BRANDS_FILE):
    """Optional file country_brands.csv -> {country: {brand: [aliases]}}"""
    result = {}
    if not os.path.exists(path):
        return result
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
            aliases = aliases + [a for a in NATIVE_ALIASES.get(name, []) if a not in aliases]
            result.setdefault(country, {})[name] = aliases
    return result


def build_countries():
    csv_brands = load_brand_csv()
    countries = {}
    for country, base in COUNTRY_SETTINGS.items():
        cfg = dict(base)
        if country in csv_brands:
            cfg["brands"] = csv_brands[country]
            cfg["brand_source"] = "country_brands.csv"
        else:
            names = list(dict.fromkeys(COUNTRY_EXTRA_BRANDS.get(country, []) + BRAND_SETS[cfg["region"]]))
            cfg["brands"] = {b: BRAND_ALIASES.get(b, [b]) + NATIVE_ALIASES.get(b, []) for b in names}
            cfg["brand_source"] = "built-in " + cfg["region"] + " list"
        countries[country] = cfg
    return countries


COUNTRIES = build_countries()


def country_name(country, lang):
    names = COUNTRIES[country]["names"]
    return names.get(lang) or names.get("en") or country


def edition(country, lang):
    """Google News edition parameters (hl, gl) for a country + language."""
    cfg = COUNTRIES[country]
    gl = cfg["gl"]
    if lang == "es":
        hl = "es-419" if cfg["region"] == "LATAM" else "es"
    elif lang == "pt":
        hl = "pt-BR" if gl == "BR" else "pt-PT"
    else:
        hl = lang
    return hl, gl


def wheel_words(country):
    """Two-wheeler words that count as 'about motorcycles' for this country (its languages + English)."""
    words = list(LANGS["en"]["wheel"])
    for lang in COUNTRIES[country]["langs"]:
        words += LANGS[lang]["wheel"]
    return list(dict.fromkeys(words))


# =====================================================================
# TEXT HELPERS
# =====================================================================
def clean_text(text):
    """Lowercase, remove accents and punctuation (works for every alphabet), so matching is simple."""
    text = unicodedata.normalize("NFKD", text or "")
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"<[^>]+>", " ", text)                  # remove html tags
    text = re.sub(r"[\W_]+", " ", text.lower())
    return " " + re.sub(r"\s+", " ", text).strip() + " "  # spaces at both ends help whole-word matching


def has_word(cleaned_text, phrase):
    """True if the phrase appears as a whole word/phrase (so 'abs' does not match 'absoluto')."""
    key = clean_text(phrase).strip()
    if SUBSTRING_SCRIPTS.search(key):
        return key in cleaned_text
    return (" " + key + " ") in cleaned_text


def has_keyword(cleaned_text, keyword):
    """Short keywords (4 letters or fewer) must match a whole word. Longer ones match the start of a
    word, so 'fabrica' also finds 'fabricas'. Thai/Arabic keywords match anywhere in the text."""
    key = clean_text(keyword).strip()
    if SUBSTRING_SCRIPTS.search(key):
        return key in cleaned_text
    if len(key) <= 4:
        return (" " + key + " ") in cleaned_text
    return (" " + key) in cleaned_text


def sanitize(value):
    return ILLEGAL_XLSX.sub("", value) if isinstance(value, str) else value


# =====================================================================
# GOOGLE NEWS
# =====================================================================
def fetch_feed(query, gl, hl, retries=3):
    """Download one Google News RSS search -> raw XML text, or None if it failed.
    429 / 5xx / network problems are retried with growing pauses. If the country edition is refused
    (HTTP 400/404) the US edition of the same language is tried instead (the country is in the query)."""
    base_lang = hl.split("-")[0]
    urls = [GOOGLE_NEWS.format(query=quote_plus(query), hl=hl, gl=gl, ceid=f"{gl}:{hl}")]
    if gl != "US":
        urls.append(GOOGLE_NEWS.format(query=quote_plus(query), hl=base_lang, gl="US", ceid=f"US:{base_lang}"))
    for url in urls:
        for attempt in range(retries + 1):
            if STOP.is_set():
                return None
            try:
                response = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
                if response.status_code == 200:
                    return response.text
                if response.status_code in (400, 404):
                    print(f"   ! HTTP {response.status_code} for edition {hl}/{gl} - trying fallback")
                    break
                print(f"   ! HTTP {response.status_code} for '{query}' (try {attempt + 1})")
            except requests.RequestException as error:
                print(f"   ! Network problem for '{query}': {error}")
            time.sleep(PAUSE_SECONDS * 2 ** (attempt + 1))
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
        if source and title.endswith(" - " + source):      # Google adds " - Source name" to titles
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


# =====================================================================
# JUDGING AND TAGGING
# =====================================================================
def find_brands(cleaned_text, brands):
    return [b for b, aliases in brands.items() if any(has_word(cleaned_text, a) for a in aliases)]


def classify_topic(cleaned_text, lang):
    """Return (main topic, all matching topics). Keywords of the article's language + English."""
    packs = [LANGS["en"]["topics"]]
    if lang != "en" and LANGS.get(lang, {}).get("topics"):
        packs.append(LANGS[lang]["topics"])
    scores = {}
    for topic in TOPIC_NAMES:
        keywords = [kw for pack in packs for kw in pack.get(topic, [])]
        hits = sum(1 for kw in set(keywords) if has_keyword(cleaned_text, kw))
        if hits:
            scores[topic] = hits
    if not scores:
        return OTHER_TOPIC, ""
    best = max(scores.values())
    main = next(t for t in TOPIC_NAMES if scores.get(t) == best)       # list order breaks ties
    others = [f"{t} ({n})" for t, n in sorted(scores.items(), key=lambda x: -x[1])]
    return main, ", ".join(others)


def judge_article(article, brands, wheel):
    """Decide whether to keep an article. Returns (keep, reason, brands, cleaned_text)."""
    cleaned = clean_text(article["title"] + " " + article["snippet"])
    found = find_brands(cleaned, brands)
    has_two_wheeler = any(has_keyword(cleaned, w) for w in wheel)

    if not found:
        if has_two_wheeler:
            return True, "", ["General / Market"], cleaned
        return False, "No brand and no two-wheeler word", [], cleaned
    if all(b.upper() in AMBIGUOUS_BRANDS for b in found) and not has_two_wheeler:
        return False, "Ambiguous brand without a two-wheeler word", found, cleaned
    return True, "", found, cleaned


def is_duplicate(cleaned_title, seen_titles):
    return any(difflib.SequenceMatcher(None, cleaned_title, old).ratio() >= SIMILAR_TITLE for old in seen_titles)


# =====================================================================
# SEARCHES
# =====================================================================
def or_group(words):
    return "(" + " OR ".join(f'"{w}"' if " " in w else w for w in words) + ")"


def build_queries(country, general_n=6, max_brands=0):
    """List of (search text, label, language). One search per brand + general market searches per language."""
    cfg = COUNTRIES[country]
    langs = cfg["langs"]
    main = langs[0]
    # words for brand searches: the first two words of each of the country's languages
    words = []
    for lang in langs:
        words += LANGS[lang]["wheel"][:2]
    words = list(dict.fromkeys(words))[:4]
    brand_items = list(cfg["brands"].items())
    if max_brands:
        brand_items = brand_items[:max_brands]

    queries = []
    for brand, aliases in brand_items:
        queries.append((f'"{aliases[0]}" {or_group(words)} {country_name(country, main)}', brand, main))
    for lang in langs:
        for template in LANGS[lang]["q"][:general_n]:
            queries.append((template.format(c=country_name(country, lang)), "General / Market", lang))
    for template in COUNTRY_EXTRA_QUERIES.get(country, []):
        queries.append((template.format(c=country_name(country, "en")), "General / Market", "en"))
    return queries


def collect_country(country, days_back, general_n, max_brands, from_file=None, translate=False):
    """Fetch + judge everything for ONE country. Returns dict(kept, rejected, stats) or None if interrupted."""
    cfg = COUNTRIES[country]
    cutoff = datetime.now(timezone.utc) - timedelta(days=days_back)
    wheel = wheel_words(country)
    kept, rejected, cs = [], [], Counter()
    seen_links, seen_titles = set(), []
    pending = []                       # (row, cleaned text) - topics are assigned after translation

    if from_file:                      # test mode: one saved XML file instead of calling Google
        with open(from_file, encoding="utf-8") as f:
            jobs = [("(file)", "(file)", cfg["langs"][0], f.read())]
    else:
        jobs = []
        for query, label, lang in build_queries(country, general_n, max_brands):
            if STOP.is_set():
                return None
            hl, gl = edition(country, lang)
            print(f"[{country}] {label:<18} | {query}")
            cs["searches"] += 1
            jobs.append((query, label, lang, fetch_feed(query, gl, hl)))
            time.sleep(PAUSE_SECONDS)

    for query, label, lang, xml_text in jobs:
        if xml_text is None:
            cs["searches failed"] += 1
            continue
        for article in parse_feed(xml_text):
            cs["articles seen"] += 1
            base = {"Search": query, "Headline": article["title"], "Source": article["source"],
                    "Published": article["published"], "Link": article["link"],
                    "Country": country, "Region": cfg["region"], "Language": lang}
            if article["published"] and article["published"] < cutoff:
                cs["too old"] += 1
                continue
            if article["link"] in seen_links:
                cs["duplicate link"] += 1
                continue
            seen_links.add(article["link"])

            keep, reason, brands, cleaned = judge_article(article, cfg["brands"], wheel)
            if not keep:
                rejected.append({**base, "Reason rejected": reason})
                cs["rejected"] += 1
                continue
            cleaned_title = clean_text(article["title"])
            if is_duplicate(cleaned_title, seen_titles):
                rejected.append({**base, "Reason rejected": "Duplicate story (similar headline)"})
                cs["duplicate story"] += 1
                continue
            seen_titles.append(cleaned_title)

            row = {**base, "Brand": ", ".join(brands), "Topic": OTHER_TOPIC, "All matching topics": ""}
            kept.append(row)
            pending.append((row, cleaned))
            cs["kept"] += 1

    if translate and pending:
        cs["translation failed"] = translate_rows([r for r, _ in pending])
    for row, cleaned in pending:       # topic tags use the original text + its English translation
        text = cleaned + clean_text(row.get(EN_COL, ""))
        row["Topic"], row["All matching topics"] = classify_topic(text, row["Language"])

    print(f"[{country}] done: seen {cs['articles seen']}, kept {cs['kept']}, failed searches {cs['searches failed']}")
    return {"kept": kept, "rejected": rejected, "stats": dict(cs)}


# =====================================================================
# TRANSLATION (optional, --translate)
# =====================================================================
_TR_LOCK = threading.Lock()
_TR_CACHE = {}


def _tr_file():
    return os.path.join(CACHE_DIR, "translations.json")


def load_translations():
    if os.path.exists(_tr_file()):
        with open(_tr_file(), encoding="utf-8") as f:
            _TR_CACHE.update(json.load(f))


def save_translations():
    os.makedirs(CACHE_DIR, exist_ok=True)
    with _TR_LOCK:
        tmp = _tr_file() + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(_TR_CACHE, f, ensure_ascii=False)
        os.replace(tmp, _tr_file())


def translate_rows(rows):
    """Fill the 'Headline (English)' column. English headlines are copied; the rest are translated
    (auto-detected source language). Returns the number of headlines that could not be translated."""
    translator = _GT(source="auto", target="en")
    failed = 0
    for row in rows:
        if row.get(EN_COL):
            continue
        text = row["Headline"]
        if row.get("Language") == "en" or not text.strip():
            row[EN_COL] = text
            continue
        with _TR_LOCK:
            hit = _TR_CACHE.get(text)
        if hit is not None:
            row[EN_COL] = hit
            continue
        row[EN_COL] = ""
        for attempt in range(3):
            if STOP.is_set():
                return failed
            try:
                out = translator.translate(text) or ""
                with _TR_LOCK:
                    _TR_CACHE[text] = out
                row[EN_COL] = out
                time.sleep(TRANSLATE_PAUSE)
                break
            except Exception:                           # noqa: BLE001 - rate limit / network: wait and retry
                time.sleep(TRANSLATE_PAUSE * 5 * (attempt + 1))
        else:
            failed += 1
    return failed


def translate_cached(result):
    """Add (or finish) translations on results loaded from the cache and refresh their topic tags."""
    failed = translate_rows(result["kept"])
    for row in result["kept"]:
        text = clean_text(row["Headline"]) + clean_text(row.get(EN_COL, ""))
        row["Topic"], row["All matching topics"] = classify_topic(text, row.get("Language", "en"))
    result["stats"]["translation failed"] = failed


# =====================================================================
# CACHE (resume)
# =====================================================================
def _to_json(row):
    row = dict(row)
    if row.get("Published"):
        row["Published"] = row["Published"].isoformat()
    return row


def _from_json(row):
    if row.get("Published"):
        row["Published"] = datetime.fromisoformat(row["Published"])
    return row


def cache_path(country):
    return os.path.join(CACHE_DIR, re.sub(r"[^A-Za-z0-9]+", "_", country) + ".json")


def save_cache(country, days, result, translated=False):
    os.makedirs(CACHE_DIR, exist_ok=True)
    payload = {"saved": datetime.now(timezone.utc).isoformat(), "days": days, "translated": translated,
               "kept": [_to_json(r) for r in result["kept"]],
               "rejected": [_to_json(r) for r in result["rejected"]], "stats": result["stats"]}
    with open(cache_path(country), "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)


def load_cache(country, days):
    path = cache_path(country)
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        payload = json.load(f)
    if payload.get("days") != days:
        return None
    return {"kept": [_from_json(r) for r in payload["kept"]],
            "rejected": [_from_json(r) for r in payload["rejected"]], "stats": payload["stats"]}


def run_country(country, args):
    """One country, never raises: returns (result or None, note)."""
    try:
        if args.resume and not args.from_file:
            cached = load_cache(country, args.days)
            if cached:
                print(f"[{country}] using saved results from {CACHE_DIR}/ (--resume)")
                note = "from cache"
                if args.translate and any(not r.get(EN_COL) for r in cached["kept"]):
                    translate_cached(cached)
                    save_cache(country, args.days, cached, True)
                    save_translations()
                    note = "from cache + translated"
                return cached, note
        result = collect_country(country, args.days, args.general, args.max_brands, args.from_file,
                                 translate=args.translate)
        if result is None:
            return None, "interrupted - not included"
        if not args.from_file:
            save_cache(country, args.days, result, args.translate)
            if args.translate:
                save_translations()
        return result, ""
    except Exception as error:                      # noqa: BLE001 - one bad country must not stop the run
        print(f"[{country}] FAILED: {error.__class__.__name__}: {error}")
        return None, f"FAILED: {error.__class__.__name__}: {error}"


# =====================================================================
# EXCEL
# =====================================================================
def write_header(ws, headers, widths):
    ws.append(headers)
    for col, width in enumerate(widths, start=1):
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
            values.append(sanitize(v))
        ws.append(values)
    headers = [c.value for c in ws[1]]
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            head = headers[cell.column - 1]
            if head == "Link" and cell.value and len(str(cell.value)) < 2000:
                cell.hyperlink = cell.value
                cell.font = Font(color="2563EB", underline="single")
            if head == "Published" and cell.value:
                cell.number_format = "dd-mmm-yyyy"
    ws.auto_filter.ref = ws.dimensions


def save_excel(kept, rejected, country_rows, path):
    wb = Workbook()

    ws = wb.active
    ws.title = "News"
    cols = ["Published", "Country", "Region", "Language", "Brand", "Topic", "Headline", "Source", "Link",
            "All matching topics", "Search"]
    widths = [13, 18, 14, 9, 18, 30, 70, 22, 40, 45, 40]
    if any(EN_COL in r for r in kept):                  # translated run: English headline next to the original
        cols.insert(7, EN_COL)
        widths.insert(7, 70)
    write_header(ws, cols, widths)
    add_rows(ws, kept, cols)

    ws2 = wb.create_sheet("Summary")
    topics = TOPIC_NAMES + [OTHER_TOPIC]
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

    ws3 = wb.create_sheet("Rejected")
    cols3 = ["Country", "Language", "Reason rejected", "Headline", "Source", "Published", "Link", "Search"]
    write_header(ws3, cols3, [18, 9, 35, 70, 22, 13, 40, 40])
    add_rows(ws3, rejected, cols3)

    # one line per country: spot countries with no results or failures at a glance
    ws4 = wb.create_sheet("Countries")
    cols4 = ["Country", "Region", "Languages", "Brands used from", "Searches", "Failed searches",
             "Articles seen", "Kept", "Translation failed", "Note"]
    write_header(ws4, cols4, [18, 14, 12, 28, 10, 14, 13, 8, 12, 50])
    for row in country_rows:
        ws4.append([sanitize(row.get(c, "")) for c in cols4])
    ws4.auto_filter.ref = ws4.dimensions

    ws5 = wb.create_sheet("Run info")
    write_header(ws5, ["Item", "Value"], [34, 30])
    ws5.append(["Run time (UTC)", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M")])
    totals = Counter()
    for row in country_rows:
        for key in ("Searches", "Failed searches", "Articles seen", "Kept"):
            totals[key] += row.get(key) or 0
    for key, value in totals.items():
        ws5.append([key, value])
    ws5.append(["Countries with problems", sum(1 for r in country_rows if str(r.get("Note", "")).startswith(("FAILED", "interrupted")))])

    wb.save(path)


# =====================================================================
# MAIN
# =====================================================================
def resolve_countries(text):
    if text.strip().lower() == "all":
        return list(COUNTRIES)
    lookup = {c.lower(): c for c in COUNTRIES}
    lookup.update(COUNTRY_ALIASES)
    selected = []
    for name in text.split(","):
        key = name.strip().lower()
        if not key:
            continue
        target = lookup.get(key)
        if target is None:
            raise SystemExit(f"Unknown country '{name.strip()}'. Available: {', '.join(COUNTRIES)}")
        if target not in selected:                 # every country only once
            selected.append(target)
    return selected


def probe(selected):
    """Quick check of the Google News edition + first general search of each country/language."""
    for country in selected:
        for lang in COUNTRIES[country]["langs"]:
            hl, gl = edition(country, lang)
            query = LANGS[lang]["q"][0].format(c=country_name(country, lang))
            xml_text = fetch_feed(query, gl, hl, retries=1)
            n = len(parse_feed(xml_text)) if xml_text else None
            print(f"{country:<20} {lang:<3} edition {hl}/{gl:<3} | '{query}' -> "
                  f"{'FAILED' if n is None else str(n) + ' articles'}")
            time.sleep(PAUSE_SECONDS)


def main():
    global _GT, TRANSLATE_PAUSE
    parser = argparse.ArgumentParser(description="Fetch and categorise two-wheeler news for many countries")
    parser.add_argument("--countries", default="all", help='comma-separated, e.g. "Thailand,Germany" (default: all)')
    parser.add_argument("--days", type=int, default=DAYS_BACK_DEFAULT, help="how many days back to keep")
    parser.add_argument("--out", default="global_2w_news.xlsx", help="Excel file to create")
    parser.add_argument("--workers", type=int, default=3, help="countries fetched in parallel (default 3)")
    parser.add_argument("--general", type=int, default=6, help="general market searches per language (default 6)")
    parser.add_argument("--max-brands", type=int, default=0, help="max brand searches per country (0 = all)")
    parser.add_argument("--resume", action="store_true", help=f"re-use countries already saved in {CACHE_DIR}/")
    parser.add_argument("--translate", action="store_true", help="add a 'Headline (English)' column (needs deep-translator)")
    parser.add_argument("--translate-pause", type=float, default=TRANSLATE_PAUSE, help="seconds between translations")
    parser.add_argument("--probe", action="store_true", help="only test the Google News edition of each country")
    parser.add_argument("--from-file", help="test mode: read a saved RSS xml file instead of calling Google")
    args = parser.parse_args()

    selected = resolve_countries(args.countries)
    if args.translate:
        try:
            from deep_translator import GoogleTranslator
        except ImportError:
            raise SystemExit("--translate needs the deep-translator package:  pip install deep-translator")
        _GT, TRANSLATE_PAUSE = GoogleTranslator, args.translate_pause
        load_translations()
    if args.probe:
        probe(selected)
        return

    searches = sum(len(build_queries(c, args.general, args.max_brands)) for c in selected)
    minutes = searches * (PAUSE_SECONDS + 1) / max(args.workers, 1) / 60
    print(f"{len(selected)} countries | about {searches} searches | {args.workers} workers | ~{minutes:.0f} min")

    results, notes = {}, {}
    executor = ThreadPoolExecutor(max_workers=max(args.workers, 1))
    futures = {executor.submit(run_country, c, args): c for c in selected}
    try:
        for fut in as_completed(futures):
            country = futures[fut]
            results[country], notes[country] = fut.result()
    except KeyboardInterrupt:
        print("\nInterrupted - finishing current searches, then saving what is complete...")
        STOP.set()
        executor.shutdown(wait=True, cancel_futures=True)
        for fut, country in futures.items():
            if country not in results and fut.done() and not fut.cancelled():
                results[country], notes[country] = fut.result()
    else:
        executor.shutdown()

    kept, rejected, country_rows = [], [], []
    for country in selected:
        result, note = results.get(country), notes.get(country, "not run (interrupted)")
        cfg = COUNTRIES[country]
        stats = (result or {}).get("stats", {})
        if result:
            kept += result["kept"]
            rejected += result["rejected"]
        country_rows.append({
            "Country": country, "Region": cfg["region"], "Languages": "/".join(cfg["langs"]),
            "Brands used from": cfg["brand_source"], "Searches": stats.get("searches", 0),
            "Failed searches": stats.get("searches failed", 0), "Articles seen": stats.get("articles seen", 0),
            "Kept": stats.get("kept", 0), "Translation failed": stats.get("translation failed", 0), "Note": note,
        })
    kept.sort(key=lambda r: r["Published"] or datetime.min.replace(tzinfo=timezone.utc), reverse=True)

    if args.translate:
        save_translations()
    save_excel(kept, rejected, country_rows, args.out)
    bad = [r["Country"] for r in country_rows if str(r["Note"]).startswith(("FAILED", "interrupted", "not run"))]
    print(f"\nDone. {len(selected) - len(bad)} countries OK, {len(bad)} not included"
          + (f": {', '.join(bad)}" if bad else ""))
    print(f"Saved {len(kept)} articles to {args.out}")


if __name__ == "__main__":
    main()
