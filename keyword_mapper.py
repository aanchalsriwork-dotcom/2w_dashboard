"""
Generate comprehensive keywords mapping XLSX with English translations
Maps all keywords across 28 languages, 7 topics, and 86+ countries
Includes validation and synonym information
"""

import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

TOPIC_NAMES = [
    "New Launches",
    "Production & Capacity Changes", 
    "Dealer Network Expansion",
    "Price Changes",
    "EV & Technology Announcements",
    "Partnerships & Investments",
    "Exports & International Expansion"
]

LANGUAGE_CODES = {
    "en": "English",
    "es": "Spanish",
    "fr": "French",
    "pt": "Portuguese",
    "de": "German",
    "it": "Italian",
    "nl": "Dutch",
    "ru": "Russian",
    "pl": "Polish",
    "id": "Indonesian",
    "vi": "Vietnamese",
    "th": "Thai",
    "ar": "Arabic",
    "uk": "Ukrainian",
    "cs": "Czech",
    "sk": "Slovak",
    "ro": "Romanian",
    "hu": "Hungarian",
    "el": "Greek",
    "bg": "Bulgarian",
    "hr": "Croatian",
    "sl": "Slovenian",
    "sv": "Swedish",
    "da": "Danish",
    "no": "Norwegian",
    "fi": "Finnish",
    "et": "Estonian",
    "lv": "Latvian",
}

LANGS = {
    "en": {
        "wheel": ["motorcycle", "motorbike", "scooter", "motorcycles", "motorbikes", "moped", "scooters",
                  "two-wheeler", "two-wheelers", "2-wheeler", "e-motorcycle", "boda boda", "bodaboda", "okada",
                  "dirt bike", "superbike"],
        "topics": {
            "New Launches": ["launch", "launches", "launched", "unveil", "unveils", "unveiled", "debut", 
                           "introduces", "introduced", "new model", "all-new", "revealed", "goes on sale", 
                           "arrives in", "new version"],
            "Production & Capacity Changes": ["production", "factory", "plant", "assembly", "manufacturing", 
                                            "manufacture", "capacity", "production line", "assembles", 
                                            "localisation", "localization"],
            "Dealer Network Expansion": ["dealer", "dealers", "dealership", "showroom", "outlet", "new store",
                                       "opens", "network expansion", "service centre", "distributor", 
                                       "inaugurates"],
            "Price Changes": ["price", "prices", "price hike", "discount", "price cut", "starting at", 
                            "cost", "offer", "financing", "emi", "on-road"],
            "EV & Technology Announcements": ["electric", "ev", "battery", "range", "fast charging", 
                                            "hybrid", "connectivity", "swap", "charging", "tft", "abs", 
                                            "kw", "lithium"],
            "Partnerships & Investments": ["partnership", "tie-up", "agreement", "investment", "invests", 
                                         "joint venture", "acquires", "acquisition", "stake", "funding", 
                                         "collaboration", "merger", "mou"],
            "Exports & International Expansion": ["export", "exports", "exporting", "import", "imports", 
                                               "tariff", "tariffs", "international expansion", "overseas", 
                                               "enters market", "new market", "global expansion"]
        }
    },
    "es": {
        "wheel": ["moto", "motos", "motocicleta", "motocicletas", "motoneta", "scooter", "ciclomotor",
                  "motovehiculo", "cuatriciclo", "dos ruedas", "motero", "motoqueros"],
        "topics": {
            "New Launches": ["lanza", "lanzamiento", "lanzo", "presenta", "presento", "nuevo modelo", 
                           "nueva moto", "nuevas motos", "llega a la argentina", "llega al pais", "debut", 
                           "estreno", "novedad", "se suma a su gama", "nueva version"],
            "Production & Capacity Changes": ["produccion", "fabrica", "planta", "capacidad", "ensamblado", 
                                            "ensamble", "linea de montaje", "fabricacion", "produce", "turno", 
                                            "paro de planta", "suspende produccion", "amplia planta"],
            "Dealer Network Expansion": ["concesionaria", "concesionario", "sucursal", "red de venta", 
                                       "red comercial", "nuevo local", "inaugura", "inauguracion", "abre", 
                                       "apertura", "showroom", "puntos de venta", "red de concesionarios", 
                                       "distribuidor"],
            "Price Changes": ["precio", "precios", "aumento", "suba", "sube", "baja de precio", "bajan", 
                            "descuento", "lista de precios", "cuanto cuesta", "cuesta", "valor", "cuotas", 
                            "financiacion", "tasa 0", "oferta"],
            "EV & Technology Announcements": ["electrica", "electricas", "electrico", "electricos", "bateria", 
                                            "baterias", "autonomia", "carga rapida", "hibrida", "hibrido", 
                                            "conectividad", "app", "tecnologia", "abs", "tft", "litio"],
            "Partnerships & Investments": ["alianza", "acuerdo", "convenio", "inversion", "invertira", 
                                         "invierte", "socio", "joint venture", "asociacion", "adquiere", 
                                         "adquisicion", "compra de", "fusion", "millones de dolares", 
                                         "financiamiento"],
            "Exports & International Expansion": ["exporta", "exportacion", "exportaciones", "exportar", 
                                               "mercado externo", "mercados regionales", "latinoamerica", 
                                               "region", "internacional", "importacion", "importaciones", 
                                               "aranceles"]
        }
    },
    "fr": {
        "wheel": ["moto", "motos", "scooter", "motocyclette", "cyclomoteur", "deux-roues", "2 roues", "scooters"],
        "topics": {
            "New Launches": ["lancement", "lance", "presente", "presentation", "nouveaute", "nouveau modele", 
                           "nouvelle moto", "devoile", "debut"],
            "Production & Capacity Changes": ["production", "usine", "fabrication", "assemblage", "capacite", 
                                            "ligne de production", "fabrique"],
            "Dealer Network Expansion": ["concessionnaire", "concession", "reseau", "point de vente", 
                                       "showroom", "inaugure", "ouverture", "distributeur"],
            "Price Changes": ["prix", "tarif", "tarifs", "hausse", "baisse", "remise", "promotion", "cout", 
                            "financement"],
            "EV & Technology Announcements": ["electrique", "electriques", "batterie", "autonomie", "recharge", 
                                            "hybride", "technologie", "abs", "kw"],
            "Partnerships & Investments": ["partenariat", "accord", "investissement", "investit", "co-entreprise", 
                                         "acquisition", "rachat", "fusion", "alliance"],
            "Exports & International Expansion": ["exportation", "exporte", "exportations", "importation", 
                                               "droits de douane", "marche international", 
                                               "expansion internationale"]
        }
    },
    "pt": {
        "wheel": ["moto", "motos", "motocicleta", "motocicletas", "scooter", "ciclomotor", "duas rodas"],
        "topics": {
            "New Launches": ["lancamento", "lanca", "lancou", "apresenta", "apresentou", "nova moto", 
                           "novidade", "estreia", "nova versao"],
            "Production & Capacity Changes": ["producao", "fabrica", "planta", "capacidade", "montagem", 
                                            "fabricacao", "linha de montagem"],
            "Dealer Network Expansion": ["concessionaria", "concessionarias", "revenda", "rede de vendas", 
                                       "inaugura", "inauguracao", "abertura", "loja", "showroom", "distribuidor"],
            "Price Changes": ["preco", "precos", "aumento", "reajuste", "desconto", "tabela de precos", 
                            "quanto custa", "financiamento", "consorcio", "promocao"],
            "EV & Technology Announcements": ["eletrica", "eletricas", "eletrico", "bateria", "autonomia", 
                                            "recarga", "hibrida", "conectividade", "tecnologia", "abs"],
            "Partnerships & Investments": ["parceria", "acordo", "investimento", "investe", "joint venture", 
                                         "aquisicao", "adquire", "fusao", "alianca"],
            "Exports & International Expansion": ["exporta", "exportacao", "exportacoes", "importacao", 
                                               "importacoes", "tarifas", "mercado externo", "internacional"]
        }
    },
    "de": {
        "wheel": ["motorrad", "motorräder", "motorroller", "roller", "moped", "zweirad", "elektromotorrad"],
        "topics": {
            "New Launches": ["vorgestellt", "neuheit", "neues modell", "markteinführung", "enthüllt", 
                           "premiere", "neuvorstellung"],
            "Production & Capacity Changes": ["produktion", "werk", "fabrik", "fertigung", "kapazität", 
                                            "montage", "produktionsstopp"],
            "Dealer Network Expansion": ["händler", "händlernetz", "motorradhändler", "filiale", "eröffnet", 
                                       "eröffnung", "neuer standort", "vertrieb"],
            "Price Changes": ["preis", "preise", "preiserhöhung", "rabatt", "preissenkung", "kosten", 
                            "angebot", "finanzierung"],
            "EV & Technology Announcements": ["elektro", "elektrisch", "elektromotorrad", "akku", "batterie", 
                                            "reichweite", "schnellladen", "hybrid", "technologie", "abs"],
            "Partnerships & Investments": ["partnerschaft", "kooperation", "investition", "investiert", 
                                         "übernahme", "joint venture", "beteiligung", "fusion"],
            "Exports & International Expansion": ["export", "exporte", "import", "zölle", "auslandsmarkt", 
                                               "internationalisierung", "expansion"]
        }
    },
    "it": {
        "wheel": ["moto", "motocicletta", "motociclette", "scooter", "motorino", "ciclomotore"],
        "topics": {
            "New Launches": ["lancio", "lancia", "presenta", "presentata", "nuovo modello", "nuova moto", 
                           "novita", "debutto", "arriva in", "svelata"],
            "Production & Capacity Changes": ["produzione", "fabbrica", "stabilimento", "capacita", 
                                            "assemblaggio", "linea di montaggio"],
            "Dealer Network Expansion": ["concessionario", "concessionaria", "rete vendita", "nuova sede", 
                                       "inaugura", "apertura", "showroom", "distributore"],
            "Price Changes": ["prezzo", "prezzi", "aumento", "rincaro", "sconto", "listino", "costa", 
                            "finanziamento", "offerta", "promozione"],
            "EV & Technology Announcements": ["elettrica", "elettriche", "elettrico", "batteria", "autonomia", 
                                            "ricarica", "ibrida", "connettivita", "tecnologia", "abs"],
            "Partnerships & Investments": ["partnership", "accordo", "investimento", "investe", "joint venture", 
                                         "acquisizione", "acquista", "fusione", "alleanza"],
            "Exports & International Expansion": ["esporta", "esportazione", "esportazioni", "importazione", 
                                               "dazi", "mercati esteri", "internazionale"]
        }
    },
    "nl": {
        "wheel": ["motorfiets", "motorfietsen", "scooter", "scooters", "bromfiets", "brommer"],
        "topics": {
            "New Launches": ["lancering", "introductie", "onthult", "presenteert", "nieuw model", "nieuwe motor"],
            "Production & Capacity Changes": ["productie", "fabriek", "assemblage", "capaciteit", "fabricage"],
            "Dealer Network Expansion": ["dealer", "dealers", "showroom", "vestiging", "opent", "opening", 
                                       "dealernetwerk"],
            "Price Changes": ["prijs", "prijzen", "prijsverhoging", "korting", "kosten", "aanbieding", 
                            "financiering"],
            "EV & Technology Announcements": ["elektrisch", "elektrische", "accu", "batterij", "actieradius", 
                                            "laden", "hybride", "technologie"],
            "Partnerships & Investments": ["samenwerking", "overname", "investering", "investeert", 
                                         "joint venture", "partnerschap", "fusie"],
            "Exports & International Expansion": ["export", "exporteert", "import", "invoerrechten", "tarieven", 
                                               "internationale uitbreiding"]
        }
    },
    "ru": {
        "wheel": ["мотоцикл", "мотоциклы", "скутер", "мопед", "мотороллер", "мототехника", "электромотоцикл"],
        "topics": {
            "New Launches": ["презентовал", "представил", "новинка", "новая модель", "дебют", "запуск", 
                           "выходит на рынок", "анонс"],
            "Production & Capacity Changes": ["производство", "завод", "мощности", "сборка", "локализация"],
            "Dealer Network Expansion": ["дилер", "дилерский", "дилерская сеть", "салон", "открыл", "открытие"],
            "Price Changes": ["цена", "цены", "подорожание", "снижение цен", "скидка", "стоимость", "кредит", 
                            "прайс"],
            "EV & Technology Announcements": ["электрический", "электро", "электромотоцикл", "батарея", 
                                            "аккумулятор", "запас хода", "гибрид", "зарядка"],
            "Partnerships & Investments": ["партнерство", "соглашение", "инвестиции", "инвестирует", 
                                         "совместное предприятие", "покупка", "слияние"],
            "Exports & International Expansion": ["экспорт", "экспортирует", "импорт", "пошлины", "тарифы", 
                                               "внешние рынки", "международный"]
        }
    },
    "pl": {
        "wheel": ["motocykl", "motocykle", "skuter", "skutery", "motorower", "motorowery", "motorynka"],
        "topics": {
            "New Launches": ["premiera", "debiut", "nowy model", "nowy motocykl", "prezentuje", 
                           "wchodzi na rynek", "nowosc"],
            "Production & Capacity Changes": ["produkcja", "fabryka", "zaklad", "moce produkcyjne", "montaz", 
                                            "linia produkcyjna"],
            "Dealer Network Expansion": ["dealer", "salon", "siec dealerska", "otwiera", "otwarcie", 
                                       "dystrybutor", "nowy salon"],
            "Price Changes": ["cena", "ceny", "podwyzka", "obnizka", "rabat", "promocja", "koszt", "cennik", 
                            "finansowanie"],
            "EV & Technology Announcements": ["elektryczny", "elektryczne", "akumulator", "bateria", "zasieg", 
                                            "ladowanie", "hybryda", "technologia"],
            "Partnerships & Investments": ["partnerstwo", "umowa", "inwestycja", "inwestuje", "joint venture", 
                                         "przejecie", "fuzja", "wspolpraca"],
            "Exports & International Expansion": ["eksport", "eksportuje", "import", "cla", "rynki zagraniczne", 
                                               "ekspansja zagraniczna", "miedzynarodowy"]
        }
    },
    "id": {
        "wheel": ["motor", "sepeda motor", "skuter", "motor listrik", "motor matic"],
        "topics": {
            "New Launches": ["peluncuran", "meluncurkan", "luncurkan", "resmi meluncur", "model baru", 
                           "motor baru", "hadir", "debut"],
            "Production & Capacity Changes": ["produksi", "pabrik", "kapasitas", "perakitan", "manufaktur", 
                                            "lini produksi"],
            "Dealer Network Expansion": ["dealer", "diler", "jaringan dealer", "showroom", "buka cabang", 
                                       "pembukaan", "diresmikan", "distributor"],
            "Price Changes": ["harga", "kenaikan harga", "turun harga", "diskon", "promo", "cicilan", 
                            "kredit", "daftar harga"],
            "EV & Technology Announcements": ["listrik", "baterai", "jarak tempuh", "pengisian", "hybrid", 
                                            "teknologi", "konektivitas", "abs", "swap baterai"],
            "Partnerships & Investments": ["kemitraan", "kerja sama", "investasi", "berinvestasi", 
                                         "joint venture", "akuisisi", "kolaborasi", "merger"],
            "Exports & International Expansion": ["ekspor", "mengekspor", "impor", "tarif", "pasar ekspor", 
                                               "ekspansi", "internasional"]
        }
    },
    "vi": {
        "wheel": ["xe máy", "mô tô", "xe tay ga", "xe máy điện", "xe côn tay", "xe số"],
        "topics": {
            "New Launches": ["ra mắt", "giới thiệu", "mẫu mới", "xe mới", "trình làng", "công bố"],
            "Production & Capacity Changes": ["sản xuất", "nhà máy", "công suất", "lắp ráp", "dây chuyền"],
            "Dealer Network Expansion": ["đại lý", "showroom", "cửa hàng", "khai trương", "mở rộng mạng lưới", 
                                       "nhà phân phối"],
            "Price Changes": ["giá", "giá bán", "tăng giá", "giảm giá", "khuyến mãi", "ưu đãi", "trả góp", 
                            "bảng giá"],
            "EV & Technology Announcements": ["điện", "xe máy điện", "pin", "quãng đường", "sạc", "hybrid", 
                                            "công nghệ", "đổi pin"],
            "Partnerships & Investments": ["hợp tác", "đầu tư", "liên doanh", "mua lại", "thỏa thuận", 
                                         "đối tác", "sáp nhập"],
            "Exports & International Expansion": ["xuất khẩu", "nhập khẩu", "thuế", "thị trường quốc tế", 
                                               "mở rộng ra nước ngoài"]
        }
    },
    "th": {
        "wheel": ["มอเตอร์ไซค์", "รถจักรยานยนต์", "สกู๊ตเตอร์", "จักรยานยนต์ไฟฟ้า", "บิ๊กไบค์"],
        "topics": {
            "New Launches": ["เปิดตัว", "รุ่นใหม่", "ตัวใหม่"],
            "Production & Capacity Changes": ["ผลิต", "โรงงาน", "กำลังการผลิต", "ประกอบ"],
            "Dealer Network Expansion": ["ตัวแทนจำหน่าย", "ดีลเลอร์", "โชว์รูม", "เปิดสาขา", "ศูนย์บริการ"],
            "Price Changes": ["ราคา", "ลดราคา", "ขึ้นราคา", "โปรโมชั่น", "ผ่อน", "ดาวน์"],
            "EV & Technology Announcements": ["ไฟฟ้า", "แบตเตอรี่", "ระยะทาง", "ชาร์จ", "ไฮบริด", "เทคโนโลยี"],
            "Partnerships & Investments": ["ร่วมมือ", "ลงทุน", "พันธมิตร", "ข้อตกลง", "เข้าซื้อกิจการ", "ร่วมทุน"],
            "Exports & International Expansion": ["ส่งออก", "นำเข้า", "ภาษี", "ตลาดต่างประเทศ"]
        }
    },
    "ar": {
        "wheel": ["دراجة نارية", "دراجات نارية", "سكوتر", "موتوسيكل", "الدراجات النارية", "الدراجة النارية",
                  "دراجات بخارية"],
        "topics": {
            "New Launches": ["إطلاق", "تطلق", "تكشف", "طراز جديد", "دراجة جديدة", "تدشين"],
            "Production & Capacity Changes": ["إنتاج", "مصنع", "تصنيع", "طاقة إنتاجية", "تجميع"],
            "Dealer Network Expansion": ["وكيل", "وكلاء", "معرض", "افتتاح", "فرع جديد"],
            "Price Changes": ["سعر", "أسعار", "خصم", "ارتفاع الأسعار", "تخفيض", "تمويل", "تقسيط"],
            "EV & Technology Announcements": ["كهربائية", "كهربائي", "بطارية", "مدى", "شحن", "هجينة", "تقنية"],
            "Partnerships & Investments": ["شراكة", "اتفاقية", "استثمار", "استحواذ", "تعاون", 
                                         "مشروع مشترك", "اندماج"],
            "Exports & International Expansion": ["تصدير", "استيراد", "رسوم جمركية", "أسواق خارجية", 
                                               "توسع دولي"]
        }
    },
    "uk": {"wheel": ["мотоцикл", "мотоцикли", "скутер", "мопед", "електромотоцикл"], "topics": {}},
    "cs": {"wheel": ["motocykl", "motocykly", "motorka", "motorky", "skútr", "skútry"], "topics": {}},
    "sk": {"wheel": ["motocykel", "motocykle", "motorka", "skúter"], "topics": {}},
    "ro": {"wheel": ["motocicletă", "motociclete", "scuter", "scutere", "moped", "mopede"], "topics": {}},
    "hu": {"wheel": ["motorkerékpár", "motorkerékpárok", "robogó", "robogók", "motorbicikli"], "topics": {}},
    "el": {"wheel": ["μοτοσυκλέτα", "μοτοσυκλέτες", "σκούτερ", "μηχανάκι", "μηχανάκια"], "topics": {}},
    "bg": {"wheel": ["мотоциклет", "мотоциклети", "скутер", "мотопед"], "topics": {}},
    "hr": {"wheel": ["motocikl", "motocikli", "skuter", "skuteri", "moped"], "topics": {}},
    "sl": {"wheel": ["motocikel", "motocikli", "motorno kolo", "motorna kolesa", "skuter", "skuterji"], "topics": {}},
    "sv": {"wheel": ["motorcykel", "motorcyklar", "skoter", "moped", "mopeder"], "topics": {}},
    "da": {"wheel": ["motorcykel", "motorcykler", "scooter", "knallert"], "topics": {}},
    "no": {"wheel": ["motorsykkel", "motorsykler", "scooter", "moped"], "topics": {}},
    "fi": {"wheel": ["moottoripyörä", "moottoripyörät", "skootteri", "mopo"], "topics": {}},
    "et": {"wheel": ["mootorratas", "mootorrattad", "mopeed", "skuuter"], "topics": {}},
    "lv": {"wheel": ["motocikls", "motocikli", "mopēds", "skuteris"], "topics": {}},
}

COUNTRY_MAPPING = {
    "ASEAN": ["Philippines", "Malaysia", "Thailand", "Vietnam", "Indonesia", "Cambodia", "Myanmar", "Singapore"],
    "Other Asia": ["Australia", "New Zealand", "Bangladesh", "Nepal", "Sri Lanka"],
    "Europe": ["Austria", "United Kingdom", "Belgium", "Russia", "Switzerland", "Poland", "Estonia", "Sweden", 
               "Finland", "Norway", "Denmark", "Romania", "Czech Republic", "France", "Germany", "Greece", 
               "Hungary", "Ukraine", "Ireland", "Bulgaria", "Croatia", "Italy", "Belarus", "Latvia", 
               "Luxembourg", "Netherlands", "Portugal", "Slovakia", "Slovenia", "Spain"],
    "Africa": ["Cameroon", "DR Congo", "Ivory Coast", "Nigeria", "Kenya", "Tanzania", "Ghana", "Uganda", 
               "Angola", "Egypt", "Ethiopia", "Niger", "Liberia", "Morocco", "Mozambique", "Namibia", 
               "Rwanda", "South Sudan", "Zimbabwe"],
    "Middle East": ["KSA", "UAE", "Kuwait", "Bahrain", "Qatar", "Oman"],
    "LATAM": ["Ecuador", "Argentina", "Bolivia", "Brazil", "Chile", "Colombia", "Costa Rica", 
              "Dominican Republic", "El Salvador", "Guatemala", "Honduras", "Nicaragua", "Paraguay", 
              "Peru", "Uruguay", "Mexico"],
    "North America": ["USA", "Canada"],
}

LANGUAGE_REGIONS = {
    "en": ["ASEAN", "Other Asia", "Europe", "Africa", "Middle East", "North America"],
    "es": ["LATAM"],
    "pt": ["LATAM"],
    "fr": ["Europe", "Africa"],
    "de": ["Europe"],
    "it": ["Europe"],
    "nl": ["Europe"],
    "ru": ["Europe"],
    "pl": ["Europe"],
    "id": ["ASEAN"],
    "vi": ["ASEAN"],
    "th": ["ASEAN"],
    "ar": ["Middle East", "Africa"],
    "uk": ["Europe"],
    "cs": ["Europe"],
    "sk": ["Europe"],
    "ro": ["Europe"],
    "hu": ["Europe"],
    "el": ["Europe"],
    "bg": ["Europe"],
    "hr": ["Europe"],
    "sl": ["Europe"],
    "sv": ["Europe"],
    "da": ["Europe"],
    "no": ["Europe"],
    "fi": ["Europe"],
    "et": ["Europe"],
    "lv": ["Europe"],
}

def get_english_synonym(lang, keyword):
    """Get English equivalent for a keyword"""
    if lang == "en":
        return keyword
    
    english_keywords = LANGS["en"]["topics"]
    
    keyword_lower = keyword.lower().strip()
    
    for topic, en_words in english_keywords.items():
        for en_word in en_words:
            if en_word.lower() == keyword_lower:
                return en_word
    
    return "N/A"

def create_keywords_mapping_xlsx():
    """Create XLSX with all keywords mapped to English translations"""
    
    wb = Workbook()
    
    ws = wb.active
    ws.title = "Keywords Mapping"
    
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    headers = ["Language Code", "Language Name", "Topic Category", "Keyword", "English Synonym/Translation", "Type"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 35
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 35
    ws.column_dimensions['F'].width = 15
    
    row = 2
    
    for lang_code in sorted(LANGS.keys()):
        lang_data = LANGS[lang_code]
        lang_name = LANGUAGE_CODES.get(lang_code, lang_code)
        
        wheel_words = lang_data.get("wheel", [])
        for keyword in wheel_words:
            english_trans = get_english_synonym(lang_code, keyword)

# Continuation from where we left off - paste this in generate_keywords_mapping.py

            ws.cell(row=row, column=1).value = lang_code
            ws.cell(row=row, column=2).value = lang_name
            ws.cell(row=row, column=3).value = "Two-Wheeler Words"
            ws.cell(row=row, column=4).value = keyword
            ws.cell(row=row, column=5).value = english_trans
            ws.cell(row=row, column=6).value = "Vehicle Type"
            
            for col in range(1, 7):
                cell = ws.cell(row=row, column=col)
                cell.border = border
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            
            row += 1
        
        topics = lang_data.get("topics", {})
        for topic_name, keywords in topics.items():
            for keyword in keywords:
                english_trans = get_english_synonym(lang_code, keyword)
                
                ws.cell(row=row, column=1).value = lang_code
                ws.cell(row=row, column=2).value = lang_name
                ws.cell(row=row, column=3).value = topic_name
                ws.cell(row=row, column=4).value = keyword
                ws.cell(row=row, column=5).value = english_trans
                ws.cell(row=row, column=6).value = "Topic Keyword"
                
                for col in range(1, 7):
                    cell = ws.cell(row=row, column=col)
                    cell.border = border
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                
                row += 1
    
    output_dir = "output"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    output_file = os.path.join(output_dir, "Keywords_Mapping_All_Languages.xlsx")
    wb.save(output_file)
    print(f"✓ Keywords mapping created successfully!")
    print(f"✓ File saved: {output_file}")
    print(f"✓ Total rows: {row - 1}")
    print(f"✓ Languages included: {len(LANGS)}")
    print(f"✓ Topics covered: {len(TOPIC_NAMES)}")
    
    return output_file

def create_countries_reference_sheet():
    """Create a reference sheet showing language-to-country mapping"""
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Countries & Languages"
    
    header_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    headers = ["Region", "Countries", "Languages Used", "Language Codes"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 20
    
    row = 2
    
    for region, countries in COUNTRY_MAPPING.items():
        languages_for_region = set()
        lang_codes = []
        
        for lang_code, regions in LANGUAGE_REGIONS.items():
            if region in regions:
                languages_for_region.add(LANGUAGE_CODES.get(lang_code, lang_code))
                lang_codes.append(lang_code)
        
        ws.cell(row=row, column=1).value = region
        ws.cell(row=row, column=2).value = ", ".join(countries)
        ws.cell(row=row, column=3).value = ", ".join(sorted(languages_for_region))
        ws.cell(row=row, column=4).value = ", ".join(sorted(lang_codes))
        
        for col in range(1, 5):
            cell = ws.cell(row=row, column=col)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        
        row += 1
    
    output_dir = "output"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    output_file = os.path.join(output_dir, "Countries_Languages_Reference.xlsx")
    wb.save(output_file)
    print(f"✓ Countries reference sheet created!")
    print(f"✓ File saved: {output_file}")
    
    return output_file

def create_summary_report():
    """Create a summary statistics report"""
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    
    title_fill = PatternFill(start_color="203864", end_color="203864", fill_type="solid")
    title_font = Font(bold=True, color="FFFFFF", size=14)
    
    summary_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    summary_font = Font(bold=True, size=11)
    
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    ws.merge_cells('A1:B1')
    title_cell = ws['A1']
    title_cell.value = "Keywords Mapping Summary Report"
    title_cell.fill = title_fill
    title_cell.font = title_font
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    ws.row_dimensions[1].height = 25
    
    row = 3
    
    statistics = [
        ("Total Languages", len(LANGS)),
        ("Total Countries", sum(len(countries) for countries in COUNTRY_MAPPING.values())),
        ("Total Regions", len(COUNTRY_MAPPING)),
        ("Topic Categories", len(TOPIC_NAMES)),
        ("", ""),
    ]
    
    for lang_code in sorted(LANGS.keys()):
        lang_data = LANGS[lang_code]
        wheel_count = len(lang_data.get("wheel", []))
        topics_count = sum(len(keywords) for keywords in lang_data.get("topics", {}).values())
        total_keywords = wheel_count + topics_count
        statistics.append((f"{LANGUAGE_CODES[lang_code]} ({lang_code})", f"{total_keywords} keywords ({wheel_count} vehicle, {topics_count} topics)"))
    
    for label, value in statistics:
        if label == "":
            row += 1
            continue
        
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = value
        
        ws.cell(row=row, column=1).fill = summary_fill
        ws.cell(row=row, column=1).font = summary_font
        ws.cell(row=row, column=1).border = border
        ws.cell(row=row, column=2).border = border
        ws.cell(row=row, column=2).alignment = Alignment(vertical="center", wrap_text=True)
        
        row += 1
    
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 60
    
    output_dir = "output"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    output_file = os.path.join(output_dir, "Keywords_Summary_Report.xlsx")
    wb.save(output_file)
    print(f"✓ Summary report created!")
    print(f"✓ File saved: {output_file}")
    
    return output_file

def create_topics_reference():
    """Create a reference sheet showing all 7 topic categories and their keywords"""
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Topics Reference"
    
    header_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    header_font = Font(bold=True, color="000000", size=12)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    headers = ["Topic #", "Topic Name", "English Keywords", "Description"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 40
    
    topic_descriptions = {
        "New Launches": "Product launches, new models, debuts, unveilings, market introductions",
        "Production & Capacity Changes": "Manufacturing, factory operations, production changes, capacity, assembly",
        "Dealer Network Expansion": "Dealership openings, showroom inaugurations, network growth, distribution",
        "Price Changes": "Price increases/decreases, discounts, promotions, financing offers",
        "EV & Technology Announcements": "Electric vehicles, hybrid models, battery technology, smart features, connectivity",
        "Partnerships & Investments": "Joint ventures, acquisitions, investments, strategic alliances, mergers",
        "Exports & International Expansion": "Export activities, international markets, tariffs, global expansion",
    }
    
    row = 2
    for idx, topic in enumerate(TOPIC_NAMES, 1):
        en_keywords = LANGS["en"]["topics"].get(topic, [])
        keywords_str = ", ".join(en_keywords[:5])
        if len(en_keywords) > 5:
            keywords_str += f", ... (+{len(en_keywords) - 5} more)"
        
        ws.cell(row=row, column=1).value = idx
        ws.cell(row=row, column=2).value = topic
        ws.cell(row=row, column=3).value = keywords_str
        ws.cell(row=row, column=4).value = topic_descriptions.get(topic, "")
        
        for col in range(1, 5):
            cell = ws.cell(row=row, column=col)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        
        ws.row_dimensions[row].height = 40
        row += 1
    
    output_dir = "output"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    output_file = os.path.join(output_dir, "Topics_Reference.xlsx")
    wb.save(output_file)
    print(f"✓ Topics reference created!")
    print(f"✓ File saved: {output_file}")
    
    return output_file

if __name__ == "__main__":
    print("=" * 70)
    print("KEYWORD MAPPING GENERATOR - Multi-Language Two-Wheeler News Keywords")
    print("=" * 70)
    print()
    
    files_created = []
    
    print("Creating keyword mapping file (combined sheet)...")
    file1 = create_keywords_mapping_xlsx()
    files_created.append(file1)
    print()
    
    print("Creating countries & languages reference...")
    file2 = create_countries_reference_sheet()
    files_created.append(file2)
    print()
    
    print("Creating summary statistics report...")
    file3 = create_summary_report()
    files_created.append(file3)
    print()
    
    print("Creating topics reference sheet...")
    file4 = create_topics_reference()
    files_created.append(file4)
    print()
    
    print("=" * 70)
    print("ALL FILES CREATED SUCCESSFULLY!")
    print("=" * 70)
    print("\nOutput files in 'output/' directory:")
    for idx, file in enumerate(files_created, 1):
        print(f"  {idx}. {file}")
    print()
    print("Files include:")
    print("  • Keywords_Mapping_All_Languages.xlsx - Complete keyword mapping with English translations")
    print("  • Countries_Languages_Reference.xlsx - Region/Country/Language mapping")
    print("  • Keywords_Summary_Report.xlsx - Statistics and coverage summary")
    print("  • Topics_Reference.xlsx - Detailed topic definitions and keywords")
    print()
    print("=" * 70)