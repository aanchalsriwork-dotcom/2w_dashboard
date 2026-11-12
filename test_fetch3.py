"""Scrapes name, price, image from fasterwheeler.com/bz/home -> Excel with embedded images.
pip install requests beautifulsoup4 openpyxl pillow"""
import io, re, time, requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font
from PIL import Image

BASE = "https://fasterwheeler.com"
HOME = f"{BASE}/bz/home"
H = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": BASE + "/",
}
S = requests.Session(); S.headers.update(H)

r = S.get(HOME, timeout=30)
print("Status:", r.status_code, "| Final URL:", r.url, "| Length:", len(r.text))
soup = BeautifulSoup(r.text, "html.parser")

LINK = re.compile(r"/(bikes|electric-bike)/[a-z0-9-]+-\d+/?$", re.I)
PRICE = re.compile(r"(?:US\$|\$|Rs\.?|INR|₹|€|£)\s?[\d,]+(?:\.\d+)?|[\d,]{3,}\s?(?:USD|INR)")

items, seen = [], set()
for a in soup.find_all("a", href=LINK):
    href = a["href"]; url = href if href.startswith("http") else BASE + href
    if url in seen: continue
    # look for the price in the link, then in its parent/grandparent
    price_txt = ""
    for node in (a, a.parent, a.parent.parent if a.parent else None):
        if node is None: continue
        m = PRICE.search(node.get_text(" ", strip=True))
        if m: price_txt = m.group(); break
    name = (a.find(["strong", "b"]) or a).get_text(" ", strip=True)
    if not name:
        img = a.find("img"); name = (img.get("alt") or img.get("title") or "") if img else ""
    seen.add(url)
    num = re.sub(r"[^\d.]", "", price_txt)
    items.append((name, price_txt, float(num) if num else None, url))

print("Product links found:", len(items))
if not items:
    open("debug.html", "w", encoding="utf-8").write(r.text)
    print("Nothing found. Saved page to debug.html - open it and check what it contains.")
    print(r.text[:500])
    raise SystemExit

wb = Workbook(); ws = wb.active; ws.title = "Vehicles"
ws.append(["Vehicle Name", "Price (as shown)", "Price (number)", "Product Page", "Image URL", "Image"])
for col, w in zip("ABCDEF", [38, 16, 14, 55, 60, 22]): ws.column_dimensions[col].width = w

for i, (name, ptxt, pnum, url) in enumerate(items, start=2):
    img_url = ""
    try:
        p = BeautifulSoup(S.get(url, timeout=30).text, "html.parser")
        og = p.find("meta", attrs={"property": "og:image"}) or p.find("meta", attrs={"name": "twitter:image"})
        img_url = og["content"] if og and og.get("content") else ""
    except Exception as e:
        print("page fail", url, e)
    ws.append([name, ptxt, pnum, url, img_url])
    ws.row_dimensions[i].height = 75
    if img_url:
        try:
            im = Image.open(io.BytesIO(S.get(img_url, timeout=30).content)).convert("RGB")
            im.thumbnail((140, 100)); buf = io.BytesIO(); im.save(buf, "PNG"); buf.seek(0)
            ws.add_image(XLImage(buf), f"F{i}")
        except Exception as e:
            print("img fail", img_url, e)
    print(f"{i-1}/{len(items)} {name}")
    time.sleep(0.5)

for c in ws[1]: c.font = Font(bold=True)
wb.save("fasterwheeler_vehicles_with_images.xlsx"); print(len(items), "vehicles saved")