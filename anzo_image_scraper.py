"""
ANZO (anzousa.com) product IMAGE scraper - built from the Warn image scraper.

For every SKU in the input list (Inventory Number, e.g. ANZ-311302):
  1. Find the product on anzousa.com (the site's product list, then its search) and open the page
  2. Collect every image of the product's main slideshow (gallery order)
  3. Download each image, make it TARGET_WIDTH px wide, save as JPG
  4. ALL images go in ONE main folder (no folder per SKU), named with a DOUBLE underscore:
         ANZ-311302__1.jpg, ANZ-311302__2.jpg, ANZ-311302__3.jpg, ANZ-311354__1.jpg ...
  5. Write a log workbook: sheet 'Image Log' (one row per image) and sheet 'Parts To Review'
     (SKUs with 0-1 images, not found, or errors)

Needs:  pip install pillow requests pandas openpyxl seleniumbase
"""
import io
import os
import re
import json
import time
import base64
import shutil
import logging
import tempfile
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, urljoin, urlparse, urlunparse, parse_qsl, urlencode

import pandas as pd
import requests
from openpyxl.utils import get_column_letter
from PIL import Image
from seleniumbase import Driver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ---------------- Constants ----------------
WEBSITE = "https://anzousa.com/"
SEARCH_URL = "https://anzousa.com/search?q={mpn}&type=product"
INPUT_FILE = r"ANZO_Scrape_List_2026-11-04.xlsx"      # <-- full path if not in the same folder
ELEMENT_WAIT_TIME = 20
PAGE_PAUSE = 3
SEARCH_RESULT_WAIT = 12        # seconds to wait for search results
MAX_CANDIDATES = 5             # how many search results are checked per SKU
SAVE_LOG_EVERY = 5

# True  = first read the whole product list of the site (products.json: every SKU -> product URL), fast and exact.
#         If the site blocks it, the script falls back to the search automatically.
USE_PRODUCT_LIST = True
# True  = if the product list was read and the SKU is not in it, mark it "not found" straight away (faster)
# False = still try the site's suggestions and search page before giving up (slower, safest)
SKIP_SEARCH_IF_NOT_IN_LIST = False

# ---- Image settings ----
TARGET_WIDTH = 1800              # horizontal size in pixels (same rule as the Warn images)
UPSCALE_SMALLER_IMAGES = True    # True  = images narrower than TARGET_WIDTH are enlarged to exactly that width
                                 #         (the Anzo originals look like 1500 px wide, so they will be enlarged)
                                 # False = smaller images are saved at their own size and flagged in the log
OUTPUT_FORMAT = "jpg"            # "jpg" or "png"
JPEG_QUALITY = 95
SKIP_EXISTING = True             # do not download again if the file already exists
REVIEW_MAX_IMAGES = 1            # SKUs with this many images or fewer go on the "Parts To Review" sheet
                                 # (0 = no images, 1 = possibly a placeholder)
NAME_SEPARATOR = "__"            # ANZ-311302__1  (double underscore, as Matt asked)

OUTPUT_ROOT = Path(os.path.expanduser("~")) / "Desktop" / f"Anzo_Images_{datetime.now():%Y-%m-%d}"
LOG_FILE = Path(os.path.expanduser("~")) / "Desktop" / f"Anzo_Image_Log_{datetime.now():%Y-%m-%d}.xlsx"

# Images of the MAIN slideshow only (not the thumbnails strip, not videos, not the zoom-lightbox copies)
GALLERY_SELECTORS = [
    '[data-media-type="main"] li[data-media-type="image"] img',
    '[data-media-type="main"] li[class*="slide"] img',
    'media-slideshow li[data-media-type="image"] img',
    'li[class*="product"][class*="media-image"] img',
    'img[class*="product"][class*="media-image"]',
]
GALLERY_SLIDES_SELECTOR = '[data-media-type="main"] li[data-media-type]'     # used only to compare counts

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s: %(message)s",
    handlers=[logging.FileHandler("anzo_image_scraper.log"), logging.StreamHandler()],
)


# ---------------- Driver ----------------
def setup_driver():
    user_data_dir = tempfile.mkdtemp()
    driver = Driver(uc=True)
    driver.user_data_dir = user_data_dir
    driver.set_page_load_timeout(30)
    return driver


def wait_for_captcha(driver):
    print("If a CAPTCHA / bot check appears, solve it manually.")
    input("Press Enter to continue...")


def cleanup(driver):
    try:
        driver.quit()
    except Exception as e:
        print(f"Driver quit error (can be ignored): {e}")
    if hasattr(driver, "user_data_dir") and os.path.exists(driver.user_data_dir):
        shutil.rmtree(driver.user_data_dir, ignore_errors=True)


# ---------------- Input ----------------
def load_mpn_list(path):
    """
    Reads the SKU list. Needs an 'Inventory Number' column (e.g. ANZ-311302).
    MPN = the file's 'MPN' column if it has one, otherwise the number after the last hyphen of the Inventory Number.
    """
    df = pd.read_excel(path, dtype=str)
    inv_col = next((c for c in df.columns if "inventory" in str(c).lower()), df.columns[0])
    mpn_col = next((c for c in df.columns if str(c).strip().lower() == "mpn"), None)

    out = pd.DataFrame()
    out["Inventory Number"] = df[inv_col].str.strip()
    derived = out["Inventory Number"].str.extract(r"-(\d+)$")[0].fillna(out["Inventory Number"])
    out["MPN"] = df[mpn_col].str.strip() if mpn_col else derived
    out = out.dropna(subset=["Inventory Number", "MPN"])
    out = out[out["MPN"].str.len() > 0]
    return out.drop_duplicates(subset="MPN").reset_index(drop=True)


# ---------------- Helpers ----------------
def clean(text):
    return re.sub(r"\s+", " ", (text or "").replace("\xa0", " ")).strip()


def safe_name(text):
    """Safe for a Windows file name (removes < > : " / \\ | ? * and trailing dots/spaces)."""
    t = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", clean(text))
    return t.strip(" .") or "Untitled"


def product_path(url):
    """https://anzousa.com/collections/x/products/abc?v=1 -> https://anzousa.com/products/abc"""
    p = urlparse(urljoin(WEBSITE, url))
    path = re.sub(r"^/collections/[^/]+/products/", "/products/", p.path)
    return f"{p.scheme}://{p.netloc}{path}"


def sku_matches(sku, mpn):
    """Site SKU equals the MPN, or the MPN plus letters only (e.g. 111637X for 111637)."""
    s, m = (sku or "").strip().upper(), (mpn or "").strip().upper()
    if not s or not m:
        return False
    return s == m or (s.startswith(m) and not re.search(r"\d", s[len(m):]))


# ---------------- Finding the product (same method as the Anzo product scraper) ----------------
_PRODUCTS_JSON_JS = r"""
const done = arguments[arguments.length - 1];
(async () => {
  const out = [];
  for (let page = 1; page <= 200; page++) {
    const r = await fetch('/products.json?limit=250&page=' + page, {credentials: 'include'});
    if (!r.ok) throw new Error('HTTP ' + r.status);
    const j = await r.json();
    if (!j.products || !j.products.length) break;
    for (const p of j.products) for (const v of (p.variants || [])) if (v.sku) out.push([v.sku, p.handle]);
  }
  done(JSON.stringify(out));
})().catch(e => done('ERR:' + e));
"""

_PRODUCT_SKUS_JS = r"""
const url = arguments[0], done = arguments[arguments.length - 1];
fetch(url.split('?')[0] + '.js', {credentials: 'include'})
  .then(r => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
  .then(p => done(JSON.stringify((p.variants || []).map(v => v.sku || ''))))
  .catch(e => done('ERR:' + e));
"""

_SUGGEST_JS = """
const q = arguments[0], done = arguments[arguments.length - 1];
fetch('/search/suggest.json?q=' + encodeURIComponent(q) + '&resources[type]=product&resources[limit]=10',
      {credentials: 'include'})
  .then(r => r.json()).then(j => done(JSON.stringify(j))).catch(e => done('ERR:' + e));
"""


def load_sku_index(driver):
    """{SKU (upper case): product URL} for the whole shop. Empty dict if the site does not allow it."""
    try:
        driver.set_script_timeout(300)
        res = driver.execute_async_script(_PRODUCTS_JSON_JS)
        if not isinstance(res, str) or res.startswith("ERR:"):
            logging.warning(f"Product list could not be read ({res}); the search will be used instead")
            return {}
        return {sku.strip().upper(): f"{WEBSITE}products/{handle}" for sku, handle in json.loads(res)}
    except Exception as e:
        logging.warning(f"Product list could not be read ({e}); the search will be used instead")
        return {}


def url_from_sku_index(sku_index, mpn):
    """Exact SKU first, then SKU = MPN + letters (e.g. 111637X)."""
    m = (mpn or "").strip().upper()
    if m in sku_index:
        return sku_index[m]
    for sku, url in sku_index.items():
        if sku_matches(sku, m):
            return url
    return ""


def _suggest_urls(driver, mpn):
    try:
        driver.set_script_timeout(30)
        res = driver.execute_async_script(_SUGGEST_JS, mpn)
        if not isinstance(res, str) or res.startswith("ERR:"):
            return []
        products = json.loads(res).get("resources", {}).get("results", {}).get("products", [])
        return [product_path(p["url"]) for p in products if p.get("url")]
    except Exception as e:
        logging.warning(f"Predictive search failed for {mpn}: {e}")
        return []


def _search_page_urls(driver, mpn):
    """Product links in the results of the search page (main content only)."""
    driver.get(SEARCH_URL.format(mpn=quote(mpn)))
    time.sleep(PAGE_PAUSE)
    if "/products/" in driver.current_url:
        return [product_path(driver.current_url)]
    sel = '#MainContent a[href*="/products/"], main a[href*="/products/"]'
    try:
        WebDriverWait(driver, SEARCH_RESULT_WAIT).until(
            lambda d: len(d.find_elements(By.CSS_SELECTOR, sel)) > 0)
    except TimeoutException:
        return []
    urls = []
    for a in driver.find_elements(By.CSS_SELECTOR, sel):
        u = product_path(a.get_attribute("href") or "")
        if "/products/" in u and u not in urls:
            urls.append(u)
    return urls


def verify_candidate(driver, url, mpn):
    """True / False if the product's SKUs say so, None when it cannot be checked."""
    try:
        driver.set_script_timeout(30)
        res = driver.execute_async_script(_PRODUCT_SKUS_JS, url)
        if not isinstance(res, str) or res.startswith("ERR:"):
            return None
        return any(sku_matches(s, mpn) for s in json.loads(res))
    except Exception:
        return None


def find_product_url(driver, mpn, sku_index=None):
    """
    (product URL, status) for an MPN: 1) the shop's product list, 2) predictive search then the search page
    (each result checked by its SKUs), 3) last resort: the first unchecked result (the caller discards it
    if the opened page does not contain the MPN).
    """
    if sku_index:
        hit = url_from_sku_index(sku_index, mpn)
        if hit:
            return hit, "OK (product list)"
        if SKIP_SEARCH_IF_NOT_IN_LIST:
            return "", "Not in the site's product list (search skipped)"

    tried, unchecked = set(), ""
    for source in (_suggest_urls, _search_page_urls):
        for url in source(driver, mpn)[:MAX_CANDIDATES]:
            if url in tried:
                continue
            tried.add(url)
            ok = verify_candidate(driver, url, mpn)
            if ok:
                return url, "OK (search)"
            if ok is None and not unchecked:
                unchecked = url
    if unchecked:
        return unchecked, "CHECK - SKU could not be verified (first result used)"
    logging.warning(f"MPN {mpn}: no matching product found")
    return "", "No search results"


def scrape_title(driver):
    for sel in ("h1.product_title", "div.product_info-container h1", "h1"):
        try:
            return clean(driver.find_element(By.CSS_SELECTOR, sel).get_attribute("textContent"))
        except NoSuchElementException:
            continue
    return ""


# ---------------- Image URLs ----------------
def _best_img_url(img):
    """src of the <img>; if it is empty / a placeholder (lazy loading) use a lazy attribute or the largest srcset entry."""
    src = (img.get_attribute("src") or "").strip()
    if src and not src.startswith("data:"):
        return src
    for attr in ("data-splide-lazy", "data-src", "data-original", "data-lazy"):
        v = (img.get_attribute(attr) or "").strip()
        if v and not v.startswith("data:"):
            return v
    srcset = (img.get_attribute("srcset") or img.get_attribute("data-srcset")
              or img.get_attribute("data-splide-lazy-srcset") or "").strip()
    best, best_w = "", -1
    for part in srcset.split(","):
        bits = part.strip().split()
        if not bits:
            continue
        w = int(bits[1][:-1]) if len(bits) > 1 and bits[1][:-1].isdigit() else 0
        if w > best_w:
            best, best_w = bits[0], w
    return best


def sized_url(url, width=TARGET_WIDTH):
    """Same image, but asking Shopify's CDN for the requested width (keeps the ?v= version tag)."""
    u = urlparse(urljoin(WEBSITE, url))
    q = dict(parse_qsl(u.query))
    q["width"] = str(width)
    q.pop("height", None)
    return urlunparse(u._replace(query=urlencode(q)))


def _image_key(url):
    """Identity of an image ignoring width/height params, used to drop duplicates."""
    u = urlparse(urljoin(WEBSITE, url))
    return u.path + "?v=" + dict(parse_qsl(u.query)).get("v", "")


def get_gallery_image_urls(driver):
    """Returns (list of image URLs in slideshow order, number of slides the slideshow shows)."""
    try:
        slides = len(driver.find_elements(By.CSS_SELECTOR, GALLERY_SLIDES_SELECTOR))
    except Exception:
        slides = 0
    expected = slides or None

    urls, seen = [], set()
    for sel in GALLERY_SELECTORS:
        for img in driver.find_elements(By.CSS_SELECTOR, sel):
            raw = _best_img_url(img)
            if not raw:
                continue
            key = _image_key(raw)
            if key in seen:
                continue
            seen.add(key)
            urls.append(sized_url(raw))
        if urls:            # first selector that finds images wins
            break
    return urls, expected


# ---------------- Download + resize ----------------
def fetch_bytes(driver, url):
    """
    Download through the browser itself (same site, same cookies, so no bot-blocking).
    Falls back to requests with the browser's cookies and user agent.
    """
    js = """
    const url = arguments[0], done = arguments[arguments.length - 1];
    fetch(url, {credentials: 'include'})
      .then(r => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.arrayBuffer(); })
      .then(buf => {
        const a = new Uint8Array(buf); let s = ''; const CH = 0x8000;
        for (let i = 0; i < a.length; i += CH) s += String.fromCharCode.apply(null, a.subarray(i, i + CH));
        done(btoa(s));
      })
      .catch(e => done('ERR:' + e));
    """
    try:
        driver.set_script_timeout(60)
        res = driver.execute_async_script(js, url)
        if isinstance(res, str) and not res.startswith("ERR:"):
            return base64.b64decode(res)
        logging.warning(f"Browser fetch failed ({res}); trying requests: {url}")
    except Exception as e:
        logging.warning(f"Browser fetch error ({e}); trying requests: {url}")

    ua = driver.execute_script("return navigator.userAgent")
    cookies = {c["name"]: c["value"] for c in driver.get_cookies()}
    r = requests.get(url, headers={"User-Agent": ua, "Referer": WEBSITE}, cookies=cookies, timeout=60)
    r.raise_for_status()
    return r.content


def process_image(raw):
    """
    Open the downloaded bytes (webp/jpg/png), flatten any transparency onto white,
    make it exactly TARGET_WIDTH wide (keeping proportions) and return (PIL image, note).
    """
    img = Image.open(io.BytesIO(raw))
    img.load()
    src_w, src_h = img.size

    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        rgba = img.convert("RGBA")
        bg = Image.new("RGB", rgba.size, (255, 255, 255))
        bg.paste(rgba, mask=rgba.split()[-1])
        img = bg
    elif img.mode != "RGB":
        img = img.convert("RGB")

    note = ""
    if src_w != TARGET_WIDTH:
        if src_w > TARGET_WIDTH or UPSCALE_SMALLER_IMAGES:
            new_h = max(1, round(src_h * TARGET_WIDTH / src_w))
            img = img.resize((TARGET_WIDTH, new_h), Image.LANCZOS)
            note = f"resized {src_w}px -> {TARGET_WIDTH}px" + (" (enlarged)" if src_w < TARGET_WIDTH else "")
        else:
            note = f"smaller than {TARGET_WIDTH}px ({src_w}px), not enlarged"
    return img, note


def save_image(img, path):
    if OUTPUT_FORMAT.lower() == "png":
        img.save(path, "PNG", optimize=True)
    else:
        img.save(path, "JPEG", quality=JPEG_QUALITY, optimize=True)


def image_path(inv, index):
    """ONE main folder, no subfolders:  <root>/ANZ-311302__1.jpg, ANZ-311302__2.jpg ..."""
    ext = "png" if OUTPUT_FORMAT.lower() == "png" else "jpg"
    return OUTPUT_ROOT / f"{safe_name(inv)}{NAME_SEPARATOR}{index}.{ext}"


# ---------------- Log ----------------
def _format_log_sheet(ws, max_width=60):
    """Autofit column widths, filter buttons on the header row, freeze row 1 and column A (cell B2)."""
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    for col_idx, column_cells in enumerate(ws.columns, start=1):
        longest = max((len(str(c.value)) for c in column_cells if c.value is not None), default=0)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(longest + 2, 8), max_width)


def save_log(log_rows, review_rows=None):
    """Sheet 'Image Log' (one row per image) and sheet 'Parts To Review' (0-1 images, not found, errors)."""
    if not log_rows and not review_rows:
        return
    log_cols = ["Inventory Number", "MPN", "Product Title", "Image #", "File Name",
                "Width", "Height", "Status", "Note", "Source URL"]
    review_cols = ["Inventory Number", "MPN", "Product Title", "Image Count", "Review Status", "Product URL"]
    with pd.ExcelWriter(LOG_FILE, engine="openpyxl") as writer:
        pd.DataFrame(log_rows or [], columns=log_cols).to_excel(writer, index=False, sheet_name="Image Log")
        pd.DataFrame(review_rows or [], columns=review_cols).to_excel(writer, index=False, sheet_name="Parts To Review")
        for ws in writer.book.worksheets:
            _format_log_sheet(ws)
    print(f"Log saved: {LOG_FILE}  ({len(review_rows or [])} SKUs to review)")


# ---------------- Main ----------------
def main(test_mpns=None):
    """
    test_mpns: optional list of MPN strings to run instead of the full list.
    Example: main(test_mpns=["111506", "511032", "211140"])
    """
    driver = None
    log_rows = []
    review_rows = []
    try:
        mpns = load_mpn_list(INPUT_FILE)
        if test_mpns:
            mpns = mpns[mpns["MPN"].isin(test_mpns)].reset_index(drop=True)
        print(f"Loaded {len(mpns)} SKUs")
        OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

        driver = setup_driver()
        driver.get(WEBSITE)
        wait_for_captcha(driver)

        sku_index = load_sku_index(driver) if USE_PRODUCT_LIST else {}
        print(f"Product list: {len(sku_index)} SKUs read from the site"
              + ("" if sku_index else " (not available - the search will be used)"))

        for i, r in mpns.iterrows():
            mpn, inv = r["MPN"], r["Inventory Number"]
            print(f"[{i + 1}/{len(mpns)}] {inv} (MPN {mpn})")
            title, url, count = "", "", 0

            def flag(review_status):
                review_rows.append({"Inventory Number": inv, "MPN": mpn, "Product Title": title,
                                    "Image Count": count, "Review Status": review_status, "Product URL": url})

            try:
                url, status = find_product_url(driver, mpn, sku_index)
                if not url:
                    logging.warning(f"{inv}: {status}")
                    log_rows.append({"Inventory Number": inv, "MPN": mpn, "Status": status})
                    flag("Product not found - check")
                else:
                    if not status.startswith("OK"):
                        logging.warning(f"{inv} needs a manual check: {status} {url}")

                    driver.get(url)
                    time.sleep(PAGE_PAUSE)
                    WebDriverWait(driver, ELEMENT_WAIT_TIME).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "h1"))
                    )
                    title = scrape_title(driver)

                    if mpn not in driver.page_source:
                        logging.error(f"{inv}: opened page '{title}' does not contain the SKU, "
                                      f"skipped (wrong product?): {url}")
                        log_rows.append({"Inventory Number": inv, "MPN": mpn, "Product Title": title,
                                         "Status": "SKU not found on page - skipped (wrong product?)"})
                        flag("Wrong product opened - check")
                    else:
                        urls, expected = get_gallery_image_urls(driver)
                        count = len(urls)

                        if count <= REVIEW_MAX_IMAGES:
                            flag("No images" if count == 0 else f"{count} image - possible placeholder")
                        if not urls:
                            logging.warning(f"{inv}: no gallery images found")
                            log_rows.append({"Inventory Number": inv, "MPN": mpn, "Product Title": title,
                                             "Status": "No images found"})
                        else:
                            if expected is not None and expected != count:
                                logging.warning(f"{inv}: slideshow has {expected} slides, found {count} images")
                            print(f"   -> {title[:60]} | {count} images (slideshow has {expected} slides)")

                            for idx, img_url in enumerate(urls, start=1):
                                path = image_path(inv, idx)
                                row = {"Inventory Number": inv, "MPN": mpn, "Product Title": title,
                                       "Image #": idx, "File Name": path.name, "Source URL": img_url}
                                try:
                                    if SKIP_EXISTING and path.exists():
                                        with Image.open(path) as existing:
                                            row.update(Width=existing.width, Height=existing.height,
                                                       Status="Skipped (already exists)")
                                    else:
                                        img, note = process_image(fetch_bytes(driver, img_url))
                                        save_image(img, path)
                                        row.update(Width=img.width, Height=img.height, Status="OK", Note=note)
                                        if img.width != TARGET_WIDTH:
                                            row["Status"] = f"CHECK - width is {img.width}, not {TARGET_WIDTH}"
                                    print(f"      {path.name}  {row.get('Width')}x{row.get('Height')}  {row['Status']}")
                                except Exception as e:
                                    logging.error(f"{inv} image {idx} failed: {e}")
                                    row["Status"] = f"ERROR: {e}"
                                log_rows.append(row)
            except Exception as e:
                logging.error(f"{inv} failed: {e}")
                log_rows.append({"Inventory Number": inv, "MPN": mpn, "Product Title": title,
                                 "Status": f"ERROR: {e}"})
                flag("ERROR - check")

            if (i + 1) % SAVE_LOG_EVERY == 0:
                save_log(log_rows, review_rows)
    finally:
        save_log(log_rows, review_rows)
        if driver:
            cleanup(driver)


if __name__ == "__main__":
    # ---- TEST MODE: the 3 test SKUs. Comment the line below and uncomment main() to run the full list ----
    # main(test_mpns=["311302", "311354"])
    main()