"""
WARN.com product IMAGE scraper (built on the same framework as warn_scraper_full.py)

For every MPN in Warn_Scrape_List_*.xlsx:
  1. Search the MPN on warn.com and open the product page
  2. Collect every image in the product gallery (main slider only, in gallery order)
  3. Download each image at 1800 px wide (Matt's requirement), convert to JPG
  4. Save in a subfolder per Part Number, as  <Product Title>_<Part Number>.jpg   (extra images: _2, _3, ...)
     Part Number = our Inventory Number (e.g. WAR-17801), same as the other sheets.
  5. Create the part's folder for EVERY part (even with no images) and write a log workbook:
     sheet 'Image Log' (one row per image) and sheet 'Parts To Review' (parts with 0-1 images,
     not found, or errors)

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
WEBSITE = "https://www.warn.com/"
SEARCH_URL = "https://www.warn.com/search?q={mpn}"
INPUT_FILE = r"Warn_Scrape_List_2026-10-27.xlsx"   # <-- full path if not in the same folder
ELEMENT_WAIT_TIME = 20
PAGE_PAUSE = 3
SEARCH_RESULT_WAIT = 12       # seconds to wait for the search results to show the MPN
SAVE_LOG_EVERY = 5

# ---- Image settings ----
TARGET_WIDTH = 1800              # horizontal size in pixels (Matt)
UPSCALE_SMALLER_IMAGES = True    # True  = images smaller than 1800 px are enlarged to exactly 1800 wide
                                 #         (softer picture, but every file is 1800 wide)
                                 # False = smaller images are saved at their own size and flagged in the log
OUTPUT_FORMAT = "jpg"            # "jpg" or "png"
JPEG_QUALITY = 95
SUBFOLDER_PER_PART = True        # True  = one subfolder per Part Number (e.g. WAR-17801) with that part's images
                                 # False = all images together in one folder
SKIP_EXISTING = True             # do not download again if the file already exists
REVIEW_MAX_IMAGES = 1            # parts with this many images or fewer go on the "Parts To Review" sheet
                                 # (0 = no images, 1 = probably a placeholder picture, as Matt suggested)

OUTPUT_ROOT = Path(os.path.expanduser("~")) / "Desktop" / f"Warn_Images_{datetime.now():%Y-%m-%d}"
LOG_FILE = Path(os.path.expanduser("~")) / "Desktop" / f"Warn_Image_Log_{datetime.now():%Y-%m-%d}.xlsx"

# Main gallery images only (NOT the thumbnails strip and NOT the zoom-modal copy of the same images)
GALLERY_SELECTORS = [
    'media-gallery ul[id^="Slider-Gallery"] li.product__media-item img',
    'media-gallery ul[id^="Slider-Gallery"] img',
    'ul[id^="Slider-Gallery"] img',
    "div.product__media img",
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s: %(message)s",
    handlers=[logging.FileHandler("warn_image_scraper.log"), logging.StreamHandler()],
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


# ---------------- Input / helpers ----------------
def load_mpn_list(path):
    df = pd.read_excel(path, dtype=str)
    inv_col = next((c for c in df.columns if "inventory" in c.lower()), df.columns[0])
    df = df[[inv_col]].copy()
    df.columns = ["Inventory Number"]
    df = df.dropna(subset=["Inventory Number"])
    df["Inventory Number"] = df["Inventory Number"].str.strip()
    # Extract the numeric portion after the last hyphen (e.g. WAR-102255 -> 102255)
    df["MPN"] = df["Inventory Number"].str.extract(r"-(\d+)$")[0].fillna(df["Inventory Number"])
    return df.drop_duplicates(subset="MPN").reset_index(drop=True)


def clean(text):
    return re.sub(r"\s+", " ", (text or "").replace("\xa0", " ")).strip()


def safe_name(text):
    """Make text safe for a Windows file/folder name (removes < > : " / \\ | ? * and trailing dots/spaces)."""
    t = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", clean(text))
    return t.strip(" .") or "Untitled"


# Product index from the site's own sitemap: every product URL, independent of the search page
_SITEMAP_JS = r"""
const done = arguments[arguments.length - 1];
const getText = u => fetch(u, {credentials: 'include'}).then(r => {
  if (!r.ok) throw new Error('HTTP ' + r.status + ' ' + u);
  return r.text();
});
const locs = t => Array.from(t.matchAll(/<loc>([^<]+)<\/loc>/g)).map(m => m[1].replace(/&amp;/g, '&'));
getText('/sitemap.xml').then(async idx => {
  const maps = locs(idx).filter(u => u.indexOf('sitemap_products') !== -1);
  let urls = [];
  for (const m of maps) {
    const t = await getText(m);
    urls = urls.concat(locs(t).filter(u => u.indexOf('/products/') !== -1));
  }
  done(JSON.stringify(urls));
}).catch(e => done('ERR:' + e));
"""


def load_product_url_index(driver):
    """List of every product URL on the site (from sitemap.xml). Empty list if it cannot be read."""
    try:
        driver.set_script_timeout(120)
        res = driver.execute_async_script(_SITEMAP_JS)
        if not isinstance(res, str) or res.startswith("ERR:"):
            logging.warning(f"Product sitemap could not be read ({res}); the search page will be used instead")
            return []
        return [u.split("?")[0] for u in json.loads(res)]
    except Exception as e:
        logging.warning(f"Product sitemap could not be read ({e}); the search page will be used instead")
        return []


def url_from_index(url_index, mpn):
    """Product URLs whose handle contains the MPN as its own word (e.g. m12-winch-17801), best match first."""
    pat = re.compile(r"(?:^|-)" + re.escape(mpn.lower()) + r"(?:-|$)")
    hits = [u for u in url_index if pat.search(u.split("/products/")[-1].lower())]
    hits.sort(key=lambda u: (not u.lower().rstrip("/").endswith("-" + mpn.lower()), len(u)))
    return hits


# Shopify predictive-search endpoint: lists the products that match the typed MPN
_SUGGEST_JS = """
const q = arguments[0], done = arguments[arguments.length - 1];
fetch('/search/suggest.json?q=' + encodeURIComponent(q) + '&resources[type]=product&resources[limit]=10',
      {credentials: 'include'})
  .then(r => r.json()).then(j => done(JSON.stringify(j))).catch(e => done('ERR:' + e));
"""


def _product_links(driver, mpn):
    """(all product links on the page, those whose URL contains the MPN) - snapshot of the page right now."""
    hrefs = []
    for a in driver.find_elements(By.CSS_SELECTOR, 'a[href*="/products/"]'):
        h = (a.get_attribute("href") or "").split("?")[0]
        if "/products/" in h and h not in hrefs:
            hrefs.append(h)
    return hrefs, [h for h in hrefs if mpn in h.split("/products/")[-1]]


def _suggest_urls(driver, mpn):
    """Product URLs returned by the site's own predictive search for this MPN."""
    try:
        driver.set_script_timeout(30)
        res = driver.execute_async_script(_SUGGEST_JS, mpn)
        if not isinstance(res, str) or res.startswith("ERR:"):
            return []
        products = json.loads(res).get("resources", {}).get("results", {}).get("products", [])
        urls = [urljoin(WEBSITE, p.get("url", "")).split("?")[0] for p in products if p.get("url")]
        return [u for u in urls if mpn in u.split("/products/")[-1]]
    except Exception as e:
        logging.warning(f"Predictive search failed for {mpn}: {e}")
        return []


def find_product_url(driver, mpn, url_index=None):
    """
    Return (product URL, status) for an MPN. Tried in this order:
      1. the site's product sitemap (url_index) - does not depend on the search page at all
      2. the search page, waiting until a result whose URL contains the MPN appears
      3. the site's predictive search
      4. last resort: first search result (the caller discards the data if the page lacks the MPN)
    """
    if url_index:
        hits = url_from_index(url_index, mpn)
        if hits:
            if len(hits) > 1:
                logging.warning(f"MPN {mpn}: {len(hits)} products match in the sitemap, using {hits[0]} (others: {hits[1:3]})")
            return hits[0], "OK (sitemap)"

    driver.get(SEARCH_URL.format(mpn=quote(mpn)))
    time.sleep(PAGE_PAUSE)

    if "/products/" in driver.current_url:          # search redirected straight to the product
        return driver.current_url.split("?")[0], "OK (direct)"

    try:
        WebDriverWait(driver, SEARCH_RESULT_WAIT).until(lambda d: _product_links(d, mpn)[1])
    except TimeoutException:
        pass
    hrefs, exact = _product_links(driver, mpn)
    if exact:
        return exact[0], "OK"

    suggested = _suggest_urls(driver, mpn)
    if suggested:
        return suggested[0], "OK (predictive search)"

    logging.warning(f"MPN {mpn}: not found in sitemap, search results or predictive search")
    if hrefs:
        return hrefs[0], "CHECK - MPN not in URL (first result used)"
    return "", "No search results"


def scrape_title(driver):
    for sel in ("div.product_title h1", "div.product_info-container h1", "h1"):
        try:
            return clean(driver.find_element(By.CSS_SELECTOR, sel).get_attribute("textContent"))
        except NoSuchElementException:
            continue
    return ""


# ---------------- Image URLs ----------------
def _best_img_url(img):
    """src of the <img>; if it is empty / a placeholder (lazy loading) use the largest srcset entry."""
    src = (img.get_attribute("src") or "").strip()
    if src and not src.startswith("data:"):
        return src
    for attr in ("data-src", "data-original"):
        v = (img.get_attribute(attr) or "").strip()
        if v and not v.startswith("data:"):
            return v
    srcset = (img.get_attribute("srcset") or img.get_attribute("data-srcset") or "").strip()
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
    """Returns (list of 1800-wide image URLs in gallery order, number the gallery says it has)."""
    expected = None
    try:
        gal = driver.find_element(By.CSS_SELECTOR, "media-gallery")
        v = (gal.get_attribute("data-num-images") or "").strip()
        expected = int(v) if v.isdigit() else None
    except NoSuchElementException:
        pass

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


def image_paths(title, inv, index):
    """
    <root>/<Part Number>/Title_Part.jpg for the first image, Title_Part_2.jpg, Title_Part_3.jpg ... for the rest.
    Part Number = our Inventory Number, e.g. Warn_Images_<date>/WAR-17801/M12 Winch_WAR-17801.jpg
    """
    base = f"{safe_name(title)}_{safe_name(inv)}"
    folder = OUTPUT_ROOT / safe_name(inv) if SUBFOLDER_PER_PART else OUTPUT_ROOT
    ext = "png" if OUTPUT_FORMAT.lower() == "png" else "jpg"
    name = f"{base}.{ext}" if index == 1 else f"{base}_{index}.{ext}"
    return folder, folder / name


# ---------------- Log ----------------
def _format_log_sheet(ws, max_width=60):
    """Autofit column widths, filter buttons on the header row, freeze row 1 and column A (cell B2)."""
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    for col_idx, column_cells in enumerate(ws.columns, start=1):
        longest = max((len(str(c.value)) for c in column_cells if c.value is not None), default=0)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(longest + 2, 8), max_width)


def save_log(log_rows, review_rows=None):
    """Sheet 'Image Log' (one row per image) and sheet 'Parts To Review' (parts with 0-1 images, not found, errors)."""
    if not log_rows and not review_rows:
        return
    log_cols = ["Inventory Number", "MPN", "Product Title", "Image #", "File Name",
                "Width", "Height", "Status", "Note", "Source URL"]
    review_cols = ["Inventory Number", "MPN", "Product Title", "Image Count", "Review Status", "Folder", "Product URL"]
    with pd.ExcelWriter(LOG_FILE, engine="openpyxl") as writer:
        pd.DataFrame(log_rows or [], columns=log_cols).to_excel(writer, index=False, sheet_name="Image Log")
        pd.DataFrame(review_rows or [], columns=review_cols).to_excel(writer, index=False, sheet_name="Parts To Review")
        for ws in writer.book.worksheets:
            _format_log_sheet(ws)
    print(f"Log saved: {LOG_FILE}  ({len(review_rows or [])} parts to review)")


# ---------------- Main ----------------
def main(test_mpns=None):
    """
    test_mpns: optional list of MPN strings to run instead of the full list.
    Example: main(test_mpns=["28500", "17801", "26502"])
    """
    driver = None
    log_rows = []
    review_rows = []
    try:
        mpns = load_mpn_list(INPUT_FILE)
        if test_mpns:
            mpns = mpns[mpns["MPN"].isin(test_mpns)].reset_index(drop=True)
        print(f"Loaded {len(mpns)} MPNs")
        OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

        driver = setup_driver()
        driver.get(WEBSITE)
        wait_for_captcha(driver)

        url_index = load_product_url_index(driver)
        print(f"Product index: {len(url_index)} product URLs read from the sitemap")

        for i, r in mpns.iterrows():
            mpn, inv = r["MPN"], r["Inventory Number"]
            print(f"[{i + 1}/{len(mpns)}] MPN {mpn} ({inv})")
            title, url, count = "", "", 0

            # The part's folder is ALWAYS created, even when there are no images (data team reviews every folder)
            part_folder = OUTPUT_ROOT / safe_name(inv) if SUBFOLDER_PER_PART else OUTPUT_ROOT
            part_folder.mkdir(parents=True, exist_ok=True)

            def flag(review_status):
                review_rows.append({"Inventory Number": inv, "MPN": mpn, "Product Title": title,
                                    "Image Count": count, "Review Status": review_status,
                                    "Folder": part_folder.name if SUBFOLDER_PER_PART else "", "Product URL": url})

            try:
                url, status = find_product_url(driver, mpn, url_index)
                if not url:
                    logging.warning(f"MPN {mpn}: {status}")
                    log_rows.append({"Inventory Number": inv, "MPN": mpn, "Status": status})
                    flag("Product not found - check")
                else:
                    if not status.startswith("OK"):
                        logging.warning(f"MPN {mpn} needs a manual check: {status} {url}")

                    driver.get(url)
                    time.sleep(PAGE_PAUSE)
                    WebDriverWait(driver, ELEMENT_WAIT_TIME).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "h1"))
                    )
                    title = scrape_title(driver)

                    if mpn not in driver.page_source:
                        logging.error(f"MPN {mpn} ({inv}): opened page '{title}' does not contain the MPN, "
                                      f"skipped (wrong product?): {url}")
                        log_rows.append({"Inventory Number": inv, "MPN": mpn, "Product Title": title,
                                         "Status": "MPN not found on page - skipped (wrong product?)"})
                        flag("Wrong product opened - check")
                    else:
                        urls, expected = get_gallery_image_urls(driver)
                        count = len(urls)

                        if count <= REVIEW_MAX_IMAGES:
                            flag("No images" if count == 0 else
                                 f"{count} image - possible placeholder")
                        if not urls:
                            logging.warning(f"MPN {mpn} ({inv}): no gallery images found")
                            log_rows.append({"Inventory Number": inv, "MPN": mpn, "Product Title": title,
                                             "Status": "No images found"})
                        else:
                            if expected is not None and expected != count:
                                logging.warning(f"MPN {mpn} ({inv}): gallery says {expected} images, found {count}")
                            print(f"   -> {title} | {count} images (gallery says {expected})")

                            for idx, img_url in enumerate(urls, start=1):
                                folder, path = image_paths(title, inv, idx)
                                row = {"Inventory Number": inv, "MPN": mpn, "Product Title": title,
                                       "Image #": idx, "File Name": path.name, "Source URL": img_url}
                                try:
                                    if SKIP_EXISTING and path.exists():
                                        with Image.open(path) as existing:
                                            row.update(Width=existing.width, Height=existing.height,
                                                       Status="Skipped (already exists)")
                                    else:
                                        folder.mkdir(parents=True, exist_ok=True)
                                        img, note = process_image(fetch_bytes(driver, img_url))
                                        save_image(img, path)
                                        row.update(Width=img.width, Height=img.height, Status="OK", Note=note)
                                        if img.width != TARGET_WIDTH:
                                            row["Status"] = f"CHECK - width is {img.width}, not {TARGET_WIDTH}"
                                    print(f"      {path.name}  {row.get('Width')}x{row.get('Height')}  {row['Status']}")
                                except Exception as e:
                                    logging.error(f"MPN {mpn} image {idx} failed: {e}")
                                    row["Status"] = f"ERROR: {e}"
                                log_rows.append(row)
            except Exception as e:
                logging.error(f"MPN {mpn} failed: {e}")
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
    # ---- TEST MODE: comment the line below and uncomment main() to run the full list ----
    # main(test_mpns=["85360", "101450", "80156"])
    main()