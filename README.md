# WARN.com Product Image Scraper

A Python web scraper that downloads product images from [warn.com](https://www.warn.com) and exports a log to Excel.

---

## What It Does

For every SKU in the input Excel file, the scraper:

1. Finds the product on warn.com using a 3-step lookup:
   - Sitemap (`sitemap.xml`) — exact MPN token match in the product URL handle
   - Search page — waits for a result whose URL contains the exact MPN
   - Predictive search (`/search/suggest.json`) — exact MPN token match only
   - Falls back to the first search result as a last-resort candidate (validated by MPN presence on the page)
2. Validates the opened page contains the MPN before downloading — wrong products are skipped and flagged
3. Collects every image from the product's main gallery slider, in gallery order
4. Downloads each image at **1800px wide**, converts to JPG
5. Saves all images in **one flat folder** on the Desktop — no subfolders per part
   - File naming uses a double underscore separator: `WAR-17801__1.jpg`, `WAR-17801__2.jpg`, etc.
6. Writes a log workbook to the Desktop:
   - **Sheet: Image Log** — one row per image (dimensions, status, source URL)
   - **Sheet: Parts To Review** — SKUs with 0–1 images, not found, wrong product, or errors

Progress is auto-saved every 5 SKUs to prevent data loss on crashes.

---

## Requirements

- Python 3.8+
- Google Chrome installed

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Setup

1. Place your SKU list Excel file in the project folder
2. The file must have an **Inventory Number** column (e.g. `WAR-85360`)
   - If absent, the number after the last hyphen is used as the MPN
3. Update `INPUT_FILE` in the script if your filename differs

---

## Usage

```bash
python warn_image_scraper.py
```

- A browser window will open — solve any CAPTCHA manually, then press **Enter** to continue
- Images are saved to your Desktop in `Warn_Images_YYYY-MM-DD/`
- Log file is saved to your Desktop as `Warn_Image_Log_YYYY-MM-DD.xlsx`

### Test Mode

Edit the bottom of the script:

```python
if __name__ == "__main__":
    main(test_mpns=["85360", "101450", "80156"])
```

---

## Configuration

Key settings at the top of the script:

| Constant | Default | Description |
|---|---|---|
| `INPUT_FILE` | `Warn_Scrape_List_2026-11-16.xlsx` | Path to your SKU list |
| `TARGET_WIDTH` | `1800` | Output image width in pixels |
| `UPSCALE_SMALLER_IMAGES` | `True` | Enlarge images narrower than TARGET_WIDTH |
| `OUTPUT_FORMAT` | `jpg` | Output format: `jpg` or `png` |
| `JPEG_QUALITY` | `95` | JPEG compression quality |
| `NAME_SEPARATOR` | `__` | Separator between inventory number and image index (double underscore) |
| `SKIP_EXISTING` | `True` | Skip images already downloaded |
| `REVIEW_MAX_IMAGES` | `1` | Flag SKUs with this many images or fewer for review |
| `SAVE_LOG_EVERY` | `5` | Auto-save log interval (every N SKUs) |

---

## Output

| Location | Contents |
|---|---|
| `~/Desktop/Warn_Images_YYYY-MM-DD/` | All images in one flat folder (`WAR-XXXXX__1.jpg`, `WAR-XXXXX__2.jpg`, ...) |
| `~/Desktop/Warn_Image_Log_YYYY-MM-DD.xlsx` | Image Log + Parts To Review sheets |

### Log Sheet Columns

| Sheet | Columns |
|---|---|
| Image Log | Inventory Number, MPN, Product Title, Image #, File Name, Width, Height, Status, Note, Source URL |
| Parts To Review | Inventory Number, MPN, Product Title, Image Count, Review Status, Product URL |

---

## Notes

- Logs are written to `warn_image_scraper.log` in the project folder
- Uses [SeleniumBase](https://github.com/seleniumbase/SeleniumBase) with undetected-chromedriver mode (`uc=True`) to handle bot protection
- Images are downloaded through the browser itself (same session cookies) to avoid bot-blocking; falls back to `requests` if the browser fetch fails
- If a wrong product is opened, the Product Title is left blank in both log sheets — the incorrect title is only retained in `warn_image_scraper.log` for debugging
- Transparency (RGBA/PNG) is flattened onto a white background before saving as JPG
