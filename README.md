# WARN.com Product Image Scraper

A Python web scraper that downloads product images from [warn.com](https://www.warn.com) and exports a log to Excel.

## What It Does

For every SKU in the input Excel file, the scraper:

1. Finds the product on warn.com
2. Collects every image from the product gallery (main slider only, in gallery order)
3. Downloads each image at **1800px wide**, converted to JPG
4. Saves images in a subfolder per Part Number: `<Product Title>_<Part Number>.jpg` (extra images: `_2`, `_3`, ...)
5. Creates a folder for **every** part (even with no images)
6. Writes a log workbook to the Desktop:
   - **Sheet: Image Log** — one row per image
   - **Sheet: Parts To Review** — parts with 0–1 images, not found, or errors

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

1. Place your SKU list Excel file in the project folder (default: `Warn_Scrape_List_2026-10-27.xlsx`)
2. The file must have an **Inventory Number** column (e.g. `WAR-85360`)
   - Optional: `MPN` column (if absent, the number after the last hyphen is used)
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

To test with specific MPNs only, edit the bottom of the script:

```python
if __name__ == "__main__":
    main(test_mpns=["85360", "101450", "80156"])
```

---

## Configuration

Key settings at the top of the script:

| Constant | Default | Description |
|---|---|---|
| `INPUT_FILE` | `Warn_Scrape_List_2026-10-27.xlsx` | Path to your SKU list |
| `TARGET_WIDTH` | `1800` | Output image width in pixels |
| `UPSCALE_SMALLER_IMAGES` | `True` | Enlarge images smaller than 1800px |
| `OUTPUT_FORMAT` | `jpg` | Output format: `jpg` or `png` |
| `JPEG_QUALITY` | `95` | JPEG compression quality |
| `SUBFOLDER_PER_PART` | `True` | One subfolder per Part Number |
| `SKIP_EXISTING` | `True` | Skip already-downloaded images |
| `REVIEW_MAX_IMAGES` | `1` | Flag parts with this many images or fewer for review |
| `SAVE_LOG_EVERY` | `5` | Auto-save interval (every N SKUs) |

---

## Output

| Location | Contents |
|---|---|
| `~/Desktop/Warn_Images_YYYY-MM-DD/` | Downloaded images, one subfolder per SKU |
| `~/Desktop/Warn_Image_Log_YYYY-MM-DD.xlsx` | Image Log + Parts To Review sheets |

---

## Notes

- Logs are written to `warn_image_scraper.log` in the project folder
- The scraper uses [SeleniumBase](https://github.com/seleniumbase/SeleniumBase) with undetected-chromedriver mode (`uc=True`) to handle bot protection
- Images are downloaded through the browser itself (same cookies) to avoid bot-blocking
