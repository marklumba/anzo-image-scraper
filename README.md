# ANZO Product Image Scraper

A Python web scraper that downloads product images from [anzousa.com](https://anzousa.com) and exports a log to Excel.

## What It Does

For every SKU in the input Excel file, the scraper:

1. Finds the product on anzousa.com (via the site's product list, then search)
2. Collects every image from the product's main slideshow (in gallery order)
3. Downloads each image at **1800px wide**, converted to JPG
4. Saves all images in **one flat folder** — no subfolders — named with a double underscore separator:
   - `ANZ-311302__1.jpg`, `ANZ-311302__2.jpg`, `ANZ-311354__1.jpg`, ...
5. Writes a log workbook to the Desktop:
   - **Sheet: Image Log** — one row per image
   - **Sheet: Parts To Review** — SKUs with 0–1 images, not found, or errors

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

1. Place your SKU list Excel file in the project folder (default: `ANZO_Scrape_List_2026-11-04.xlsx`)
2. The file must have an **Inventory Number** column (e.g. `ANZ-311302`)
   - Optional: `MPN` column (if absent, the number after the last hyphen is used)
3. Update `INPUT_FILE` in the script if your filename differs

---

## Usage

```bash
python anzo_image_scraper.py
```

- A browser window will open — solve any CAPTCHA manually, then press **Enter** to continue
- Images are saved to your Desktop in `Anzo_Images_YYYY-MM-DD/`
- Log file is saved to your Desktop as `Anzo_Image_Log_YYYY-MM-DD.xlsx`

### Test Mode

To test with specific MPNs only, edit the bottom of the script:

```python
if __name__ == "__main__":
    main(test_mpns=["311302", "311354"])
```

---

## Configuration

Key settings at the top of the script:

| Constant | Default | Description |
|---|---|---|
| `INPUT_FILE` | `ANZO_Scrape_List_2026-11-04.xlsx` | Path to your SKU list |
| `TARGET_WIDTH` | `1800` | Output image width in pixels |
| `UPSCALE_SMALLER_IMAGES` | `True` | Enlarge images smaller than 1800px |
| `OUTPUT_FORMAT` | `jpg` | Output format: `jpg` or `png` |
| `JPEG_QUALITY` | `95` | JPEG compression quality |
| `SKIP_EXISTING` | `True` | Skip already-downloaded images |
| `REVIEW_MAX_IMAGES` | `1` | Flag SKUs with this many images or fewer for review |
| `USE_PRODUCT_LIST` | `True` | Read the full product list first (faster) |
| `SKIP_SEARCH_IF_NOT_IN_LIST` | `False` | Skip search if SKU not in product list |
| `SAVE_LOG_EVERY` | `5` | Auto-save interval (every N SKUs) |
| `NAME_SEPARATOR` | `__` | Separator between SKU and image number in filename |

---

## Output

| Location | Contents |
|---|---|
| `~/Desktop/Anzo_Images_YYYY-MM-DD/` | All downloaded images in one flat folder |
| `~/Desktop/Anzo_Image_Log_YYYY-MM-DD.xlsx` | Image Log + Parts To Review sheets |

---

## Notes

- Logs are written to `anzo_image_scraper.log` in the project folder
- The scraper uses [SeleniumBase](https://github.com/seleniumbase/SeleniumBase) with undetected-chromedriver mode (`uc=True`) to handle bot protection
- Images are downloaded through the browser itself (same cookies) to avoid bot-blocking
