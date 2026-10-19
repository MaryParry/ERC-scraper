import argparse
import logging
import os
import re
import time
from pathlib import Path
import openpyxl
from playwright.sync_api import Page, sync_playwright

# --- SETTINGS ---
START_ROW = 2
HEADED = True
EXCEL_FILE = "data/ERC_Stock.xlsx"
ORTE_STATE_FILE = "data/orte_state.json"
LOG_FILE = "data/uploader.log"
NEW_PRODUCT_URL = "https://orte.ge/seller/products/new"
# ----------------

os.makedirs("data", exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)

# Fixes for GTranslate artifacts in ERC scraped specs (desc_1)
KEY_FIXES = {
    "პარტიების მხარდაჭერა": "ასპექტის თანაფარდობა",
    "კაშკაშა": "სიკაშკაშე",
    "აპლიკაცია": "დანიშნულება",
    "შეყვანები": "პორტები",
    "LED განათება, LED განათება, უკანა განათება.": "განათების ტიპი",
}

VAL_FIXES = {
    "ვეტერანთა საქმეების დეპარტამენტი": "VA",
    "ჩვენების პორტი": "DisplayPort",
    "დისპლეიპორტი": "DisplayPort",
    "სწრაფი IPS": "Fast IPS",
    "შავი IPS": "IPS Black",
    "(нормальний режим)": "(სტანდარტული რეჟიმი)",
}

# Known typos in ERC_Stock.xlsx categories/subcategories
CATEGORY_TYPO_FIXES = {
    "კომპიუეტერი და აქსესუარები": "კომპიუტერი და აქსესუარები",
    "თნის მოვლა": "თმის მოვლა",
    "თავის მოვლა": "სილამაზე და ჯანმრთელობა",
}


def format_short_description(raw: str | None) -> str:
    """Clean and format badly scraped ERC specification table text (desc_1)."""
    if not raw or not str(raw).strip():
        return ""
    text = str(raw).replace("\xa0", " ").strip()
    lines = text.splitlines()
    items: list[list] = []  # [[key, [val1, val2, ...]], ...]
    free_lines: list[str] = []

    for line in lines:
        s = line.strip()
        if not s:
            continue
        # Strip ERC webpage UI headers and toggle button labels
        if (
            s.startswith("ტექნიკური პარამეტრები")
            or s in ("მთავარი მახასიათებლები", "მთავარი მახასიათებლების სრული სია")
        ):
            continue

        if "\t" in line:
            parts = line.split("\t", 1)
            k = parts[0].strip().rstrip(":")
            v = parts[1].strip()
            for old, new in VAL_FIXES.items():
                v = v.replace(old, new)

            if not k and v and items:
                # Continuation line belonging to previous spec key
                items[-1][1].append(v)
            elif k:
                k = KEY_FIXES.get(k, k)
                items.append([k, [v] if v else []])
        else:
            free_lines.append(s)

    out = []
    for k, vals in items:
        non_empty = [v for v in vals if v]
        if not non_empty:
            continue
        val_str = ", ".join(non_empty)
        val_str = re.sub(r'(\d)\s+"', r'\1"', val_str)
        out.append(f"• {k}: {val_str}")

    if out:
        return "\n".join(out)
    return "\n".join(free_lines)


def format_detailed_description(raw: str | None) -> str:
    """Clean detailed description (desc_2)."""
    if not raw or not str(raw).strip():
        return ""
    text = str(raw).replace("\xa0", " ").strip()
    lines = text.splitlines()
    # Remove redundant first line if it starts with 'აღწერა ' or 'მონიტორის აღწერა '
    if lines and (
        lines[0].strip().startswith("აღწერა ")
        or lines[0].strip().startswith("მონიტორის აღწერა ")
    ):
        lines = lines[1:]
    return "\n".join(lines).strip()


def find_model_images(model: str) -> list[str]:
    """Return sorted list of image file paths in data/<model>/."""
    safe_model = re.sub(r'[\\/*?:"<>|]', "_", model.strip())
    model_dir = Path("data") / safe_model
    if not model_dir.is_dir():
        return []

    valid_exts = {".jpg", ".jpeg", ".png", ".webp"}
    files = [
        p for p in model_dir.iterdir()
        if p.is_file() and p.suffix.lower() in valid_exts
    ]

    def sort_key(p: Path):
        nums = re.findall(r"\d+", p.stem)
        return (int(nums[0]) if nums else 999999, p.name)

    files.sort(key=sort_key)
    return [str(p.resolve()) for p in files]


def parse_price(promo_val, rrp_val) -> str | None:
    """Return Promo RRP GEL (H) if present and valid, else RRP GEL (G)."""
    for val in (promo_val, rrp_val):
        if val is None:
            continue
        s = str(val).strip().replace(",", ".")
        if not s:
            continue
        try:
            num = float(s)
            if num > 0:
                return f"{num:g}" if num.is_integer() else f"{num:.2f}"
        except ValueError:
            continue
    return None


def select_category(page: Page, category: str | None, subcategory: str | None, prefer_sub: bool = False) -> str:
    """Open #cat dropdown and select the matching category option."""
    cat_clean = CATEGORY_TYPO_FIXES.get(category.strip(), category.strip()) if category and str(category).strip() else None
    sub_clean = CATEGORY_TYPO_FIXES.get(subcategory.strip(), subcategory.strip()) if subcategory and str(subcategory).strip() else None

    target_primary = sub_clean if (prefer_sub and sub_clean) else (cat_clean or sub_clean)
    target_fallback = cat_clean if target_primary == sub_clean else sub_clean

    if not target_primary:
        raise ValueError("Both Category (C) and Sub-category (D) are empty")

    cat_btn = page.locator("#cat")
    if cat_btn.get_attribute("aria-expanded") != "true":
        cat_btn.click()

    page.wait_for_selector('[role="listbox"] [role="option"]', timeout=5000)
    options = page.locator('[role="listbox"] [role="option"]')
    count = options.count()

    def try_click(target: str) -> bool:
        for i in range(count):
            opt = options.nth(i)
            raw_txt = opt.inner_text().strip()
            # Parent options look like: "ელექტრო ტექნიკა\n16 ქვე" or "ელექტრო ტექნიკა16 ქვე"
            # Child options look like: "└\nკომპიუტერი და აქსესუარები" or "└კომპიუტერი და აქსესუარები"
            cleaned_txt = re.sub(r"\d+\s*ქვე$", "", raw_txt).strip()
            cleaned_txt = cleaned_txt.lstrip("└").strip()
            if cleaned_txt == target:
                opt.scroll_into_view_if_needed()
                opt.click()
                return True
        return False

    if try_click(target_primary):
        return target_primary
    if target_fallback and try_click(target_fallback):
        log.warning("Primary category %r not found, used fallback %r", target_primary, target_fallback)
        return target_fallback

    raise RuntimeError(f"Could not find category {target_primary!r} (or fallback {target_fallback!r}) in #cat listbox")


def fill_and_upload_product(
    page: Page,
    *,
    name: str,
    model: str,
    brand: str,
    category: str | None,
    subcategory: str | None,
    short_desc: str,
    detailed_desc: str,
    price: str,
    image_paths: list[str],
    prefer_subcategory: bool = False,
    dry_run: bool = False,
    screenshot_path: str | None = None,
) -> bool:
    """Navigate to /seller/products/new, fill all fields, upload images, and optionally submit."""
    page.goto(NEW_PRODUCT_URL, wait_until="domcontentloaded", timeout=30000)
    try:
        page.wait_for_selector("#name", timeout=10000)
    except Exception:
        log.error("Failed to reach #name on %s (current URL: %s). Session may be unauthenticated.", NEW_PRODUCT_URL, page.url)
        return False

    # 1. Images
    if image_paths:
        file_input = page.locator('input[type="file"]')
        file_input.set_input_files(image_paths)
        page.wait_for_timeout(500)
        log.info("Attached %d image(s): %s", len(image_paths), [os.path.basename(p) for p in image_paths])
    else:
        log.warning("No images found for model %s", model)

    # 2. Product Name = Name (F) + " " + Model (E)
    clean_name = re.sub(r"\s+", " ", name.strip())
    clean_model = model.strip()
    full_name = f"{clean_name} {clean_model}" if clean_model not in clean_name else f"{clean_name} {clean_model}"
    page.locator("#name").fill(full_name)

    # 3. Brand (B)
    page.locator("#brand").fill(brand.strip() if brand else "")

    # 4. Short Description (#short) <- formatted desc_1
    page.locator("#short").fill(short_desc)

    # 5. Detailed Description (#desc) <- formatted desc_2
    page.locator("#desc").fill(detailed_desc)

    # 6. Category (#cat) <- Category column (C)
    selected_cat = select_category(page, category, subcategory, prefer_sub=prefer_subcategory)
    log.info("Selected category: %s", selected_cat)

    # 7. Price (#price) <- Promo RRP GEL (H) else RRP GEL (G)
    page.locator("#price").fill(price)

    # 8. Stock (#stock) <- always 10
    page.locator("#stock").fill("10")

    # 9. Weight (#weight) <- optional, leave empty
    page.locator("#weight").fill("")

    # 10. Delivery days (#delivery-days) <- always 3
    page.locator("#delivery-days").fill("3")

    # 11. Status (#status) <- always active
    status_sel = page.locator("#status")
    if status_sel.count() > 0 and status_sel.input_value() != "active":
        status_sel.select_option("active")

    page.wait_for_timeout(500)

    if screenshot_path:
        page.screenshot(path=screenshot_path, full_page=True)
        log.info("Saved form preview screenshot to %s", screenshot_path)

    if dry_run:
        log.info("[DRY RUN] Form filled for %s (%s). Skipping submit.", model, full_name)
        return True

    # Submit form
    submit_btn = page.locator('button[type="submit"]')
    submit_btn.click()

    # Wait for navigation away from /seller/products/new or success toast
    deadline = time.time() + 30.0
    while time.time() < deadline:
        if "/seller/products/new" not in page.url:
            log.info("Successfully uploaded product %s -> redirected to %s", model, page.url)
            return True
        page.wait_for_timeout(500)

    log.error("Timed out waiting for product submission to complete for %s (still on %s)", model, page.url)
    return False


def find_or_create_col(ws, header_name: str) -> int:
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col).value == header_name:
            return col
    col = ws.max_column + 1
    ws.cell(row=1, column=col, value=header_name)
    return col


def main():
    parser = argparse.ArgumentParser(description="Upload ERC products to orte.ge")
    parser.add_argument("--row", type=int, help="Upload only a specific Excel row number (e.g. 2)")
    parser.add_argument("--model", type=str, help="Upload only a specific Model (e.g. 2E-R2723BV-01.UA)")
    parser.add_argument("--dry-run", action="store_true", help="Fill form and take screenshot without submitting")
    parser.add_argument("--batch", action="store_true", help="Run batch upload across all ready rows")
    parser.add_argument("--headless", action="store_true", help="Run browser in headless mode")
    parser.add_argument("--use-subcategory", action="store_true", help="Select Sub-category (Col D) instead of Category (Col C)")
    parser.add_argument("--screenshot", type=str, default="data/form_preview.png", help="Screenshot path for dry-run/test")
    args = parser.parse_args()

    if not args.row and not args.model and not args.batch:
        parser.error("Please specify --row <N>, --model <MODEL>, or --batch (add --dry-run to test without submitting).")

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["Stock"]
    desc1_col = find_or_create_col(ws, "desc_1")
    desc2_col = find_or_create_col(ws, "desc_2")
    uploaded_col = find_or_create_col(ws, "orte_uploaded")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless)
        context_kwargs = {"viewport": {"width": 1440, "height": 1080}}
        if os.path.exists(ORTE_STATE_FILE):
            context_kwargs["storage_state"] = ORTE_STATE_FILE

        context = browser.new_context(**context_kwargs)
        page = context.new_page()

        try:
            for row in range(START_ROW, ws.max_row + 1):
                raw_model = ws.cell(row=row, column=5).value
                if not raw_model or not str(raw_model).strip():
                    continue
                model = str(raw_model).strip()

                if args.row and row != args.row:
                    continue
                if args.model and model != args.model.strip():
                    continue

                if args.batch and ws.cell(row=row, column=uploaded_col).value:
                    continue

                brand = str(ws.cell(row=row, column=2).value or "").strip()
                category = ws.cell(row=row, column=3).value
                subcategory = ws.cell(row=row, column=4).value
                name = str(ws.cell(row=row, column=6).value or "").strip()
                rrp_gel = ws.cell(row=row, column=7).value
                promo_gel = ws.cell(row=row, column=8).value
                raw_desc1 = ws.cell(row=row, column=desc1_col).value
                raw_desc2 = ws.cell(row=row, column=desc2_col).value

                price = parse_price(promo_gel, rrp_gel)
                if not price:
                    log.warning("[Row %d] Skipping %s: no valid price (G=%r, H=%r)", row, model, rrp_gel, promo_gel)
                    continue

                short_desc = format_short_description(raw_desc1)
                detailed_desc = format_detailed_description(raw_desc2)
                image_paths = find_model_images(model)

                if args.batch and not (short_desc or detailed_desc):
                    log.info("[Row %d] Skipping %s: descriptions not scraped yet", row, model)
                    continue

                log.info(
                    "[Row %d] Preparing %s | Price=%s GEL | Images=%d | short_desc=%d chars | detailed_desc=%d chars",
                    row,
                    model,
                    price,
                    len(image_paths),
                    len(short_desc),
                    len(detailed_desc),
                )

                ok = fill_and_upload_product(
                    page,
                    name=name,
                    model=model,
                    brand=brand,
                    category=category,
                    subcategory=subcategory,
                    short_desc=short_desc,
                    detailed_desc=detailed_desc,
                    price=price,
                    image_paths=image_paths,
                    prefer_subcategory=args.use_subcategory,
                    dry_run=args.dry_run,
                    screenshot_path=args.screenshot if (args.dry_run or args.row or args.model) else None,
                )

                if ok and not args.dry_run:
                    ws.cell(row=row, column=uploaded_col, value="yes")
                    wb.save(EXCEL_FILE)
                    context.storage_state(path=ORTE_STATE_FILE)

                if args.row or args.model:
                    break
        finally:
            browser.close()


if __name__ == "__main__":
    main()
