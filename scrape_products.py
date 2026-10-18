import logging
import os
import re
import time
from urllib.parse import quote, urljoin
import openpyxl
from playwright.sync_api import Page, sync_playwright

# --- SETTINGS ---
START_ROW = 2  # Starting row in Excel
HEADED = True  # Set to False to run in headless mode
EXCEL_FILE = "data/ERC_Stock.xlsx"
STATE_FILE = "data/cf_state.json"
LOG_FILE = "data/scraper.log"
BASE_URL = "https://erc-distribution.ge"
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


def pass_cloudflare_challenge(page: Page, timeout: int = 30) -> bool:
    """Detect and solve Cloudflare Turnstile challenge if present."""
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            title = page.title()
        except Exception:
            page.wait_for_timeout(500)
            continue

        if title and "Just a moment" not in title:
            return True

        # Look for the Cloudflare Turnstile iframe and click its checkbox
        for frame in page.frames:
            if "challenges.cloudflare.com" in frame.url:
                try:
                    box = frame.frame_element().bounding_box()
                    if box:
                        log.info("Solving Cloudflare Turnstile challenge...")
                        # Checkbox is located ~28px from the left edge, vertically centered
                        page.mouse.click(
                            box["x"] + 28,
                            box["y"] + box["height"] / 2,
                        )
                        page.wait_for_timeout(3000)
                        break
                except Exception:
                    pass

        page.wait_for_timeout(1000)

    try:
        return "Just a moment" not in page.title()
    except Exception:
        return False


def wait_for_georgian(page: Page, timeout: float = 8.0) -> None:
    """Wait until GTranslate finishes translating product blocks into Georgian."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            translated = page.evaluate(
                """() => {
                    const el = document.querySelector('.product-char td, section.product-description .pub-content');
                    return !!(el && /[\\u10d0-\\u10ff]/.test(el.innerText));
                }"""
            )
            if translated:
                page.wait_for_timeout(500)
                return
        except Exception:
            pass
        page.wait_for_timeout(300)


def find_desc_cols(ws) -> tuple[int, int]:
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col).value in ("desc", "desc_1"):
            ws.cell(row=1, column=col, value="desc_1")
            ws.cell(row=1, column=col + 1, value="desc_2")
            return col, col + 1
    col = ws.max_column + 1
    ws.cell(row=1, column=col, value="desc_1")
    ws.cell(row=1, column=col + 1, value="desc_2")
    return col, col + 1


def main():
    log.info("Loading workbook %s", EXCEL_FILE)
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["Stock"]
    desc1_col, desc2_col = find_desc_cols(ws)
    log.info(
        "Starting scrape from row %d to %d (desc_1 col=%d, desc_2 col=%d)",
        START_ROW,
        ws.max_row,
        desc1_col,
        desc2_col,
    )

    with sync_playwright() as p:
        launch_args = ["--disable-blink-features=AutomationControlled"]
        if not HEADED:
            # Use Chromium's new headless mode instead of legacy headless to avoid detection
            launch_args.append("--headless=new")

        browser = p.chromium.launch(
            headless=False,
            args=launch_args,
        )

        # Match the User-Agent's Chrome major version and OS to the actual browser binary
        major_version = browser.version.split(".")[0]
        user_agent = (
            f"Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
            f"(KHTML, like Gecko) Chrome/{major_version}.0.0.0 Safari/537.36"
        )

        context_kwargs = {
            "viewport": {"width": 1920, "height": 1080},
            "user_agent": user_agent,
        }
        if os.path.exists(STATE_FILE):
            context_kwargs["storage_state"] = STATE_FILE

        context = browser.new_context(**context_kwargs)
        page = context.new_page()

        try:
            for row in range(START_ROW, ws.max_row + 1):
                raw_model = ws.cell(row=row, column=5).value
                if not raw_model or not str(raw_model).strip():
                    continue
                model = str(raw_model).strip()

                try:
                    search_url = f"{BASE_URL}/search/?q={quote(model)}&_locale=ge"
                    page.goto(search_url, wait_until="domcontentloaded", timeout=45000)
                    if not pass_cloudflare_challenge(page):
                        log.warning("[Row %d] Cloudflare challenge failed on search for %s", row, model)
                        continue
                    os.makedirs(os.path.dirname(STATE_FILE), exist_ok=True)
                    context.storage_state(path=STATE_FILE)

                    if "/ware/" not in page.url:
                        link_loc = page.locator(".product-item__title a").first
                        if link_loc.count() == 0:
                            log.warning("[Row %d] Not found: %s", row, model)
                            continue
                        product_url = urljoin(BASE_URL, link_loc.get_attribute("href"))
                        sep = "&" if "?" in product_url else "?"
                        page.goto(
                            f"{product_url}{sep}_locale=ge",
                            wait_until="domcontentloaded",
                            timeout=45000,
                        )
                        if not pass_cloudflare_challenge(page):
                            log.warning("[Row %d] Cloudflare challenge failed on product %s", row, model)
                            continue

                    wait_for_georgian(page)

                    char_loc = page.locator(".product-char").first
                    desc_1 = char_loc.inner_text().strip() if char_loc.count() > 0 else ""

                    desc_2_parts = [
                        el.inner_text().strip()
                        for el in page.locator("section.product-description").all()
                        if el.inner_text().strip()
                    ]
                    desc_2 = "\n\n".join(desc_2_parts)

                    ws.cell(row=row, column=desc1_col, value=desc_1)
                    ws.cell(row=row, column=desc2_col, value=desc_2)

                    img_urls = []
                    for a in page.locator(
                        '.product-slider__big-item:not(.slick-cloned) a[data-fancybox="gallery"]'
                    ).all():
                        href = a.get_attribute("href")
                        if href:
                            full_url = urljoin(BASE_URL, href)
                            if full_url not in img_urls:
                                img_urls.append(full_url)

                    saved_imgs = 0
                    if img_urls:
                        safe_model = re.sub(r'[\\/*?:"<>|]', "_", model)
                        model_dir = os.path.join("data", safe_model)
                        os.makedirs(model_dir, exist_ok=True)
                        for idx, img_url in enumerate(img_urls, start=1):
                            resp = context.request.get(img_url, timeout=30000)
                            if resp.ok:
                                with open(os.path.join(model_dir, f"{idx}.jpg"), "wb") as f:
                                    f.write(resp.body())
                                saved_imgs += 1
                            else:
                                log.warning("[Row %d] Failed to download image %s (HTTP %d)", row, img_url, resp.status)

                    log.info(
                        "[Row %d/%d] %s: desc_1=%d chars, desc_2=%d chars, images=%d/%d",
                        row,
                        ws.max_row,
                        model,
                        len(desc_1),
                        len(desc_2),
                        saved_imgs,
                        len(img_urls),
                    )

                    if row % 10 == 0:
                        wb.save(EXCEL_FILE)
                        log.info("[Row %d] Checkpoint saved to %s", row, EXCEL_FILE)

                except Exception:
                    log.exception("[Row %d] Error scraping %s", row, model)

        finally:
            wb.save(EXCEL_FILE)
            log.info("Workbook saved to %s", EXCEL_FILE)
            browser.close()


if __name__ == "__main__":
    main()

