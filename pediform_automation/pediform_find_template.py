import asyncio
from playwright.async_api import async_playwright
import config.settings as cfg

async def login(page):
    await page.goto(cfg.PEDIFORM_URL, wait_until="commit", timeout=90000)
    await page.locator("#admin-practice").wait_for(state="visible", timeout=90000)
    await page.locator("#admin-practice").fill(cfg.PEDIFORM_ORG)
    await page.locator("#admin-email").fill(cfg.PEDIFORM_EMAIL)
    await page.locator("#admin-password").fill(cfg.PEDIFORM_PASSWORD)
    await page.locator("button.patient-portal-submit").click()
    await page.get_by_role("link", name="Today's Patients", exact=True).wait_for(state="visible", timeout=90000)
    print("Logged in!")

async def find_template(page):
    print("Going to Today's Patients...")
    await page.get_by_role("link", name="Today's Patients", exact=True).click()
    await asyncio.sleep(3)

    page_text = await page.inner_text("body")
    print("\n--- RELEVANT PAGE TEXT (template/download/sample/example) ---")
    for line in page_text.splitlines():
        line = line.strip()
        if line and any(word in line.lower() for word in ["template", "download", "sample", "example", "csv", "excel", "xlsx", "format"]):
            print(f"  {line}")

    print("\n--- ALL LINKS ---")
    links = await page.locator("a").all()
    for a in links:
        href = await a.get_attribute("href") or ""
        text = (await a.inner_text()).strip().replace("\n", " ")[:80]
        if text or href:
            print(f"  href='{href}'  text='{text}'")

    print("\n--- DOWNLOAD LINKS ---")
    dl_links = await page.locator("[download]").all()
    for el in dl_links:
        href = await el.get_attribute("href") or ""
        text = (await el.inner_text()).strip()[:80]
        print(f"  href='{href}'  text='{text}'")

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=cfg.HEADLESS,
            slow_mo=cfg.SLOW_MO,
            args=["--disable-blink-features=AutomationControlled"],
        )
        context = await browser.new_context(
            ignore_https_errors=True,
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
        )
        page = await context.new_page()
        await login(page)
        await find_template(page)
        print("\nDone! Browser staying open.")
        await asyncio.sleep(99999)

asyncio.run(main())
