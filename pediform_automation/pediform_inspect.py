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

async def inspect_page(page):
    print("\nNavigating to Today's Patients...")
    await page.get_by_role("link", name="Today's Patients", exact=True).click()
    await asyncio.sleep(3)

    await page.screenshot(path="todays_patients.png")
    print("Screenshot saved: todays_patients.png")

    print("\n--- INPUT ELEMENTS ---")
    inputs = await page.locator("input").all()
    for el in inputs:
        id_   = await el.get_attribute("id") or ""
        name  = await el.get_attribute("name") or ""
        type_ = await el.get_attribute("type") or ""
        cls   = await el.get_attribute("class") or ""
        print(f"  <input> id='{id_}'  name='{name}'  type='{type_}'  class='{cls[:60]}'")

    print("\n--- BUTTONS ---")
    buttons = await page.locator("button").all()
    for btn in buttons:
        id_  = await btn.get_attribute("id") or ""
        cls  = await btn.get_attribute("class") or ""
        text = (await btn.inner_text()).strip().replace("\n", " ")[:60]
        print(f"  <button> id='{id_}'  class='{cls[:50]}'  text='{text}'")

    print("\n--- LINKS ---")
    links = await page.locator("a").all()
    for a in links:
        href = await a.get_attribute("href") or ""
        text = (await a.inner_text()).strip().replace("\n", " ")[:60]
        print(f"  <a> href='{href}'  text='{text}'")

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
        await inspect_page(page)
        print("\nDone! Check the output above and the screenshot: todays_patients.png")
        await asyncio.sleep(99999)

asyncio.run(main())
