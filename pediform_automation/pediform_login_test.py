import asyncio
from playwright.async_api import async_playwright
import config.settings as cfg

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

        print("Opening PediForm login page...")
        await page.goto(cfg.PEDIFORM_URL, wait_until="commit", timeout=90000)
        print("Server responded — waiting for login form to render...")

        await page.locator("#admin-practice").wait_for(state="visible", timeout=90000)
        print("Login form visible — filling credentials...")

        await page.locator("#admin-practice").fill(cfg.PEDIFORM_ORG)
        print("Organisation filled")

        await page.locator("#admin-email").fill(cfg.PEDIFORM_EMAIL)
        print("Email filled")

        await page.locator("#admin-password").fill(cfg.PEDIFORM_PASSWORD)
        print("Password filled")

        await page.locator("button.patient-portal-submit").click()
        print("Sign In clicked — waiting for dashboard...")

        await page.get_by_role("link", name="Today's Patients", exact=True).wait_for(
            state="visible", timeout=90000
        )
        print("SUCCESS — Logged in! Today's Patients page loaded.")

        await asyncio.sleep(99999)

asyncio.run(main())
