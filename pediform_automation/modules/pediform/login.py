"""
modules/pediform/login.py
-------------------------
Login to PediForm Pro and navigate to Today's Patients.
"""

import asyncio
from playwright.async_api import Page, BrowserContext

from config.settings import PEDIFORM_URL, PEDIFORM_ORG, PEDIFORM_EMAIL, PEDIFORM_PASSWORD
from modules.pediform.selectors import (
    LOGIN_ORG, LOGIN_EMAIL, LOGIN_PASSWORD, LOGIN_SUBMIT, NAV_TODAYS_PATIENTS
)
from utils.logger import get_logger

logger = get_logger(__name__)


async def login(page: Page) -> None:
    """Log into PediForm. Raises on timeout or bad credentials."""
    logger.info("Opening PediForm login page...")
    await page.goto(PEDIFORM_URL, wait_until="commit", timeout=90000)

    logger.info("Waiting for login form...")
    await page.locator(LOGIN_ORG).wait_for(state="visible", timeout=90000)

    await page.locator(LOGIN_ORG).fill(PEDIFORM_ORG)
    await page.locator(LOGIN_EMAIL).fill(PEDIFORM_EMAIL)
    await page.locator(LOGIN_PASSWORD).fill(PEDIFORM_PASSWORD)
    await page.locator(LOGIN_SUBMIT).click()

    await page.get_by_role("link", name=NAV_TODAYS_PATIENTS).wait_for(
        state="visible", timeout=90000
    )
    logger.info("PediForm login successful.")


async def navigate_to_todays_patients(page: Page) -> None:
    """Click Today's Patients nav link and wait for the table."""
    await page.get_by_role("link", name=NAV_TODAYS_PATIENTS).click()
    await asyncio.sleep(2)
    logger.debug("Navigated to Today's Patients.")


async def new_pediform_context(playwright):
    """Launch a Playwright browser configured to avoid bot detection."""
    from config.settings import HEADLESS, SLOW_MO
    browser = await playwright.chromium.launch(
        headless=HEADLESS,
        slow_mo=SLOW_MO,
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
    return browser, context
