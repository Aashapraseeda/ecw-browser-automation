"""
ecw/login.py
------------
Shared eCW login + License Alert dismissal.

Extracted from the reference project, where this exact block was
duplicated inside ecw_export_schedule() and ecw_upload_forms(). Behavior
is unchanged - just called from one place now.
"""

import asyncio

from config import settings
from utils.logger import get_logger

log = get_logger(__name__)


async def ecw_login(page):
    log.info("Logging into eCW...")
    await page.goto(
        settings.ECW_LOGIN_URL,
        timeout=120000, wait_until="domcontentloaded"
    )
    await asyncio.sleep(3)
    await page.get_by_role("textbox", name="Enter username to continue").fill(settings.ECW_USERNAME)
    await page.get_by_role("button", name="Next").click()
    await asyncio.sleep(8)
    await page.click('input[type="password"]')
    await page.keyboard.type(settings.ECW_PASSWORD)
    await page.keyboard.press("Enter")
    log.info("Login submitted...")

    await page.wait_for_selector('#jellybean-panelLink33', timeout=120000)
    log.info("Home page loaded!")

    log.info("Waiting for eCW to fully load...")
    try:
        await page.wait_for_selector('#load', state='hidden', timeout=120000)
        log.info("eCW fully loaded!")
    except Exception:
        log.info("Loading screen already hidden!")

    await dismiss_license_alert(page)


async def dismiss_license_alert(page):
    log.info("Checking for License Alert...")
    dismissed = False
    for _ in range(20):
        try:
            if await page.locator("#providerLicense button.clsMyButton").is_visible():
                await page.click("#providerLicense button.clsMyButton")
                log.info("License Alert dismissed!")
                dismissed = True
                break
        except Exception:
            pass
        await asyncio.sleep(1)
    if not dismissed:
        log.info("No License Alert, continuing...")
    await asyncio.sleep(2)
