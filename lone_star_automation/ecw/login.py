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

    await _wait_for_loading_overlay_gone(page)

    await dismiss_license_alert(page)

    # (2026-07-23 fix) The "#load" / "Building your user experience" overlay
    # can re-render and reappear AFTER the first hidden-check resolves -
    # eCW keeps loading asynchronously post-login/post-license-alert. A
    # live production run showed the overlay still intercepting pointer
    # events on the very next click (#jellybean-panelLink4), ~25+ seconds
    # after the code had already logged "Loading screen already hidden!".
    # Re-verify right before handing control back to callers, since they
    # click navigation elements immediately.
    await _wait_for_loading_overlay_gone(page)


async def _wait_for_loading_overlay_gone(page, timeout_ms=60000, retries=3):
    """
    Robust wait for eCW's '#load' overlay to be hidden. A single
    wait_for_selector(state='hidden') call can resolve or except
    prematurely if the overlay toggles visibility multiple times during a
    complex page load - retries a few times rather than trusting one shot
    and silently assuming "already hidden" on any exception (that
    assumption was directly disproved live - the overlay was still
    blocking clicks well after the exception fired).
    """
    for attempt in range(retries):
        try:
            await page.wait_for_selector('#load', state='hidden', timeout=timeout_ms)
            log.info("Loading overlay confirmed hidden.")
            return
        except Exception:
            log.info(f"Loading overlay still present or check unstable (attempt {attempt + 1}/{retries}) - re-checking...")
            await asyncio.sleep(2)
    log.info("Proceeding despite loading-overlay uncertainty after retries.")


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
