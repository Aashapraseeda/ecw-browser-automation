"""
ecw/schedule_export.py
-----------------------
STEP 0 - eCW: export the weekly schedule (Encounter Patient Download).

Ported from the reference project's ecw_export_schedule(). The only
workflow addition is the Facility filter step (facility_filter.apply),
inserted AFTER the date range is set (using the unchanged reference date
logic) and BEFORE the single OK/Finish click that submits both dates and
facility together - see ecw/facility_filter.py.
"""

import asyncio
from datetime import date, timedelta

from playwright.async_api import async_playwright

from config import settings
from ecw.login import ecw_login
from ecw.facility_filter import apply_facility_filter
from utils.logger import get_logger

log = get_logger(__name__)


async def click_calendar_option(iframe, day_str, descriptions_to_try, label):
    """Try clicking calendar option with multiple description variants then fallback."""
    for desc in descriptions_to_try:
        try:
            await iframe.get_by_role("option", name=day_str, description=desc, exact=True).click(timeout=5000)
            log.info(f"{label} date set! (description='{desc}')")
            return
        except Exception:
            pass
    # Final fallback - no description
    try:
        await iframe.get_by_role("option", name=day_str, exact=True).first.click(timeout=5000)
        log.info(f"{label} date set! (no description fallback)")
        return
    except Exception as e:
        raise RuntimeError(f"Could not set {label} date to day {day_str}: {e}")


async def run(window_days=7):
    """
    window_days: size of the export window (today -> today+window_days).
    Production uses 7 (default); the demo pipeline uses a shorter 3-day
    window (window_days=2, i.e. today/+1/+2), matching the reference
    project's main_1.py demo variant.
    """
    log.info("=" * 50)
    log.info(f"STEP 0 - ECW: EXPORTING SCHEDULE (TODAY + {window_days} DAYS)")
    log.info("=" * 50)

    # Both 2-letter and 3-letter day description variants
    day_desc_2 = {0: "Mo", 1: "Tu", 2: "We", 3: "Th", 4: "Fr", 5: "Sa", 6: "Su"}
    day_desc_3 = {0: "Mon", 1: "Tue", 2: "Wed", 3: "Thu", 4: "Fri", 5: "Sat", 6: "Sun"}

    start_date = date.today()
    end_date = start_date + timedelta(days=window_days)
    start_day_str = str(start_date.day)
    end_day_str = str(end_date.day)
    start_descs = [day_desc_2[start_date.weekday()], day_desc_3[start_date.weekday()]]
    end_descs = [day_desc_2[end_date.weekday()], day_desc_3[end_date.weekday()]]
    log.info(f"Date range: {start_date} to {end_date}")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        await ecw_login(page)

        # --- NAVIGATE TO EBO REPORTS ---
        log.info("Navigating to eBO Reports...")
        await page.locator("#jellybean-panelLink4").click()
        await asyncio.sleep(1)
        await page.get_by_text("Menu", exact=True).click()
        await asyncio.sleep(1)
        await page.locator("#pane6").get_by_text("Reports").click()
        await asyncio.sleep(1)

        async with page.expect_popup(timeout=60000) as page1_info:
            await page.get_by_text("eBO Reports CTRL + ALT + E").click()
        page1 = await page1_info.value
        await page1.goto(
            settings.ECW_EBO_HOME_URL,
            timeout=60000, wait_until="domcontentloaded"
        )
        await asyncio.sleep(5)
        log.info("eBO Reports opened!")

        # --- NAVIGATE TO ENCOUNTER PATIENT DOWNLOAD ---
        await page1.get_by_role("link", name="eCWEBO", exact=True).click()
        await asyncio.sleep(2)
        await page1.get_by_role("link", name="- Administrative Reports").click()
        await asyncio.sleep(2)
        await page1.get_by_role("link", name="- Encounter Patient Download").click()
        await asyncio.sleep(8)
        log.info("Encounter Patient Download opened!")

        # --- WAIT IF REPORT ALREADY RUNNING ---
        try:
            if await page1.get_by_text("Your report is running").is_visible():
                log.info("Report already running - waiting...")
                for i in range(90):
                    if not await page1.get_by_text("Your report is running").is_visible():
                        log.info("Report finished.")
                        break
                    log.info(f"Still running... ({i+1}/90)")
                    await asyncio.sleep(2)
        except Exception:
            pass

        # --- WAIT FOR IFRAME ---
        log.info("Waiting for iframe to load...")
        iframe = page1.locator("iframe[name=\"iD6D96C5E47F347C9B95828AC68A2D69B\"]").content_frame
        await iframe.get_by_role("img").first.wait_for(timeout=60000)
        await asyncio.sleep(3)
        log.info("Iframe loaded!")

        # --- SET DATES (unchanged from reference project) ---
        log.info(f"Setting start date: {start_date}")
        await iframe.get_by_role("img").first.click()
        await asyncio.sleep(2)
        await click_calendar_option(iframe, start_day_str, start_descs, "Start")
        await asyncio.sleep(1)

        log.info(f"Setting end date: {end_date}")
        await iframe.get_by_role("img").nth(1).click()
        await asyncio.sleep(2)
        await click_calendar_option(iframe, end_day_str, end_descs, "End")
        await asyncio.sleep(1)

        # --- NEW STEP: FACILITY FILTER (Lone Star specific) - after dates, before OK ---
        await apply_facility_filter(iframe)

        # --- CLICK OK (submits both dates and facility together) ---
        await iframe.get_by_role("button", name="OK").click()
        await asyncio.sleep(2)
        log.info(f"Date range confirmed: {start_date} to {end_date}")

        # --- WAIT FOR REPORT TO GENERATE ---
        log.info("Waiting for report to generate (2-3 minutes)...")
        for i in range(120):
            try:
                is_disabled = await page1.locator("button[aria-label='Select a format']").get_attribute("disabled")
                if is_disabled is None:
                    log.info("Report ready!")
                    break
                log.info(f"Report still running... ({i+1}/120)")
            except Exception:
                pass
            await asyncio.sleep(2)

        await asyncio.sleep(1)

        # --- DOWNLOAD EXCEL ---
        log.info("Clicking Select a format...")
        await page1.get_by_role("button", name="Select a format").click()
        await asyncio.sleep(3)
        await page1.get_by_role("link", name="Excel data").wait_for(timeout=15000)

        log.info("Clicking Excel data - waiting for download (may take 2-3 minutes)...")
        async with page1.expect_download(timeout=300000) as download_info:
            await page1.get_by_role("link", name="Excel data").click()
        download = await download_info.value
        await download.save_as(settings.EXCEL_PATH)
        log.info(f"Excel saved to: {settings.EXCEL_PATH}")

        await browser.close()
        log.info("Schedule exported successfully!")


if __name__ == "__main__":
    asyncio.run(run())
