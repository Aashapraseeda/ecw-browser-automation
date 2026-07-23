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


async def _wait_for_report_running_modal_gone(page1, max_checks=90):
    """
    Waits out eCW's "Your report is running" modal (a leftover/queued
    report execution, from this account's own prior report requests,
    still processing server-side). This modal sits on top of the whole
    report screen and blocks every click underneath it - including the
    date picker - which can look exactly like an unrelated UI bug if you
    only check for it once, early, and it appears again later.
    """
    try:
        if await page1.get_by_text("Your report is running").is_visible():
            log.info("Report already running - waiting...")
            for i in range(max_checks):
                if not await page1.get_by_text("Your report is running").is_visible():
                    log.info("Report finished.")
                    return
                log.info(f"Still running... ({i + 1}/{max_checks})")
                await asyncio.sleep(2)
            log.info("Report still running after max wait - proceeding anyway.")
    except Exception:
        pass


async def click_calendar_option(iframe, day_str, descriptions_to_try, label, retries=3):
    """
    Try clicking calendar option with multiple description variants then
    fallback.

    (2026-07-24) Live diagnostics confirmed the day option itself is
    genuinely visible, enabled, and has a real bounding box when this
    fails - the click is being blocked by something transiently
    overlapping it (likely a brief internal loading/refresh state of the
    date-picker widget), not a missing/broken element. This is the exact
    same pattern already solved elsewhere in this codebase (the Facility
    tab click in ecw/facility_filter.py, blocked by an empty sibling
    badge) - force=True bypasses Playwright's actionability/overlap check
    and clicks the element's coordinates directly. Kept the outer
    retry/wait loop too as a second layer, in case the widget is still
    mid-render rather than just overlapped.
    """
    last_error = None
    for attempt in range(retries):
        for desc in descriptions_to_try:
            try:
                await iframe.get_by_role("option", name=day_str, description=desc, exact=True).click(timeout=5000, force=True)
                log.info(f"{label} date set! (description='{desc}')")
                return
            except Exception:
                pass
        # Final fallback - no description
        try:
            await iframe.get_by_role("option", name=day_str, exact=True).first.click(timeout=5000, force=True)
            log.info(f"{label} date set! (no description fallback)")
            return
        except Exception as e:
            last_error = e
            log.info(f"Could not click {label} date option (attempt {attempt + 1}/{retries}) - retrying...")
            await asyncio.sleep(2)
    raise RuntimeError(f"Could not set {label} date to day {day_str} after {retries} attempts: {last_error}")


async def run(window_days=7, start_offset_days=0):
    """
    start_offset_days: how many days after today the window STARTS
    (default 0 = starts today).
    window_days: size of the export window, counted from the start date
    (start_date -> start_date+window_days).

    Production (main.py) calls this with start_offset_days=1, window_days=2
    (2026-07-21 change: tomorrow through +3 days, i.e. today+1/+2/+3 -
    never includes today). The demo pipeline (main_demo.py) is unchanged -
    window_days=2 with the default start_offset_days=0 (today/+1/+2),
    matching the reference project's main_1.py demo variant.
    """
    log.info("=" * 50)
    log.info(f"STEP 0 - ECW: EXPORTING SCHEDULE ({'TODAY' if start_offset_days == 0 else f'TODAY + {start_offset_days} DAYS'} THROUGH +{window_days} MORE DAYS)")
    log.info("=" * 50)

    # Both 2-letter and 3-letter day description variants
    day_desc_2 = {0: "Mo", 1: "Tu", 2: "We", 3: "Th", 4: "Fr", 5: "Sa", 6: "Su"}
    day_desc_3 = {0: "Mon", 1: "Tue", 2: "Wed", 3: "Thu", 4: "Fri", 5: "Sat", 6: "Sun"}

    start_date = date.today() + timedelta(days=start_offset_days)
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
        # (2026-07-24) force=True on these: live runs showed different
        # transient backdrop elements (#pnBackDrop, seen after #load's own
        # overlay was already confirmed hidden) intermittently blocking
        # these clicks - a different element name each time, not the same
        # bug recurring. Same fix already proven for the calendar day
        # click and the Facility tab click elsewhere in this codebase.
        log.info("Navigating to eBO Reports...")
        await page.locator("#jellybean-panelLink4").click(force=True)
        await asyncio.sleep(1)
        await page.get_by_text("Menu", exact=True).click(force=True)
        await asyncio.sleep(1)
        await page.locator("#pane6").get_by_text("Reports").click(force=True)
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
        await _wait_for_report_running_modal_gone(page1)

        # --- WAIT FOR IFRAME ---
        log.info("Waiting for iframe to load...")
        iframe = page1.locator("iframe[name=\"iD6D96C5E47F347C9B95828AC68A2D69B\"]").content_frame
        await iframe.get_by_role("img").first.wait_for(timeout=120000)

        # (2026-07-24 fix) A live diagnostic showed the iframe's own
        # content is unstable right after this point - images can appear
        # then disappear again moments later as the report prompt panel
        # (tabs: Additional Prompts / Facility / Provider / Payer /
        # Patient / Others) keeps re-rendering. The date-range controls
        # need TWO calendar icons (img().first for start, img().nth(1) for
        # end) - a live run showed the count stabilizing at just 1, which
        # the previous version of this check accepted as "done" since it
        # only required non-zero, not the actual expected count. Now waits
        # up to 90s (was 30) and requires the count to reach >= 2 before
        # considering it settled - same stabilization pattern already
        # proven for the Facility results list in ecw/facility_filter.py,
        # now also checking for completeness, not just stability.
        previous_count = -1
        stable_checks = 0
        target_reached = False
        for _ in range(90):
            count = await iframe.get_by_role("img").count()
            if count >= 2 and count == previous_count:
                stable_checks += 1
                if stable_checks >= 3:
                    target_reached = True
                    break
            else:
                stable_checks = 0
            previous_count = count
            await asyncio.sleep(1)
        if target_reached:
            log.info(f"Iframe content stabilized at {previous_count} image(s).")
        else:
            log.info(f"Iframe content never reached 2+ stable images after extended wait "
                     f"(last count: {previous_count}) - proceeding anyway.")

        # (2026-07-24 fix) The "Your report is running" modal was observed
        # live appearing AFTER the check above already passed - likely a
        # report from an earlier attempt still executing/queued
        # server-side, only surfacing once the page caught up. That modal
        # sits on top of the date picker and blocks every click under it
        # (including the day option), which looked like a calendar bug but
        # wasn't - re-check right before interacting with dates.
        await _wait_for_report_running_modal_gone(page1)

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
