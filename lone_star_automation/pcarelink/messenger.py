"""
pcarelink/messenger.py
--------------------------
Ported from the reference project's pcarelink_send_messages() EXACTLY,
now wired into both main.py and main_demo.py (per explicit confirmation).

PCARELINK_PRACTICE is set to "Pediatric Center Of Round Rock" (the
reference clinic's practice name, unchanged) - per explicit instruction,
this is reused as-is rather than adapted for Lone Star, since PCareLink/
ReachMyDr is a single shared account (same login, aasha@painmedpa.com)
that only has this one practice configured regardless of which clinic/
facility a patient belongs to in eCW or Patient Forms Now. The practice-
filter click below is the exact original hardcoded pattern
(f"{PCARELINK_PRACTICE}Round Rock, us"), not adapted for Lone Star.
"""

from playwright.async_api import async_playwright

from config import settings
from utils.logger import get_logger

log = get_logger(__name__)


async def send_messages(patients):
    log.info("=" * 50)
    log.info("STEP 2 - PCARELINK: SENDING MESSAGES")
    log.info("=" * 50)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=300)
        context = await browser.new_context()
        page = await context.new_page()

        log.info("Logging into pcarelink...")
        await page.goto("https://app.pcarelink.com/login", timeout=60000, wait_until="domcontentloaded")
        await page.wait_for_timeout(3000)
        await page.get_by_role("textbox", name="Enter email id").fill(settings.PCARELINK_EMAIL)
        await page.get_by_role("textbox", name="Enter password").fill(settings.PCARELINK_PASSWORD)
        await page.locator('[data-test-id="pcl-login-signInButton"]').click()
        await page.wait_for_timeout(5000)
        log.info("Logged in!")

        await page.locator('[data-test-id="pcl-menuDropDownComponent"]').click()
        await page.locator('[data-test-id="pcl-dashboard-popOver1"]').click()
        await page.wait_for_load_state("networkidle")

        await page.get_by_role("button", name="Filter by Practice").click()
        await page.get_by_text(f"{settings.PCARELINK_PRACTICE}Round Rock, us").click()
        await page.wait_for_timeout(2000)
        log.info(f"Filtered by practice: {settings.PCARELINK_PRACTICE}")

        for patient in patients:
            try:
                log.info(f"Sending message for {patient['acct_no']} ({patient['last_name']} {patient['first_name']})")
                search_box = page.get_by_role("searchbox", name="Enter patient first name or")
                await search_box.click()
                await search_box.fill(patient["acct_no"])
                await page.wait_for_timeout(2000)
                try:
                    await page.get_by_text(f"{patient['last_name'].upper()}, {patient['first_name'].upper()}").first.click(timeout=10000)
                except Exception:
                    try:
                        await page.locator(".patient-result, .search-result").first.click(timeout=5000)
                    except Exception:
                        log.info(f"Patient {patient['acct_no']} not found - skipping")
                        continue
                await page.wait_for_timeout(1000)
                await page.locator('[data-test-id="pcl-payments-sendMessageLinkGuarantorDrawer"]').click()
                await page.wait_for_timeout(1000)
                try:
                    await page.get_by_role("button", name=settings.PCARELINK_PRACTICE).click(timeout=5000)
                    await page.wait_for_timeout(500)
                    await page.get_by_role("menuitem", name="Appointment Scheduling").get_by_role("radio").check(timeout=5000)
                    await page.locator("#menu- > div").first.click(timeout=5000)
                    await page.wait_for_timeout(500)
                except Exception:
                    log.info("Message type selection skipped")
                message_box = page.get_by_role("textbox", name="Type your response and send")
                await message_box.click()
                await message_box.fill(settings.PCARELINK_MESSAGE)
                log.info("Message typed!")
                await page.locator('[data-test-id="pcl-payments-sendMessageButton"]').click()
                log.info("Message sent!")
                await page.wait_for_timeout(1000)
                await page.locator('[data-test-id="pcl-appointments-closePatientsDetails"]').click()
                await page.wait_for_timeout(1000)
            except Exception as e:
                log.info(f"Error: {e}")
                try:
                    await page.locator('[data-test-id="pcl-appointments-closePatientsDetails"]').click()
                except Exception:
                    pass
                continue

        await browser.close()
        log.info("Messages sent!")
