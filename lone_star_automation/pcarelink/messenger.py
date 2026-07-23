"""
pcarelink/messenger.py
--------------------------
Ported from the reference project's pcarelink_send_messages(), wired into
both main.py and main_demo.py (per explicit confirmation).

**ARCHITECTURE CHANGE (2026-07-21)**: the "Filter by Practice" selection
is no longer a single fixed value (settings.PCARELINK_PRACTICE). It is now
resolved PER PATIENT from that patient's own "facility" field (carried
from the eCW Excel export - see patient_forms_now/form_sender.py and
main_demo.py) via settings.resolve_practice_for_facility(). One shared
PCareLink account (aasha@painmedpa.com, same login used by both clinics'
projects) covers multiple practices, and different patients can belong to
different practices - a single fixed filter was wrong for anyone not at
that one practice.

Lone Star's own facility ("Lone Star Pediatrics Midlothian") is NOT
currently in settings.FACILITY_TO_PRACTICE (it doesn't match any of the
practices confirmed present in ReachMyDr's dropdown - see settings.py's
comment) - so under this change, EVERY Lone Star patient will be skipped
with a logged warning rather than messaged under a guessed/wrong practice,
until the correct mapping is known.
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

        for patient in patients:
            try:
                practice = settings.resolve_practice_for_facility(patient.get("facility"))

                # --- TEMPORARY DEBUG LOGGING (per explicit request - remove once verified live) ---
                log.info(f"[DEBUG] Patient: {patient['first_name']} {patient['last_name']} | "
                         f"Facility: {patient.get('facility')!r} | Practice: {practice!r}")

                if not practice:
                    log.warning(f"No ReachMyDr practice mapping for facility {patient.get('facility')!r} "
                                f"(acct {patient['acct_no']}) - skipping message, NOT guessing a practice.")
                    continue

                log.info(f"Sending message for {patient['acct_no']} ({patient['last_name']} {patient['first_name']})")
                await page.get_by_role("button", name="Filter by Practice").click()
                await page.get_by_text(practice, exact=False).click()
                await page.wait_for_timeout(2000)
                log.info(f"Filtered by practice: {practice}")

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
                    await page.get_by_role("button", name=practice).click(timeout=5000)
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
