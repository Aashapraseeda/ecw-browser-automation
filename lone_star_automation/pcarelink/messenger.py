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
from utils import failure_report

log = get_logger(__name__)

PATIENT_SEARCH_ATTEMPTS = 3
PATIENT_SEARCH_RETRY_WAIT_MS = 1500


async def _find_and_open_patient(page, patient):
    """
    Searches for a patient by account number and opens their result.
    Retries the WHOLE search cycle (re-click, re-fill, re-attempt) up to
    PATIENT_SEARCH_ATTEMPTS times rather than a single fill + fixed sleep +
    one lookup - live logs (from the identical code path in the reference
    project) showed roughly 1 in 3 lookups failing on the first attempt
    with no consistent pattern by patient, the signature of a timing/race
    issue (or a search box that doesn't always register a single .fill()
    the same way real keystrokes would) rather than a genuine "this
    patient isn't in ReachMyDr" case. Returns True/False.
    """
    target_text = f"{patient['last_name'].upper()}, {patient['first_name'].upper()}"
    search_box = page.get_by_role("searchbox", name="Enter patient first name or")
    for attempt in range(PATIENT_SEARCH_ATTEMPTS):
        await search_box.click()
        await search_box.fill("")
        await search_box.fill(patient["acct_no"])
        try:
            await page.get_by_text(target_text).first.click(timeout=8000)
            return True
        except Exception:
            if attempt < PATIENT_SEARCH_ATTEMPTS - 1:
                log.info(f"Patient {patient['acct_no']} not found on attempt "
                         f"{attempt + 1}/{PATIENT_SEARCH_ATTEMPTS} - retrying...")
                await page.wait_for_timeout(PATIENT_SEARCH_RETRY_WAIT_MS)
    return False


async def send_messages(patients):
    log.info("=" * 50)
    log.info("STEP 2 - PCARELINK: SENDING MESSAGES")
    log.info("=" * 50)

    messaged, failed = [], []

    # (fix 2026-07-30) Previously had no session-level guard - if login (or
    # anything else outside the per-patient loop below) raised, the
    # exception propagated all the way out uncaught, which meant the
    # failure report below never even got built/saved. Now, on a total
    # session crash, every patient not already accounted for by the
    # per-patient loop gets recorded as failed with a clear reason, so the
    # report always reflects reality - even in the crash case - instead of
    # silently not existing for this run.
    try:
        await _send_messages_session(patients, messaged, failed)
    except Exception as e:
        log.error(f"ReachMyDr session failed: {e}")
        processed_accts = {p["acct_no"] for p in messaged} | {f["acct_no"] for f in failed}
        for patient in patients:
            if patient.get("acct_no") not in processed_accts:
                failure_report.record_failure(failed, patient, f"ReachMyDr session failed: {e}")

    failure_report.print_and_save_failure_report(failed, settings.LOG_DIR)

    return messaged, failed


async def _send_messages_session(patients, messaged, failed):
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

        # The practice filter button's own visible text is NOT stable across
        # patients: it reads "Filter by Practice" only before anything has
        # ever been selected, then switches to showing whichever practice
        # name was picked last (e.g. "Lone Star Pediatrics") and stays that
        # way - it never reverts to the placeholder. Track what it's
        # currently showing ourselves (we're the ones who set it) instead of
        # hardcoding either the placeholder or any specific practice name,
        # so this works for any practice, in any order, indefinitely.
        current_filter_label = "Filter by Practice"

        for patient in patients:
            practice = None
            try:
                practice = settings.resolve_practice_for_facility(patient.get("facility"))

                # --- TEMPORARY DEBUG LOGGING (per explicit request - remove once verified live) ---
                log.info(f"[DEBUG] Patient: {patient['first_name']} {patient['last_name']} | "
                         f"Facility: {patient.get('facility')!r} | Practice: {practice!r}")

                if not practice:
                    log.warning(f"No ReachMyDr practice mapping for facility {patient.get('facility')!r} "
                                f"(acct {patient['acct_no']}) - skipping message, NOT guessing a practice.")
                    failure_report.record_failure(failed, patient, "Practice mapping not found")
                    continue

                log.info(f"Sending message for {patient['acct_no']} ({patient['last_name']} {patient['first_name']})")

                # --- practice filter selection (mechanics UNCHANGED from the
                # earlier fix - only wrapped here to tag WHICH step failed) ---
                try:
                    try:
                        await page.get_by_role("button", name=current_filter_label).click(timeout=10000)
                    except Exception:
                        # Fallback in case our tracked label ever gets out of
                        # sync with what's actually on screen - the placeholder
                        # is the only other state this button can be in.
                        await page.get_by_role("button", name="Filter by Practice").click(timeout=10000)
                    try:
                        # Scoped to the open menu's own container
                        # (#menu-clinics, confirmed from a live strict-mode
                        # error dump) - a bare page-wide get_by_text matches
                        # BOTH the freshly-opened menu item AND the closed
                        # dropdown's own leftover display text whenever the
                        # same practice is chosen twice in a row, which
                        # crashes with a strict-mode violation and leaves
                        # the dropdown open for the next patient.
                        await page.locator("#menu-clinics").get_by_text(practice, exact=False).click(timeout=8000)
                    except Exception:
                        # Last-resort fallback so a DOM change here degrades
                        # to "may pick the wrong element" rather than
                        # crashing the whole patient - .first guarantees no
                        # strict-mode error.
                        await page.get_by_text(practice, exact=False).first.click(timeout=5000)
                    await page.wait_for_timeout(2000)
                    current_filter_label = practice
                    log.info(f"Filtered by practice: {practice}")
                except Exception as e:
                    log.info(f"Error selecting practice: {e}")
                    reason = "Timeout during practice selection" if "Timeout" in str(e) else "Practice selection failed"
                    failure_report.record_failure(failed, patient, reason, practice=practice)
                    continue

                # --- patient search (retry helper UNCHANGED from the earlier fix) ---
                found = await _find_and_open_patient(page, patient)
                if not found:
                    log.info(f"Patient {patient['acct_no']} not found after {PATIENT_SEARCH_ATTEMPTS} attempts - skipping")
                    failure_report.record_failure(failed, patient, "Patient not found in ReachMyDr", practice=practice)
                    continue
                await page.wait_for_timeout(1000)

                # --- open the message drawer ---
                try:
                    await page.locator('[data-test-id="pcl-payments-sendMessageLinkGuarantorDrawer"]').click()
                    await page.wait_for_timeout(1000)
                except Exception as e:
                    log.info(f"Error opening message drawer: {e}")
                    reason = "Timeout opening message drawer" if "Timeout" in str(e) else f"Error opening message drawer: {e}"
                    failure_report.record_failure(failed, patient, reason, practice=practice)
                    continue

                # message-type selection - UNCHANGED, still non-fatal
                try:
                    await page.get_by_role("button", name=practice).click(timeout=5000)
                    await page.wait_for_timeout(500)
                    await page.get_by_role("menuitem", name="Appointment Scheduling").get_by_role("radio").check(timeout=5000)
                    await page.locator("#menu- > div").first.click(timeout=5000)
                    await page.wait_for_timeout(500)
                except Exception:
                    log.info("Message type selection skipped")

                # --- fill and send the message ---
                try:
                    message_box = page.get_by_role("textbox", name="Type your response and send")
                    await message_box.click()
                    await message_box.fill(settings.PCARELINK_MESSAGE)
                    log.info("Message typed!")
                    await page.locator('[data-test-id="pcl-payments-sendMessageButton"]').click()
                    log.info("Message sent!")
                    messaged.append(patient)
                except Exception as e:
                    log.info(f"Error sending message: {e}")
                    reason = "Timeout while sending the message" if "Timeout" in str(e) else f"Error sending message: {e}"
                    failure_report.record_failure(failed, patient, reason, practice=practice)
                    continue

                # --- best-effort cleanup only - the message already sent
                # successfully above, so a failure here must NOT retroactively
                # mark this patient as failed (that would put them in BOTH
                # the messaged and failed lists).
                await page.wait_for_timeout(1000)
                try:
                    await page.locator('[data-test-id="pcl-appointments-closePatientsDetails"]').click()
                except Exception:
                    pass
                await page.wait_for_timeout(1000)
            except Exception as e:
                # True catch-all for anything unexpected not already
                # handled by one of the specific blocks above.
                log.info(f"Error: {e}")
                failure_report.record_failure(failed, patient, f"Unexpected error: {e}", practice=practice)
                try:
                    await page.locator('[data-test-id="pcl-appointments-closePatientsDetails"]').click()
                except Exception:
                    pass
                continue

        await browser.close()
        log.info("Messages sent!")
