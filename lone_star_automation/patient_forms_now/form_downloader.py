"""
patient_forms_now/form_downloader.py
----------------------------------------
STEP 3 - Patient Forms Now: check for completed forms and download PDFs.

**(2026-07-23) MIGRATED to match the reference clinic's (Nurture Kids)
navigation pattern**: searches "Today's Patients" (the SAME list used for
sending forms) by account number, opens View, and checks for a
"Completed" status - no longer uses the separate "Completed Forms" nav
section at all (confirmed live, read-only, that this function never
navigates there anymore).

One real, confirmed-live difference from Nurture Kids that could NOT be
eliminated: Nurture Kids checks for "Completed" text on the Today's
Patients LIST ROW itself, before opening View (an efficiency
short-circuit). Lone Star's list row shows a DIFFERENT status vocabulary
there (its own "Form status" column shows values like "Downloaded", never
the literal word "Completed" - confirmed live) - so the "Completed" check
here happens AFTER opening View instead, against the View page's own
"Sent forms" table (which DOES show a "Completed" badge there, matching
Nurture Kids' check exactly, just relocated). This costs one extra
navigation (View is always opened, even for not-yet-completed patients)
but was necessary - checking the list row for "Completed" text would have
incorrectly skipped every genuinely-completed Lone Star patient.

Also confirmed live: the section that actually lists individual form
submissions on the View page - "Submission Exports" - renders
ASYNCHRONOUSLY, after Playwright's `networkidle` already reports done. An
earlier read-only inspection this session caught this section still
empty ("No submissions linked") despite two real completed submissions
existing, purely because the read happened before the async render
finished. A fixed extra wait after opening View works around this - see
SUBMISSION_EXPORTS_RENDER_WAIT_MS below.

Each submission renders as its own `<div class="card">` (confirmed via
DOM inspection) containing: a UUID + status text (e.g. "(exported)" or
"(in_progress)"), an "Export PDF" button, and a "View/Edit Responses"
toggle that - once clicked (read-only click; "Save Responses" is never
clicked) - reveals a "Template: <name> vN" line INSIDE THAT SAME CARD.
This template name is the ONLY place a submission's specific form is
identified (the card's UUID/status line alone doesn't say which form it
is) - it's matched (normalized: lowercased, non-alphanumeric stripped) to
this patient's expected form names, giving the same accurate, form-
specific filenames as the previous "Completed Forms"-based implementation
(e.g. "Alford_Astoria_ASQ_36_Months.pdf"), not a generic fallback.
Submissions still "(in_progress)" are skipped entirely (not yet ready to
export). Filenames are deterministic, so an already-downloaded submission
is detected via os.path.exists and skipped, not re-downloaded.

A patient is only marked 'downloaded' (state_db.mark_downloaded, which
hands them to the upload step) once EVERY expected form has a matching
file already present in their folder - a patient with e.g. ASQ done but
M-CHAT/TB still pending stays in 'form_sent' and is re-checked again next
run, rather than being finalized on a partial capture. See
patient_forms_now/form_sender.py's module docstring for the limitation
this implies if a parent never completes every expected form.

Single pass only - cron re-invoking the whole pipeline provides the
"check again later" behavior.
"""

import os
import re

from playwright.async_api import async_playwright

from config import settings
from database import state_db
from patient_forms_now.login import pfn_login
from utils.logger import get_logger

log = get_logger(__name__)

# Confirmed live (2026-07-23): the "Submission Exports" section on a
# patient's View page renders asynchronously - a bare `networkidle` wait
# was observed catching it still empty. This fixed extra wait after
# opening View works around that race.
SUBMISSION_EXPORTS_RENDER_WAIT_MS = 3000


def ensure_patient_folder(patient):
    folder_path = os.path.join(settings.DOC_FOLDER, patient["folder_name"])
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        log.info(f"Created folder: {folder_path}")
    return folder_path


def _normalize(text):
    return re.sub(r"[^a-z0-9]", "", text.lower())


async def check_and_download_completed(page, patients):
    """
    Searches "Today's Patients" by account number (same list/search box
    used for sending), opens View, and checks the "Sent forms" table for
    a "Completed" badge - see module docstring for why this check happens
    after opening View rather than on the list row beforehand (a
    necessary deviation from Nurture Kids' exact check position, forced
    by Lone Star's list row using different status text).

    Downloads EVERY completed submission for a patient, not just the
    first. Each submission is its own card in "Submission Exports"; for
    ones not still "(in_progress)", this expands "View/Edit Responses"
    (read-only) to read that card's own "Template: <name> vN" line and
    match it to this patient's expected forms - see module docstring for
    the full DOM/matching detail. Filenames are deterministic, so an
    already-downloaded submission is detected via os.path.exists and
    skipped rather than re-downloaded.
    """
    log.info("Checking for completed forms...")
    newly_downloaded = []

    for patient in patients:
        try:
            search_box = page.get_by_role("textbox", name="Search…")
            await search_box.click()
            await search_box.fill(patient["acct_no"])
            await page.wait_for_timeout(1500)

            view_count = await page.get_by_role("link", name="View").count()
            if view_count == 0:
                log.info(f"Patient {patient['acct_no']} not found")
                await search_box.fill("")
                await page.wait_for_timeout(500)
                continue

            await page.get_by_role("link", name="View").first.click()
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(SUBMISSION_EXPORTS_RENDER_WAIT_MS)

            completed_visible = await page.get_by_text("Completed", exact=True).count()
            if completed_visible == 0:
                log.info(f"Patient {patient['acct_no']} not completed yet")
                await page.get_by_role("link", name="← Back to today's patients").click()
                await page.wait_for_load_state("networkidle")
                continue

            log.info(f"Patient {patient['acct_no']} has Completed form(s) - checking Submission Exports...")

            folder_path = ensure_patient_folder(patient)
            # Reconstructed from the comma-joined summary built in
            # read_eligible_patients_from_excel() - safe to split on ","
            # since individual form names never contain commas.
            expected_names = [n.strip() for n in patient.get("form_name", "").split(",") if n.strip()]
            expected_norm = {_normalize(n): n for n in expected_names}

            # Anchor on each specific "Export PDF" button first, then scope
            # to its NEAREST enclosing div.card via an ancestor lookup.
            # (2026-07-23 fix, found via live testing: filtering
            # `div.card` broadly by "has an Export PDF descendant" also
            # matched an outer wrapper card containing ALL submissions'
            # buttons nested inside it - not just the innermost per-
            # submission card - causing a strict-mode multi-match error.
            # Anchoring on the button and walking up avoids this entirely.)
            export_buttons = page.get_by_role("button", name="Export PDF")
            card_count = await export_buttons.count()
            log.info(f"Found {card_count} submission(s) in Submission Exports for {patient['acct_no']}")

            for i in range(card_count):
                button = export_buttons.nth(i)
                card = button.locator("xpath=ancestor::div[@class='card'][1]")
                try:
                    card_text = await card.inner_text()
                except Exception:
                    continue

                if "in_progress" in card_text.lower():
                    log.info(f"Submission {i + 1}/{card_count} still in progress - skipping")
                    continue

                form_label = None
                try:
                    view_edit_btn = card.get_by_role("button", name="View/Edit Responses")
                    if await view_edit_btn.count() > 0:
                        await view_edit_btn.click()
                        await page.wait_for_timeout(1500)
                        card_text = await card.inner_text()
                    for line in card_text.splitlines():
                        stripped = line.strip()
                        if stripped.startswith("Template:"):
                            template_name = stripped.split(":", 1)[1].strip().split(" ")[0]
                            matched = expected_norm.get(_normalize(template_name))
                            if matched:
                                form_label = matched
                            break
                except Exception:
                    pass
                if form_label is None:
                    form_label = f"completed form {i + 1}"

                filename_part = re.sub(r"[^A-Za-z0-9]+", "_", form_label).strip("_")
                file_name = f"{patient['last_name']}_{patient['first_name']}_{filename_part}.pdf"
                save_path = os.path.join(folder_path, file_name)

                if os.path.exists(save_path):
                    log.info(f"Already downloaded: {file_name} - skipping")
                    continue

                try:
                    log.info(f"Downloading submission {i + 1}/{card_count} for {patient['acct_no']} ({form_label})...")
                    async with page.expect_download() as download_info:
                        await button.click()
                    download = await download_info.value
                    await download.save_as(save_path)
                    log.info(f"Saved: {save_path}")
                except Exception as e:
                    log.info(f"Error downloading submission {i + 1} for {patient['acct_no']}: {e}")
                    continue

            existing_files = os.listdir(folder_path) if os.path.exists(folder_path) else []
            all_captured = all(
                any(_normalize(name) in _normalize(f) for f in existing_files)
                for name in expected_names
            ) if expected_names else len(existing_files) > 0

            if all_captured:
                state_db.mark_downloaded(patient["acct_no"], patient["appointment_date"])
                newly_downloaded.append(patient)
                log.info(f"All expected form(s) captured for {patient['acct_no']} - ready for upload.")
            else:
                log.info(f"Not all expected forms captured yet for {patient['acct_no']} - will re-check next run.")

            await page.get_by_role("link", name="← Back to today's patients").click()
            await page.wait_for_load_state("networkidle")

        except Exception as e:
            log.info(f"Error downloading for {patient['acct_no']}: {e}")
            try:
                await page.get_by_role("link", name="← Back to today's patients").click()
            except Exception:
                pass
            continue

    return newly_downloaded


async def run(patients):
    """
    One single pass over `patients` (pending patients pulled from state_db,
    not just today's Excel). Returns list of newly-downloaded patients.
    """
    log.info("=" * 50)
    log.info("STEP 3 - PATIENT FORMS NOW: CHECKING FOR COMPLETED FORMS")
    log.info("=" * 50)

    if not patients:
        log.info("No patients pending form completion.")
        return []

    try:
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=False, slow_mo=300)
            context = await browser.new_context()
            page = await context.new_page()

            await pfn_login(page)

            newly_downloaded = await check_and_download_completed(page, patients)
            await browser.close()
            return newly_downloaded

    except Exception as e:
        log.info(f"Check failed: {e}")
        log.info("Will retry on next scheduled cron run.")
        return []
