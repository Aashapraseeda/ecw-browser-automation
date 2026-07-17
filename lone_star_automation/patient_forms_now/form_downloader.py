"""
patient_forms_now/form_downloader.py
----------------------------------------
STEP 3 - Patient Forms Now: check for completed forms and download PDFs.

REDESIGNED vs. the reference project: the reference clinic's Pediforms
checks completion from the SAME "Today's Patients" search/View flow used
for sending. Lone Star's PFN instead has a dedicated top-level "Completed
Forms" nav section (confirmed via a live screenshot: columns Patient |
Form | Status | Reference # | Completed on | Actions, with a Download PDF
button per row once a submission is actually done - rows still in
progress show "Available once completed" instead). This module now
navigates there and searches by patient name, rather than searching
"Today's Patients" and clicking View.

Single pass only - cron re-invoking the whole pipeline provides the
"check again later" behavior.
"""

import os

from playwright.async_api import async_playwright

from config import settings
from database import state_db
from patient_forms_now.login import pfn_login
from utils.logger import get_logger

log = get_logger(__name__)


def ensure_patient_folder(patient):
    folder_path = os.path.join(settings.DOC_FOLDER, patient["folder_name"])
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        log.info(f"Created folder: {folder_path}")
    return folder_path


async def check_and_download_completed(page, patients):
    """
    Navigates to the "Completed Forms" section and searches each patient
    by name. If a Download PDF button is present for them, downloads it.
    """
    log.info("Checking for completed forms...")
    newly_downloaded = []

    await page.get_by_role("link", name="Completed Forms").click()
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(1000)

    for patient in patients:
        # Not finding/being able to search is expected whenever the parent
        # hasn't completed the form yet - treat it the same as "not
        # completed", not as an error, and just check again next run.
        try:
            search_box = page.get_by_placeholder("Search by patient, form, or reference #...")
            await search_box.click(timeout=10000)
            await search_box.fill(f"{patient['first_name']} {patient['last_name']}")
            await page.wait_for_timeout(2000)
        except Exception:
            log.info(f"Patient {patient['acct_no']} not completed yet")
            continue

        try:
            download_buttons = page.get_by_role("button", name="Download PDF")
            download_count = await download_buttons.count()

            if download_count == 0:
                log.info(f"Patient {patient['acct_no']} not completed/downloadable yet")
                await search_box.fill("")
                await page.wait_for_timeout(500)
                continue

            log.info(f"Patient {patient['acct_no']} has a downloadable completed form")

            folder_path = ensure_patient_folder(patient)
            file_name = f"{patient['last_name']}_{patient['first_name']}_{patient['form_filename']}.pdf"
            save_path = os.path.join(folder_path, file_name)

            log.info(f"Downloading completed form for {patient['acct_no']}...")
            async with page.expect_download() as download_info:
                await download_buttons.first.click()
            download = await download_info.value
            await download.save_as(save_path)
            log.info(f"Saved: {save_path}")

            state_db.mark_downloaded(patient["acct_no"], patient["appointment_date"])
            newly_downloaded.append(patient)

            await search_box.fill("")
            await page.wait_for_timeout(500)

        except Exception as e:
            log.info(f"Error downloading for {patient['acct_no']}: {e}")
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
