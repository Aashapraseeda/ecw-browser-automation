"""
modules/pediform/form_downloader.py
------------------------------------
Polls the PediForm Submissions page for completed forms and downloads PDFs.
"""

import asyncio
import os
from typing import List, Dict

from playwright.async_api import Page

from config.settings import COMPLETED_FORMS_DOWNLOAD_DIR
from modules.pediform.selectors import (
    NAV_SUBMISSIONS,
    SUBMISSION_TABLE_ROWS,
    STATUS_COMPLETED,
    DOWNLOAD_PDF_BTN,
)
from utils.logger import get_logger

logger = get_logger(__name__)

os.makedirs(COMPLETED_FORMS_DOWNLOAD_DIR, exist_ok=True)


async def download_completed_forms(page: Page, patients: List[Dict]) -> List[Dict]:
    """
    Navigate to /staff/submissions, find rows with Completed status,
    download the PDF for each patient in our eligible list.

    Args:
        page:     Playwright page logged into PediForm.
        patients: Eligible patient dicts (must include 'full_name').

    Returns:
        List of patient dicts for whom a PDF was successfully downloaded.
    """
    logger.info("Navigating to Submissions page...")
    await page.locator(NAV_SUBMISSIONS).click()
    await asyncio.sleep(3)

    patient_names = {p["full_name"].strip().lower() for p in patients}
    downloaded = []

    rows = await page.locator(SUBMISSION_TABLE_ROWS).all()
    logger.info(f"Found {len(rows)} submission row(s).")

    for row in rows:
        # Check status
        completed_span = row.locator(STATUS_COMPLETED)
        if await completed_span.count() == 0:
            continue  # not completed

        # Extract patient name from the row (first cell)
        try:
            row_name = (await row.locator("td").first.inner_text()).strip().lower()
        except Exception:
            continue

        if row_name not in patient_names:
            continue

        # Download PDF
        try:
            download_btn = row.locator(DOWNLOAD_PDF_BTN)
            if await download_btn.count() == 0:
                logger.warning(f"No download button for {row_name}")
                continue

            safe_name = row_name.replace(" ", "_").replace("/", "-")
            filename = os.path.join(COMPLETED_FORMS_DOWNLOAD_DIR, f"{safe_name}.pdf")

            async with page.expect_download() as dl_info:
                await download_btn.click()
            download = await dl_info.value
            await download.save_as(filename)

            logger.info(f"Downloaded: {filename}")
            patient = next(
                (p for p in patients if p["full_name"].strip().lower() == row_name),
                None,
            )
            if patient:
                downloaded.append(patient)

        except Exception as e:
            logger.error(f"Error downloading PDF for {row_name}: {e}")

    logger.info(f"PDFs downloaded: {len(downloaded)} / {len(patients)} eligible")
    return downloaded
