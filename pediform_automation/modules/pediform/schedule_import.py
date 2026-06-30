"""
modules/pediform/schedule_import.py
------------------------------------
Uploads the PediForm import Excel and waits for the patient table to populate.
"""

import asyncio
from playwright.async_api import Page

from modules.pediform.selectors import SCHEDULE_FILE_INPUT, PATIENT_TABLE_ROWS
from utils.logger import get_logger

logger = get_logger(__name__)


async def import_schedule(page: Page, excel_path: str) -> int:
    """
    Upload the PediForm-format Excel and return the number of table rows after import.

    Args:
        page:       Playwright page already on Today's Patients.
        excel_path: Path to the PediForm import Excel (dates as text).

    Returns:
        Number of patient rows visible in the table after import.
    """
    logger.info(f"Importing schedule: {excel_path}")

    # Set the file on the hidden file input
    file_input = page.locator(SCHEDULE_FILE_INPUT)
    await file_input.wait_for(timeout=15000)
    await file_input.set_input_files(excel_path)
    logger.info("File selected.")

    # Click Import schedule button
    await page.get_by_role("button", name="Import schedule").click()
    logger.info("Import schedule clicked — waiting for table...")

    # Wait for the table to render (up to 60 s)
    await page.wait_for_selector(PATIENT_TABLE_ROWS, timeout=60000)
    await asyncio.sleep(2)

    row_count = await page.locator(PATIENT_TABLE_ROWS).count()
    logger.info(f"Table updated: {row_count} row(s) visible after import.")
    return row_count
