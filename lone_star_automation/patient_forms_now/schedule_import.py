"""
patient_forms_now/schedule_import.py
---------------------------------------
Uploads the FULL (unfiltered) eCW export into Patient Forms Now.

Workflow change vs. the reference project: the reference pre-filtered the
Excel to ASQ-eligible rows before uploading (FILTERED_EXCEL_PATH). For
Lone Star, the entire exported schedule is imported as-is; Well-Check
eligibility is decided locally (patient_forms_now/form_sender.py) using
the same VISIT_TYPE_TO_FORM mapping, driving which patients get searched
and sent a form - it no longer gates what gets uploaded.
"""

from config import settings
from utils.logger import get_logger

log = get_logger(__name__)


async def import_schedule(page):
    log.info("Uploading full schedule Excel...")
    await page.get_by_role("button", name="Choose File").set_input_files(settings.EXCEL_PATH)
    await page.get_by_role("button", name="Import schedule").click()
    await page.wait_for_load_state("networkidle")
    log.info("Schedule imported!")

    await page.get_by_role("combobox").nth(4).select_option("week")
    await page.wait_for_timeout(1000)
