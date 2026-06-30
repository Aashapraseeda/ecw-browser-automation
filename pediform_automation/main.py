"""
main.py
-------
Automation orchestrator — currently runs Phase 1 + Phase 2.

Phase 1: Read ECW export → filter eligible patients → build PediForm import Excel
Phase 2: PediForm — login → import schedule → send age-appropriate forms

Pending (not yet active):
  Phase 3: ReachMyDr reminders
  Phase 4: Download completed PDFs
  Phase 5: ECW Chart Documents upload
"""

import asyncio
import os
from playwright.async_api import async_playwright

from config.settings import SCHEDULE_EXCEL_PATH, PEDIFORM_IMPORT_PATH
from utils.excel_reader import get_eligible_patients
from utils.ecw_to_pediform import build_pediform_excel
from utils.logger import get_logger

from modules.pediform.login import login as pf_login, navigate_to_todays_patients, new_pediform_context
from modules.pediform.schedule_import import import_schedule
from modules.pediform.form_sender import send_forms_for_all

logger = get_logger("main")


async def run():
    # ── Phase 1: Read ECW export & filter eligible patients ────────────────────
    logger.info("=" * 60)
    logger.info("PHASE 1: Reading ECW schedule")
    logger.info("=" * 60)
    patients = get_eligible_patients(SCHEDULE_EXCEL_PATH)
    if not patients:
        logger.warning("No eligible patients found — nothing to do.")
        return

    for p in patients:
        logger.info(f"  {p['full_name']} | {p['age_bucket']} | Form: {p['form_name']}")

    # ── Phase 1b: Build PediForm import Excel ──────────────────────────────────
    logger.info("Building PediForm import Excel...")
    os.makedirs(os.path.dirname(PEDIFORM_IMPORT_PATH), exist_ok=True)
    pediform_excel = build_pediform_excel(patients, PEDIFORM_IMPORT_PATH)
    logger.info(f"Import file ready: {pediform_excel}")

    async with async_playwright() as p:
        browser, context = await new_pediform_context(p)
        page = await context.new_page()

        # ── Phase 2: PediForm — login, import, send forms ─────────────────────
        logger.info("=" * 60)
        logger.info("PHASE 2: PediForm — import schedule & send forms")
        logger.info("=" * 60)
        await pf_login(page)
        await navigate_to_todays_patients(page)

        row_count = await import_schedule(page, pediform_excel)
        logger.info(f"Schedule imported: {row_count} rows in table")

        sent_patients = await send_forms_for_all(page, patients)

        logger.info("=" * 60)
        logger.info(f"DONE — Forms sent: {len(sent_patients)} / {len(patients)} patient(s)")
        for p in sent_patients:
            logger.info(f"  ✓ {p['full_name']} → {p['form_name']}")
        logger.info("=" * 60)

        await browser.close()

    logger.info("Automation complete.")


if __name__ == "__main__":
    asyncio.run(run())
