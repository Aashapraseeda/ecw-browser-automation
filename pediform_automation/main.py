"""
main.py
-------
Full automation orchestrator.

Phase 1: Read ECW export → build PediForm import Excel
Phase 2: PediForm — import schedule → send age-appropriate forms (Send button DISABLED)
Phase 3: ReachMyDr — send reminders (placeholder, skipped)
Phase 4: PediForm — poll for completed forms → download PDFs
Phase 5: ECW — upload PDFs to Chart Documents (SKIPPED — implement later)
"""

import asyncio
from playwright.async_api import async_playwright

from config.settings import SCHEDULE_EXCEL_PATH, PEDIFORM_IMPORT_PATH
from utils.excel_reader import get_eligible_patients
from utils.ecw_to_pediform import build_pediform_excel
from utils.logger import get_logger

from modules.pediform.login import login as pf_login, navigate_to_todays_patients, new_pediform_context
from modules.pediform.schedule_import import import_schedule
from modules.pediform.form_sender import send_forms_for_all
from modules.pediform.form_downloader import download_completed_forms
from modules.reachmydr.messenger import ReachMyDrMessenger

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

    # ── Phase 1b: Build PediForm import Excel ──────────────────────────────────
    logger.info("Building PediForm import Excel (dates as text)...")
    pediform_excel = build_pediform_excel(patients, PEDIFORM_IMPORT_PATH)

    async with async_playwright() as p:
        browser, context = await new_pediform_context(p)
        page = await context.new_page()

        # ── Phase 2: PediForm ─────────────────────────────────────────────────
        logger.info("=" * 60)
        logger.info("PHASE 2: PediForm — import & send forms")
        logger.info("=" * 60)
        await pf_login(page)
        await navigate_to_todays_patients(page)
        row_count = await import_schedule(page, pediform_excel)
        logger.info(f"Import complete: {row_count} rows in table")

        sent_patients = await send_forms_for_all(page, patients)
        logger.info(f"Forms processed: {len(sent_patients)} patient(s)")

        # ── Phase 3: ReachMyDr reminders (placeholder) ────────────────────────
        logger.info("=" * 60)
        logger.info("PHASE 3: ReachMyDr reminders (placeholder)")
        logger.info("=" * 60)
        messenger = ReachMyDrMessenger()
        await messenger.send_reminders(sent_patients)

        # ── Phase 4: Download completed PDFs ──────────────────────────────────
        logger.info("=" * 60)
        logger.info("PHASE 4: Downloading completed form PDFs")
        logger.info("=" * 60)
        downloaded = await download_completed_forms(page, sent_patients)
        logger.info(f"PDFs downloaded: {len(downloaded)}")

        # ── Phase 5: ECW upload (SKIPPED) ─────────────────────────────────────
        logger.info("=" * 60)
        logger.info("PHASE 5: ECW upload — SKIPPED (not yet implemented)")
        logger.info("=" * 60)

        await browser.close()

    logger.info("Automation complete.")


if __name__ == "__main__":
    asyncio.run(run())
