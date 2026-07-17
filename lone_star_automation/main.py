"""
main.py
--------
Lone Star Pediatrics automation pipeline.

Ported from the reference project (ECW_automation/main.py)'s main(), with
these workflow changes:
  1. eCW schedule export applies a Facility filter (Lone Star Pediatrics
     Midlothian), inserted after the date range is set - see
     ecw/facility_filter.py.
  2. The full (unfiltered) export is imported into Patient Forms Now, and
     Well-Check eligibility is decided INSIDE Patient Forms Now from the
     imported table's own Visit Type + DOB + Appointment columns - NOT by
     reading the local Excel in Python, and NOT by assuming Visit Type
     encodes age (Lone Star's Visit Type is generic, e.g. "New patient" -
     age is computed from DOB instead). See patient_forms_now/form_sender.py's
     module docstring for the full explanation.
  3. PCareLink/ReachMyDr reminder messaging is wired in, immediately after
     a batch of forms is sent and marked form_sent - see
     pcarelink/messenger.py.

Single pass per run - cron (running this script every few hours) provides
the "check again later" behavior. There is no internal wait loop.
"""

import asyncio

import ecw.schedule_export as schedule_export
import ecw.chart_upload as chart_upload
import patient_forms_now.form_sender as form_sender
import patient_forms_now.form_downloader as form_downloader
import pcarelink.messenger as pcarelink_messenger
from config import settings
from database import state_db
from utils.logger import get_logger

log = get_logger(__name__)


async def main():
    log.info("=" * 50)
    log.info("LONE STAR PEDIATRICS - PRODUCTION PIPELINE")
    log.info("=" * 50)

    # --- STEP 0: eCW export (with Facility filter) ---
    await schedule_export.run()

    # --- STEP 1: Patient Forms Now - import full schedule, determine
    # Well-Check eligibility from the PFN table itself, and send forms ---
    sent_patients = await form_sender.run(demo_only=False)

    if sent_patients:
        for p in sent_patients:
            state_db.insert_form_sent(p)

        # --- STEP 2: PCareLink - reminder messages for patients just sent a form ---
        await pcarelink_messenger.send_messages(sent_patients)
    else:
        log.info("No new patients had a form sent this run.")

    # --- STEP 3: check ALL pending patients (from DB, not just today's export) ---
    pending = state_db.get_pending_patients()
    log.info(f"{len(pending)} patient-visits pending form completion (across all runs)")

    if pending:
        await form_downloader.run(pending)

        # --- STEP 4: upload anything downloaded (this run or a prior retry) ---
        to_upload = state_db.get_patients_needing_upload()
        if to_upload:
            uploaded_ok = await chart_upload.run(to_upload)
            for p in uploaded_ok:
                state_db.mark_completed(p["acct_no"], p["appointment_date"])
            log.info(f"{len(uploaded_ok)}/{len(to_upload)} uploaded and marked completed.")

    # --- Housekeeping: drop completed records past the retention window ---
    deleted = state_db.cleanup_old_completed(settings.STATE_RETENTION_DAYS)
    if deleted:
        log.info(f"Cleaned up {deleted} completed record(s) older than {settings.STATE_RETENTION_DAYS} days.")

    log.info("=" * 50)
    log.info("RUN COMPLETE!")
    log.info("=" * 50)


if __name__ == "__main__":
    asyncio.run(main())
