"""
main.py
--------
Lone Star Pediatrics automation pipeline.

Ported from the reference project (ECW_automation/main.py)'s main(), with
these workflow changes:
  1. eCW schedule export applies a Facility filter (Lone Star Pediatrics
     Midlothian), inserted after the date range is set - see
     ecw/facility_filter.py.
  2. **(2026-07-21) Well-Check eligibility is decided from the eCW Excel
     export itself** (patient_forms_now.form_sender.read_eligible_patients_from_excel),
     the same way the reference clinic does it - NOT from Patient Forms
     Now's own table. A live production debug run proved PFN's "Visit
     type" column is a generic patient-status field ("New patient" /
     "Follow-up" / "Sick visit"), never the clinical visit type, which is
     why the original PFN-table approach found 0 eligible patients every
     run. The full (unfiltered) Excel is still imported into PFN exactly
     as before - only the source driving which patients get searched and
     sent a form changed. See patient_forms_now/form_sender.py's module
     docstring for the full explanation.
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
    # (2026-07-21) tomorrow through +3 days - was a 7-day window starting today.
    await schedule_export.run(window_days=2, start_offset_days=1)

    # --- STEP 1a: determine Well-Check eligibility from the Excel itself
    # (Visit Type/Visit Reason + DOB-based age, 9-48 months inclusive) ---
    exported_patients = form_sender.read_eligible_patients_from_excel(settings.EXCEL_PATH)
    log.info(f"Found {len(exported_patients)} ASQ-eligible patients in this export")

    new_patients = [
        p for p in exported_patients
        if not state_db.is_known(p["acct_no"], p["appointment_date"])
    ]
    already_known = len(exported_patients) - len(new_patients)
    log.info(f"{len(new_patients)} new patient-visits, {already_known} already processed (skipping resend)")

    sent_patients = []
    if new_patients:
        # --- STEP 1b: Patient Forms Now - import full schedule, then
        # search + send only for the patients determined eligible above ---
        sent_patients = await form_sender.run_from_excel_list(new_patients)
        for p in sent_patients:
            state_db.insert_form_sent(p)

        # --- STEP 2: PCareLink - reminder messages for patients just sent a form ---
        if sent_patients:
            await pcarelink_messenger.send_messages(sent_patients)
    else:
        log.info("No new patients to send forms to this run.")

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
