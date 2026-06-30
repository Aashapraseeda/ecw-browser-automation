"""
utils/excel_reader.py
---------------------
Reads the ECW "Encounter Patient Download" Excel export and returns a list of
eligible patient dicts for the next N days.

ECW export column mapping (confirmed from real file):
  A  = Appointment Date      (datetime)
  F  = Appointment Start Time (time)
  L  = Visit Type            (e.g. "12 MONTHWC : 12 MONTH WC")
  N  = Visit Status          (e.g. "PEN : Pending", "CAN : Cancelled")
  U  = Patient First Name
  V  = Patient Last Name
  Y  = Patient DOB           (datetime)
"""

from datetime import datetime, date
from typing import List, Dict, Optional

import openpyxl

from config.settings import (
    SCHEDULE_EXCEL_PATH,
    APPOINTMENT_WINDOW_DAYS,
    AGE_FORM_MAP,
    SKIP_STATUSES,
)
from utils.date_utils import appointment_within_days, age_bucket_label, to_date
from utils.logger import get_logger

logger = get_logger(__name__)

# ── Column indices (0-based) ───────────────────────────────────────────────────
COL_APPT_DATE   = 0   # A
COL_APPT_TIME   = 5   # F
COL_VISIT_TYPE  = 11  # L
COL_STATUS      = 13  # N
COL_FIRST_NAME  = 20  # U
COL_LAST_NAME   = 21  # V
COL_DOB         = 24  # Y


def _parse_status_code(raw: Optional[str]) -> str:
    """Extract the short code before ' : ' from ECW status strings."""
    if not raw:
        return ""
    return str(raw).split(":")[0].strip().upper()


def _parse_visit_type(raw: Optional[str]) -> str:
    """Return the human-readable part after ' : ', e.g. '12 MONTH WC'."""
    if not raw:
        return ""
    parts = str(raw).split(":", 1)
    return parts[1].strip() if len(parts) > 1 else str(raw).strip()


def read_schedule(excel_path: str = SCHEDULE_EXCEL_PATH) -> List[Dict]:
    """
    Load all rows from the ECW export.
    Returns raw list of dicts (no filtering applied yet).
    """
    logger.info(f"Reading ECW schedule: {excel_path}")
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active

    patients = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        # Skip completely empty rows
        if not any(row):
            continue

        first_name  = str(row[COL_FIRST_NAME] or "").strip()
        last_name   = str(row[COL_LAST_NAME]  or "").strip()
        raw_dob     = row[COL_DOB]
        raw_appt    = row[COL_APPT_DATE]
        raw_time    = row[COL_APPT_TIME]
        raw_status  = row[COL_STATUS]
        raw_vtype   = row[COL_VISIT_TYPE]

        if not first_name or not last_name:
            continue

        dob       = to_date(raw_dob)
        appt_date = to_date(raw_appt)

        patients.append({
            "first_name":  first_name,
            "last_name":   last_name,
            "full_name":   f"{first_name} {last_name}",
            "dob":         dob,
            "appt_date":  appt_date,
            "appt_time":  raw_time,          # datetime.time or None
            "visit_type": _parse_visit_type(raw_vtype),
            "status_code": _parse_status_code(raw_status),
            "row":         row_idx,
        })

    logger.info(f"Total rows read: {len(patients)}")
    return patients


def get_eligible_patients(excel_path: str = SCHEDULE_EXCEL_PATH) -> List[Dict]:
    """
    Filter to patients whose:
      - appointment is within APPOINTMENT_WINDOW_DAYS from today
      - status is NOT in SKIP_STATUSES (cancelled, no-show, etc.)
      - DOB is parseable
      - age maps to a form in AGE_FORM_MAP (None entries are skipped)

    Returns list of dicts with added keys: age_bucket, form_name.
    """
    all_patients = read_schedule(excel_path)
    eligible = []

    for p in all_patients:
        # 1. Skip cancelled / no-show
        if p["status_code"] in SKIP_STATUSES:
            logger.debug(f"Skipping {p['full_name']}: status={p['status_code']}")
            continue

        # 2. Skip if appointment date unknown or outside window
        if p["appt_date"] is None:
            logger.debug(f"Skipping {p['full_name']}: no appointment date")
            continue
        if not appointment_within_days(p["appt_date"], APPOINTMENT_WINDOW_DAYS):
            logger.debug(
                f"Skipping {p['full_name']}: appt {p['appt_date']} "
                f"outside {APPOINTMENT_WINDOW_DAYS}-day window"
            )
            continue

        # 3. Skip if DOB missing
        if p["dob"] is None:
            logger.warning(f"Skipping {p['full_name']}: could not parse DOB")
            continue

        # 4. Determine age bucket and form
        bucket = age_bucket_label(p["dob"])
        if bucket is None:
            logger.info(f"Skipping {p['full_name']}: age out of well-child range")
            continue

        form_name = AGE_FORM_MAP.get(bucket)
        if form_name is None:
            logger.info(
                f"Skipping {p['full_name']}: age bucket '{bucket}' has no form assigned"
            )
            continue

        p["age_bucket"] = bucket
        p["form_name"]  = form_name
        eligible.append(p)
        logger.info(
            f"Eligible: {p['full_name']} | DOB {p['dob']} | "
            f"Appt {p['appt_date']} | Bucket: {bucket} | Form: {form_name}"
        )

    logger.info(f"Eligible patients: {len(eligible)} / {len(all_patients)}")
    return eligible
