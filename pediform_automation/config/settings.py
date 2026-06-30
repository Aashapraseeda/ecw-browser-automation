"""
config/settings.py
------------------
Central configuration loaded from .env.
All other modules import from here — never read os.environ directly.
"""

import os
from dotenv import load_dotenv

load_dotenv()

# ── PediForm credentials ───────────────────────────────────────────────────────
PEDIFORM_URL      = os.getenv("PEDIFORM_URL", "https://admin.pediformpro.com/staff/login")
PEDIFORM_ORG      = os.getenv("PEDIFORM_ORG", "")
PEDIFORM_EMAIL    = os.getenv("PEDIFORM_EMAIL", "")
PEDIFORM_PASSWORD = os.getenv("PEDIFORM_PASSWORD", "")

# ── ECW credentials ────────────────────────────────────────────────────────────
ECW_URL      = os.getenv("ECW_URL", "")
ECW_USERNAME = os.getenv("ECW_USERNAME", "")
ECW_PASSWORD = os.getenv("ECW_PASSWORD", "")

# ── ReachMyDr credentials (placeholder) ───────────────────────────────────────
REACHMYDR_URL      = os.getenv("REACHMYDR_URL", "")
REACHMYDR_USERNAME = os.getenv("REACHMYDR_USERNAME", "")
REACHMYDR_PASSWORD = os.getenv("REACHMYDR_PASSWORD", "")

# ── File paths ─────────────────────────────────────────────────────────────────
SCHEDULE_EXCEL_PATH = os.getenv(
    "SCHEDULE_EXCEL_PATH",
    r"C:\Users\prase\Downloads\ecw_patient_automation\daily_schedule.xlsx",
)
PEDIFORM_IMPORT_PATH = os.getenv(
    "PEDIFORM_IMPORT_PATH",
    r"C:\Users\prase\Downloads\ecw_patient_automation\pediform_import.xlsx",
)
COMPLETED_FORMS_DOWNLOAD_DIR = os.getenv(
    "COMPLETED_FORMS_DOWNLOAD_DIR",
    r"C:\Users\prase\Downloads\ecw_patient_automation\completed_forms",
)
FORM_LINK = os.getenv("FORM_LINK", "")

# ── Automation settings ────────────────────────────────────────────────────────
APPOINTMENT_WINDOW_DAYS = int(os.getenv("APPOINTMENT_WINDOW_DAYS", "3"))
HEADLESS = os.getenv("HEADLESS", "false").lower() == "true"
SLOW_MO  = int(os.getenv("SLOW_MO", "200"))

# ── Visit statuses to SKIP ────────────────────────────────────────────────────
# ECW format: "CAN : Cancelled", "NCNS : No Show", "LCN : Late Cancel"
SKIP_STATUSES = {"CAN", "NCNS", "LCN", "DNA"}

# ── Age bucket → PediForm form name mapping ────────────────────────────────────
# Values must exactly match PediForm's label text in the Send Form panel.
# Confirmed labels (2026-06-30 inspect):
#   ASQ9Mos, ASQ 12 Months, ASQ 18 Months, ASQ 24 Months,
#   ASQ30, ASQ 36 Months, ASQ 48 Months, EPDS, LEAD, TB
# None = no form sent for that age bucket.
AGE_FORM_MAP: dict = {
    "NEWBORN":    "EPDS",
    "1 WEEK":     "EPDS",
    "2 MONTH":    "EPDS",
    "4 MONTH":    None,
    "6 MONTH":    "LEAD",
    "9 MONTH":    "ASQ9Mos",
    "12 MONTH":   "TB",
    "15 MONTH":   "TB",
    "18 MONTH":   "ASQ 18 Months",
    "24 MONTH":   "ASQ 24 Months",
    "30 MONTH":   "ASQ30",
    "3 YEAR":     "ASQ 36 Months",
    "4 YEAR":     "ASQ 48 Months",
    "5-6 YEAR":   "TB",
    "7-11 YEAR":  "TB",
    "12-18 YEAR": "PHQ-9",
}
