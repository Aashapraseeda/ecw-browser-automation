"""
config/settings.py
-------------------
Single source of truth for all credentials, paths, and static mappings.
Every other module reads configuration from here rather than calling
os.getenv() directly (ported convention from the reference project).
"""

import os
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# --- eCW CREDENTIALS ---
ECW_USERNAME = os.getenv("ECW_USERNAME")
ECW_PASSWORD = os.getenv("ECW_PASSWORD")
ECW_LOGIN_URL = "https://txsnmbapp.ecwcloud.com/mobiledoc/jsp/webemr/login/newLogin.jsp"
ECW_EBO_HOME_URL = "https://txsnmbebo.ecwcloud.com/bi/?perspective=home"

# --- FACILITY FILTER (Lone Star specific) ---
FACILITY_KEYWORD = os.getenv("FACILITY_KEYWORD", "lone")
FACILITY_NAME = os.getenv("FACILITY_NAME", "Lone Star Pediatrics Midlothian")

# TEMPORARY (demo-only) safety net: the current shared test schedule mixes
# Lone Star's own test patient with Nurture Kids test patients (same eCW
# tenant). main_demo.py's allowlist additionally requires each row's
# Appointment Facility Name to match FACILITY_NAME so only Lone Star's own
# test patient is ever touched - see patient_forms_now/form_sender.py's
# _build_demo_allowlist(). Does NOT affect main.py (production) either way.
# Set to False (or remove the check) once the shared test schedule no
# longer mixes clinics.
DEMO_RESTRICT_TO_OWN_FACILITY = os.getenv("DEMO_RESTRICT_TO_OWN_FACILITY", "true").strip().lower() != "false"

# TEMPORARY (demo-only): the actual test schedule provided by the clinic
# has NO "test" / "<N> year test" values in Visit Reason at all (real
# clinical-looking reasons like "* 3 YEAR WELL CHILD CHECK" instead) - so
# the Visit-Reason-based is_demo_patient() check cannot identify anything
# in this round. Set back to True once test patients are tagged with a
# real Visit Reason value again - does NOT affect main.py (production).
DEMO_REQUIRE_VISIT_REASON = os.getenv("DEMO_REQUIRE_VISIT_REASON", "false").strip().lower() == "true"

# EXPLICIT allowlist of known test-patient account numbers (Chart #) -
# the ONLY patients main_demo.py will ever touch. DEMO_RESTRICT_TO_OWN_FACILITY
# alone is NOT sufficient: Lone Star's own eCW export already applies the
# Facility filter at the report level, so every row in EXCEL_PATH already
# belongs to Lone Star Midlothian - that check is nearly a no-op and would
# match every real Lone Star patient with a valid Well-Check DOB match,
# not just the intended test one. Update this set as the actual test
# patients change; leave empty to process nothing.
DEMO_TEST_ACCOUNT_NUMBERS = {"156671"}  # Astoria Alford

# --- PATIENT FORMS NOW (PFN) CREDENTIALS ---
PFN_ORG = os.getenv("PFN_ORG")
PFN_EMAIL = os.getenv("PFN_EMAIL")
PFN_PASSWORD = os.getenv("PFN_PASSWORD")
PFN_LOGIN_URL = os.getenv("PFN_LOGIN_URL", "https://admin.lonestar.patientformsnow.com/staff/login")

# --- PCARELINK CREDENTIALS (wired into main.py and main_demo.py) ---
PCARELINK_EMAIL = os.getenv("PCARELINK_EMAIL")
PCARELINK_PASSWORD = os.getenv("PCARELINK_PASSWORD")
PCARELINK_PRACTICE = os.getenv("PCARELINK_PRACTICE")
PCARELINK_MESSAGE = os.getenv("PCARELINK_MESSAGE")

# --- PATHS ---
EXCEL_PATH = os.getenv("EXCEL_PATH", os.path.join(BASE_DIR, "ecw_schedule.xlsx"))
DOC_FOLDER = os.getenv("ECW_PATIENTS_DOC_FOLDER", os.path.join(BASE_DIR, "patients_doc"))
STATE_DB_PATH = os.getenv("STATE_DB_PATH", os.path.join(BASE_DIR, "data", "patients_state.db"))
LOG_DIR = os.getenv("LOG_DIR", os.path.join(BASE_DIR, "logs"))

# --- SETTINGS ---
STATE_RETENTION_DAYS = int(os.getenv("STATE_RETENTION_DAYS", "30"))

# --- ASQ AGE BRACKETS (DOB-based) ---
# Lone Star's Patient Forms Now table does NOT encode age in the Visit
# Type text (unlike the reference project's "9 MONTH WC" style values -
# confirmed via a live screenshot showing generic values like "New
# patient"). Eligibility/form selection is computed from DOB + Appointment
# date instead - see utils/date_utils.match_asq_bracket(), which maps age
# in months to one of these via explicit ranges (matching the boundary
# logic of automation_pd_forms/utils/date_utils.py's age_bucket_label()),
# not an arbitrary +/- tolerance.
ASQ_AGE_BRACKETS_MONTHS = [9, 12, 15, 18, 24, 30, 36, 48]

# bracket (months) -> (form_name shown in PFN's "+ Send a form" checkboxes,
# filename-safe form label). Label TEXT corrected from a live screenshot
# of Lone Star's own PFN checkbox list - it uses a DIFFERENT naming
# convention than the reference clinic's Pediforms account (mostly
# hyphenated, e.g. "ASQ-36 Months", except 48-month which is "ASQ 48
# Months" with a space, no hyphen). The 15-month bracket intentionally
# reuses the 18-month form (there's no separate ASQ-15-months checkbox).
ASQ_BRACKET_TO_FORM = {
    9: ("ASQ-9 Months", "ASQ_9_Months"),
    12: ("ASQ-12 Months", "ASQ_12_Months"),
    15: ("ASQ-18 Months", "ASQ_18_Months"),
    18: ("ASQ-18 Months", "ASQ_18_Months"),
    24: ("ASQ-24 Months", "ASQ_24_Months"),
    30: ("ASQ-30 Months", "ASQ_30_Months"),
    36: ("ASQ-36 Months", "ASQ_36_Months"),
    48: ("ASQ 48 Months", "ASQ_48_Months"),
}
