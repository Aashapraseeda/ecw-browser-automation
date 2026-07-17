"""
patient_forms_now/form_sender.py
------------------------------------
STEP 1 - Patient Forms Now: determine Well-Check eligibility INSIDE PFN
(from the imported patient table's DOB + Visit Type columns) and send forms.

ARCHITECTURE: eligibility is decided entirely from the live Patient Forms
Now table, never from the local Excel:
  1. Visit Type (column 7, "Visit type") must indicate a Well Check.
  2. DOB (column 5) and Appointment (column 6) drive the age-in-months
     calculation (utils.date_utils), matched against the supported ASQ
     brackets (config.settings.ASQ_BRACKET_TO_FORM) - NOT visit-type text,
     since Lone Star's Visit Type does not encode age (confirmed via a
     live screenshot of the real table).
  3. Chart # (column 3) is used directly as the account number - no local
     Excel lookup needed for it, DOB, name, or appointment date; all come
     straight from the table row.

The Excel is read in ONE narrow case: demo_only=True builds a safety
allowlist of test-patient account numbers (Visit Reason == "test" or
"<N> year test") so main_demo.py never touches real patients - this does
not affect which patients are eligible, only which eligible patients get
touched during testing.

Confirmed column order from a live screenshot of the real table:
  First name | Last name | Chart # | Region / Clinic | DOB | Appointment |
  Visit type | Form status | Parent portal | Action

**PARTIALLY UNVERIFIED**: the column order above is confirmed, but the
exact Visit Type text used for an actual Well Check row has not been seen
(the screenshot example was a "New patient" visit) - the match below is a
case-insensitive substring check for "well", which needs confirming
against a real Well Check row via main_demo.py. Also unconfirmed: that
role "row"/"cell" locators actually match this table's real markup. Age
bracket boundaries (utils.date_utils.match_asq_bracket) are explicit
ranges matching automation_pd_forms' age_bucket_label(), confirmed
correct by design, not live data.
"""

import re

import openpyxl
from playwright.async_api import async_playwright

from config import settings
from database import state_db
from patient_forms_now.login import pfn_login
from patient_forms_now.schedule_import import import_schedule
from utils import date_utils
from utils.logger import get_logger

log = get_logger(__name__)

_DEMO_YEAR_TEST_RE = re.compile(r"^\d+\s*year\s*test$")

# Confirmed column positions (see module docstring)
COL_FIRST_NAME = 0
COL_LAST_NAME = 1
COL_CHART_NO = 2
COL_REGION = 3
COL_DOB = 4
COL_APPOINTMENT = 5
COL_VISIT_TYPE = 6
MIN_EXPECTED_COLUMNS = 7


def is_demo_patient(visit_reason):
    """Ported from the reference project's main_1.py demo pipeline."""
    if not visit_reason:
        return False
    vr = str(visit_reason).strip().lower()
    if vr == "test":
        return True
    if _DEMO_YEAR_TEST_RE.match(vr):
        return True
    return False


def _build_demo_allowlist():
    """
    Set of account numbers identifying which patients main_demo.py is
    allowed to touch, read once from the local export. ONLY used to
    restrict main_demo.py to the intended test patient(s) - the Excel
    plays no role in deciding Well-Check eligibility itself (see module
    docstring).

    Two independent, revertible (demo-only) toggles control identification:
      - settings.DEMO_REQUIRE_VISIT_REASON (default False): the actual
        test schedule provided by the clinic has no "test" / "<N> year
        test" values in Visit Reason at all (real clinical-looking text
        like "* 3 YEAR WELL CHILD CHECK" instead) - so this check is off
        for now. Set back to True once test patients carry a real Visit
        Reason marker again.
      - settings.DEMO_RESTRICT_TO_OWN_FACILITY (default True): requires
        Appointment Facility Name to match settings.FACILITY_NAME - this
        is what actually identifies the Lone Star patient(s) in the
        current shared test schedule (which also contains Nurture Kids/
        River Ridge test patients).

    Neither toggle affects main.py (production) - this function is only
    ever called when main_demo.py runs with demo_only=True.
    """
    wb = openpyxl.load_workbook(settings.EXCEL_PATH, read_only=True)
    ws = wb.active
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    col = {name: idx for idx, name in enumerate(headers)}
    if "Patient Acct No" not in col:
        log.info("Excel missing 'Patient Acct No' column - demo allowlist will be empty")
        return set()

    require_visit_reason = settings.DEMO_REQUIRE_VISIT_REASON and "Visit Reason" in col
    if settings.DEMO_REQUIRE_VISIT_REASON and "Visit Reason" not in col:
        log.info("DEMO_REQUIRE_VISIT_REASON is on but Excel has no 'Visit Reason' column - skipping that check")

    restrict_to_own_facility = settings.DEMO_RESTRICT_TO_OWN_FACILITY and "Appointment Facility Name" in col
    if settings.DEMO_RESTRICT_TO_OWN_FACILITY and "Appointment Facility Name" not in col:
        log.info("DEMO_RESTRICT_TO_OWN_FACILITY is on but Excel has no 'Appointment Facility Name' column - skipping that check")
    own_facility_norm = settings.FACILITY_NAME.strip().lower()

    allowlist = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        acct_no = row[col["Patient Acct No"]]
        if not acct_no:
            continue
        if require_visit_reason and not is_demo_patient(row[col["Visit Reason"]]):
            continue
        if restrict_to_own_facility:
            facility_name_raw = row[col["Appointment Facility Name"]]
            facility_name_norm = str(facility_name_raw).strip().lower() if facility_name_raw else ""
            if facility_name_norm != own_facility_norm:
                log.info(f"Demo allowlist: acct {acct_no} - facility '{facility_name_raw}' != '{settings.FACILITY_NAME}' - excluding")
                continue
        acct_no_norm = str(acct_no).strip()
        if acct_no_norm not in settings.DEMO_TEST_ACCOUNT_NUMBERS:
            # Facility-exclusion alone is not a "test patient" filter - Lone
            # Star's own export already only contains Lone Star Midlothian
            # appointments, so every real patient there would otherwise pass.
            # DEMO_TEST_ACCOUNT_NUMBERS is the actual, explicit safety gate.
            continue
        allowlist.add(acct_no_norm)
    return allowlist


async def _send_form_for_open_patient(page, patient):
    """Assumes the patient's detail page is already open (View just clicked)."""
    try:
        await page.get_by_role("button", name="+ Send a form").click()
        await page.wait_for_timeout(500)
        try:
            await page.locator("label").filter(has_text=patient["form_name"]).first.click(timeout=10000)
        except Exception:
            log.info(f"Form '{patient['form_name']}' not found - skipping")
            await page.get_by_role("link", name="← Back to today's patients").click()
            await page.wait_for_load_state("networkidle")
            return False
        await page.get_by_role("button", name="Send form").click()
        log.info(f"Sent '{patient['form_name']}' successfully!")
        await page.get_by_role("link", name="← Back to today's patients").click()
        await page.wait_for_load_state("networkidle")
        return True
    except Exception as e:
        log.info(f"Error: {e}")
        try:
            await page.get_by_role("link", name="← Back to today's patients").click()
            await page.wait_for_load_state("networkidle")
        except Exception:
            pass
        return False


async def determine_and_send_from_pfn_table(page, demo_only=False):
    """
    Scans the imported "Today's Patients" table row by row, decides
    Well-Check eligibility from each row's Visit Type + DOB + Appointment
    columns, and sends the form immediately for eligible/new rows
    (clicking that row's own View link rather than re-searching).

    demo_only=True additionally restricts to test patients via the
    Excel-sourced allowlist (see module docstring).

    Returns the list of patient dicts a form was actually sent to.
    """
    demo_allowlist = _build_demo_allowlist() if demo_only else None
    if demo_only:
        log.info(f"Demo mode: restricting to {len(demo_allowlist)} test patient acct(s): {sorted(demo_allowlist)}")

    rows = page.get_by_role("row")
    row_count = await rows.count()
    log.info(f"Scanning {row_count} row(s) in the PFN patient table...")

    sent_patients = []
    for i in range(row_count):
        row = rows.nth(i)
        try:
            cells = await row.get_by_role("cell").all_inner_texts()
        except Exception:
            continue
        if len(cells) < MIN_EXPECTED_COLUMNS:
            continue  # not a data row

        first_name = cells[COL_FIRST_NAME].strip()
        last_name = cells[COL_LAST_NAME].strip()
        chart_no = cells[COL_CHART_NO].strip()
        dob_text = cells[COL_DOB].strip()
        appt_text = cells[COL_APPOINTMENT].strip()
        visit_type_text = cells[COL_VISIT_TYPE].strip()

        if first_name.lower() == "first name":
            continue  # header row

        if "well" not in visit_type_text.lower():
            continue  # not a Well Check visit type - skip

        if not chart_no:
            log.info(f"Row matched Well Check visit type ('{visit_type_text}') but Chart # is blank - skipping")
            continue

        if demo_allowlist is not None and chart_no not in demo_allowlist:
            continue

        dob = date_utils.parse_date_flexible(dob_text)
        appt_date = date_utils.parse_date_flexible(appt_text)
        if not dob or not appt_date:
            log.info(f"Chart#{chart_no}: could not parse DOB={dob_text!r} / Appointment={appt_text!r} - skipping")
            continue

        age_months = date_utils.age_in_months(dob, appt_date)
        bracket = date_utils.match_asq_bracket(age_months)
        if bracket is None:
            log.info(f"Chart#{chart_no}: age {age_months} month(s) at appointment - no matching ASQ bracket, skipping")
            continue

        form_name, form_filename = settings.ASQ_BRACKET_TO_FORM[bracket]
        appointment_date_iso = appt_date.isoformat()

        patient = {
            "acct_no": chart_no,
            "appointment_date": appointment_date_iso,
            "last_name": last_name,
            "first_name": first_name,
            "folder_name": f"{last_name} {first_name}_doc".strip(),
            "search_name": f"{last_name},{first_name}".strip(),
            "visit_type": visit_type_text,
            "form_name": form_name,
            "form_filename": form_filename,
        }

        if state_db.is_known(patient["acct_no"], patient["appointment_date"]):
            log.info(f"Chart#{chart_no} / {patient['appointment_date']} already processed - skipping resend")
            continue

        log.info(f"Row eligible: chart#{chart_no}, age {age_months}mo -> {bracket}mo ASQ '{form_name}'")
        try:
            await row.get_by_role("link", name="View").click(timeout=10000)
        except Exception:
            log.info(f"Could not click View for chart#{chart_no} - skipping")
            continue

        sent_ok = await _send_form_for_open_patient(page, patient)
        if sent_ok:
            sent_patients.append(patient)

    log.info(f"PFN table scan complete - sent forms to {len(sent_patients)} patient(s)")
    return sent_patients


async def run(demo_only=False):
    """Owns the browser session: login -> import full schedule -> determine
    eligibility from the PFN table -> send forms. Returns the patients a
    form was actually sent to (drives state_db bookkeeping + PCareLink)."""
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=300)
        context = await browser.new_context()
        page = await context.new_page()

        await pfn_login(page)
        await import_schedule(page)
        sent_patients = await determine_and_send_from_pfn_table(page, demo_only=demo_only)

        await browser.close()
    return sent_patients
