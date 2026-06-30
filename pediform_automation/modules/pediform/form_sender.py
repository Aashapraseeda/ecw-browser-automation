"""
modules/pediform/form_sender.py
--------------------------------
For each eligible patient in the Today's Patients table:
  1. Click the patient's View link (href = /staff/patients/{uuid})
  2. Click '+ Send a form' to open the form panel
  3. Check the age-appropriate form checkbox (matched by label text)
  4. [DISABLED] Click 'Send form' — uncomment when ready to go live
  5. Click 'Cancel' to close the panel
  6. Go back to Today's Patients and continue

Confirmed selectors (from live inspect 2026-06-30):
  - '+ Send a form' button: get_by_role("button", name="+ Send a form")
  - Form checkboxes: get_by_label(<form_name>) — labels matched by text
  - 'Send form' submit: button.btn
  - 'Cancel' button:  get_by_role("button", name="Cancel")
"""

import asyncio
from typing import Dict, List

from playwright.async_api import Page

from utils.logger import get_logger

logger = get_logger(__name__)

PEDIFORM_BASE = "https://admin.pediformpro.com"


# ── Helpers ───────────────────────────────────────────────────────────────────

async def _open_send_form_panel(page: Page) -> bool:
    """Click '+ Send a form'. Returns True if the panel appeared."""
    try:
        btn = page.get_by_role("button", name="+ Send a form")
        await btn.wait_for(state="visible", timeout=10000)
        await btn.click()
        await asyncio.sleep(1)
        return True
    except Exception as e:
        logger.error(f"Could not open Send Form panel: {e}")
        return False


async def _select_form_checkbox(page: Page, form_name: str) -> bool:
    """
    Find the checkbox associated with the given form label and check it.
    Returns True if the checkbox was found and checked.
    """
    try:
        checkbox = page.get_by_label(form_name, exact=True)
        await checkbox.wait_for(state="visible", timeout=10000)
        if not await checkbox.is_checked():
            await checkbox.check()
        logger.debug(f"Checkbox checked: '{form_name}'")
        return True
    except Exception as e:
        logger.error(f"Could not find/check checkbox for form '{form_name}': {e}")
        return False


async def _cancel_panel(page: Page) -> None:
    """Close the Send Form panel without sending."""
    try:
        cancel = page.get_by_role("button", name="Cancel")
        await cancel.wait_for(state="visible", timeout=5000)
        await cancel.click()
        await asyncio.sleep(1)
    except Exception:
        pass


async def _submit_send_form(page: Page) -> None:
    """Click the final 'Send form' button (the primary submit in the panel)."""
    # Use button text to avoid matching the Cancel button
    submit = page.get_by_role("button", name="Send form", exact=True)
    await submit.wait_for(state="visible", timeout=10000)
    logger.info("  >> About to click 'Send form' — pausing 3s so you can verify...")
    await asyncio.sleep(3)
    await submit.click()
    logger.info("  >> 'Send form' clicked.")
    await asyncio.sleep(2)   # wait for confirmation / panel to close


# ── Per-patient function ──────────────────────────────────────────────────────

async def send_form_for_patient(
    page: Page,
    patient_href: str,
    form_name: str,
    patient_name: str,
) -> bool:
    """
    Navigate to one patient's page, open the Send Form panel,
    select the correct form checkbox, and close WITHOUT sending.

    Args:
        page:          Playwright page (logged into PediForm).
        patient_href:  URL path like /staff/patients/{uuid}.
        form_name:     Exact label text in PediForm (e.g. 'TB', 'ASQ 36 Months').
        patient_name:  Display name for logging only.

    Returns:
        True on success, False on any error.
    """
    logger.info(f"[Form] {patient_name} → '{form_name}'")
    try:
        await page.goto(PEDIFORM_BASE + patient_href, wait_until="commit", timeout=30000)
        await asyncio.sleep(2)

        if not await _open_send_form_panel(page):
            return False

        if not await _select_form_checkbox(page, form_name):
            await _cancel_panel(page)
            return False

        await _submit_send_form(page)
        logger.info(f"[Form] {patient_name}: '{form_name}' sent successfully")
        return True

    except Exception as e:
        logger.error(f"[Form] {patient_name}: unexpected error — {e}")
        return False


# ── Batch function ────────────────────────────────────────────────────────────

def _build_lookup(patients: List[Dict]) -> Dict[str, Dict]:
    """
    Build a name→patient lookup that handles both "First Last" and "Last, First"
    display formats, since PediForm shows names as "Last, First".
    """
    lookup: Dict[str, Dict] = {}
    for p in patients:
        first = p["first_name"].strip()
        last  = p["last_name"].strip()
        # "First Last"
        lookup[f"{first} {last}".lower()] = p
        # "Last, First" (PediForm's display format)
        lookup[f"{last}, {first}".lower()] = p
        # "Last First" (no comma, just in case)
        lookup[f"{last} {first}".lower()] = p
    return lookup


async def _extract_table_entries(page: Page) -> List[tuple]:
    """
    Extract all (first_name, last_name, href) triples from the patient table
    in a single JavaScript call.

    PediForm table columns (confirmed 2026-06-30):
      td[0] = First name
      td[1] = Last name
      td[2] = DOB or other field
    """
    entries = await page.evaluate("""
        () => {
            const rows = document.querySelectorAll('table tbody tr');
            return Array.from(rows).map(row => {
                const tds = row.querySelectorAll('td');
                const first = tds[0]?.innerText?.trim() || '';
                const last  = tds[1]?.innerText?.trim() || '';
                const link  = row.querySelector('a[href*="/staff/patients/"]');
                const href  = link?.getAttribute('href') || '';
                return [first, last, href];
            }).filter(([f, l, h]) => f && h);
        }
    """)
    return [tuple(e) for e in entries]


async def send_forms_for_all(page: Page, patients: List[Dict]) -> List[Dict]:
    """
    Process forms for every eligible patient.

    Uses a single JS call to extract all (name, href) pairs from the
    table, matches against our eligible patients (handling "Last, First"
    format), then visits each patient page to select the form.

    Args:
        page:     Playwright page logged into PediForm, on Today's Patients.
        patients: Output of excel_reader.get_eligible_patients().

    Returns:
        List of patient dicts successfully processed.
    """
    lookup = _build_lookup(patients)
    processed: set = set()
    succeeded: List[Dict] = []

    await page.get_by_role("link", name="Today's Patients", exact=True).click()
    await asyncio.sleep(2)

    # Extract ALL table entries in one fast JS call
    logger.info("Extracting patient table via JavaScript (fast)...")
    table_entries = await _extract_table_entries(page)
    logger.info(
        f"Table has {len(table_entries)} row(s). "
        f"Eligible patients to match: {len(patients)}"
    )

    # Match our patients to table rows
    # table_entries = [(first, last, href), ...]
    matched: List[tuple] = []   # (display_name, href, patient_dict)
    for first_pd, last_pd, href in table_entries:
        # Build lookup keys from what PediForm shows
        f = first_pd.strip().lower()
        l = last_pd.strip().lower()
        # Try "First Last", "Last, First", "Last First", and first-name-only
        for key in (f"{f} {l}", f"{l}, {f}", f"{l} {f}", f):
            patient = lookup.get(key)
            if patient:
                display = f"{first_pd} {last_pd}".strip()
                if key not in processed:
                    matched.append((display, href, patient))
                    processed.add(key)
                    logger.info(
                        f"Matched: '{display}' (PediForm) "
                        f"→ '{patient['full_name']}' (Excel) "
                        f"→ form '{patient['form_name']}'"
                    )
                break

    if not matched:
        sample = [
            f"'{f} {l}'" for f, l, _ in table_entries[:20]
        ]
        logger.warning(
            "No eligible patients found in the table.\n"
            f"  Expected (from Excel): {[p['full_name'] for p in patients]}\n"
            "  Sample first+last from PediForm table (first 20):\n"
            + "\n".join(f"    {s}" for s in sample)
        )
        return []

    logger.info(f"Matched {len(matched)} / {len(patients)} patients in table.")

    # Process each matched patient
    processed_names: set = set()
    for display_name, href, patient in matched:
        if display_name.lower() in processed_names:
            continue

        success = await send_form_for_patient(
            page,
            patient_href=href,
            form_name=patient["form_name"],
            patient_name=display_name,
        )

        if success:
            succeeded.append(patient)
            processed_names.add(display_name.lower())

        # Return to Today's Patients for next patient
        await page.get_by_role("link", name="Today's Patients", exact=True).click()
        await asyncio.sleep(2)

    logger.info(f"Forms processed: {len(succeeded)} / {len(patients)}")
    return succeeded
