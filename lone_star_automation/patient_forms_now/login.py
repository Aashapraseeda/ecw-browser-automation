"""
patient_forms_now/login.py
----------------------------
Shared Patient Forms Now (PFN) login.

Extracted from the reference project, where this exact block was
duplicated inside pediforms_send_forms() and pediforms_check_and_download().

Verified live (read-only page inspection, no credentials submitted) against
https://admin.lonestar.patientformsnow.com/staff/login:
  - That URL goes straight to the real login form - no landing-page
    click-through needed despite the marketing page shown at the bare
    root domain.
  - The Organization field has NO accessible label/aria-label, only a
    placeholder ("e.g. Lone Star Pediatrics") - get_by_role(name=...)
    would not match it, unlike the reference project's PediformPro page
    which has a real "Organization name" label. Uses get_by_placeholder
    instead.
  - Email, Password, and the "Sign in" button are unchanged.
"""

import asyncio

from config import settings
from utils.logger import get_logger

log = get_logger(__name__)

ORG_FIELD_PLACEHOLDER = "e.g. Lone Star Pediatrics"


async def pfn_login(page):
    log.info("Logging into Patient Forms Now...")
    await page.goto(settings.PFN_LOGIN_URL, timeout=60000, wait_until="domcontentloaded")
    await asyncio.sleep(3)
    await page.get_by_placeholder(ORG_FIELD_PLACEHOLDER).fill(settings.PFN_ORG)
    await page.get_by_role("textbox", name="Email").fill(settings.PFN_EMAIL)
    await page.get_by_role("textbox", name="Password").fill(settings.PFN_PASSWORD)
    await page.get_by_role("button", name="Sign in").click()
    await page.wait_for_load_state("networkidle")
    log.info("Logged in!")
