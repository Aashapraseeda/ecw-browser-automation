"""
pediform_import_test.py
-----------------------
Step-by-step test:
  1. Read the real ECW Encounter Patient Download Excel
  2. Filter to eligible patients (next N days, age-matched forms)
  3. Build a clean PediForm import Excel (dates as text strings)
  4. Log into PediForm
  5. Upload the PediForm Excel and verify the table populates

All credentials and paths are read from .env — nothing is hardcoded here.

Run:
    python pediform_import_test.py
"""

import asyncio
import os
from playwright.async_api import async_playwright

import config.settings as cfg
from utils.excel_reader import get_eligible_patients
from utils.ecw_to_pediform import build_pediform_excel


async def main():
    # ── Step 1: Parse ECW Excel ───────────────────────────────────────────────
    print("\n" + "="*60)
    print("STEP 1: Reading ECW export and filtering patients")
    print("="*60)
    patients = get_eligible_patients(cfg.SCHEDULE_EXCEL_PATH)

    if not patients:
        print("\nNo eligible patients found in the schedule.")
        print("Check that the appointment dates are within the next "
              f"{cfg.APPOINTMENT_WINDOW_DAYS} days and that Visit Status is not cancelled/no-show.")
        return

    print(f"\nEligible patients ({len(patients)}):")
    for p in patients:
        print(f"  {p['full_name']} | DOB: {p['dob']} | "
              f"Appt: {p['appt_date']} | Bucket: {p['age_bucket']} | Form: {p['form_name']}")

    # ── Step 2: Build PediForm Excel ──────────────────────────────────────────
    print("\n" + "="*60)
    print("STEP 2: Building PediForm import Excel")
    print("="*60)
    os.makedirs(os.path.dirname(cfg.PEDIFORM_IMPORT_PATH), exist_ok=True)
    pediform_excel = build_pediform_excel(patients, cfg.PEDIFORM_IMPORT_PATH)
    print(f"Saved: {pediform_excel}")

    # ── Step 3: Upload to PediForm ────────────────────────────────────────────
    print("\n" + "="*60)
    print("STEP 3: Uploading to PediForm")
    print("="*60)

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=cfg.HEADLESS,
            slow_mo=cfg.SLOW_MO,
            args=["--disable-blink-features=AutomationControlled"],
        )
        context = await browser.new_context(
            ignore_https_errors=True,
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
        )
        page = await context.new_page()

        print("Logging in...")
        await page.goto(cfg.PEDIFORM_URL, wait_until="commit", timeout=90000)
        await page.locator("#admin-practice").wait_for(state="visible", timeout=90000)
        await page.locator("#admin-practice").fill(cfg.PEDIFORM_ORG)
        await page.locator("#admin-email").fill(cfg.PEDIFORM_EMAIL)
        await page.locator("#admin-password").fill(cfg.PEDIFORM_PASSWORD)
        await page.locator("button.patient-portal-submit").click()
        await page.get_by_role("link", name="Today's Patients", exact=True).wait_for(
            state="visible", timeout=90000
        )
        print("Logged in!")

        await page.get_by_role("link", name="Today's Patients", exact=True).click()
        await asyncio.sleep(2)

        print("Uploading PediForm Excel...")
        file_input = page.locator("input[type='file']")
        await file_input.wait_for(timeout=15000)
        await file_input.set_input_files(pediform_excel)

        await page.get_by_role("button", name="Import schedule").click()
        print("Import clicked — waiting for result...")
        await asyncio.sleep(4)

        body_text = await page.inner_text("body")
        if "inserted" in body_text.lower():
            for line in body_text.splitlines():
                if "inserted" in line.lower() or "upload" in line.lower():
                    print(f"\nPediForm response: {line.strip()}")

        try:
            await page.wait_for_selector("table tbody tr", timeout=10000)
            row_count = await page.locator("table tbody tr").count()
            print(f"\nSUCCESS — {row_count} patient row(s) in table after import.")
        except Exception:
            print("\nTable did not populate — check PediForm response above.")

        print("\nBrowser staying open for inspection. Press Ctrl+C to exit.")
        await asyncio.sleep(99999)


asyncio.run(main())
