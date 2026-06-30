"""
pediform_send_form_test.py
--------------------------
Tests the full form-sending flow for all eligible patients.
All credentials and paths are read from .env — nothing is hardcoded here.

Run:
    python pediform_send_form_test.py
"""

import asyncio
import os
from playwright.async_api import async_playwright

import config.settings as cfg
from utils.excel_reader import get_eligible_patients
from utils.ecw_to_pediform import build_pediform_excel
from modules.pediform.form_sender import send_forms_for_all


async def main():
    # Step 1: Get eligible patients
    print("\n" + "="*60)
    print("STEP 1: Reading ECW export")
    print("="*60)
    patients = get_eligible_patients(cfg.SCHEDULE_EXCEL_PATH)
    if not patients:
        print("No eligible patients found.")
        return

    for p in patients:
        print(f"  {p['full_name']} | {p['age_bucket']} | Form: {p['form_name']}")

    # Step 2: Build PediForm Excel
    os.makedirs(os.path.dirname(cfg.PEDIFORM_IMPORT_PATH), exist_ok=True)
    pediform_excel = build_pediform_excel(patients, cfg.PEDIFORM_IMPORT_PATH)
    print(f"\nPediForm Excel: {pediform_excel}")

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

        # Step 3: Login
        print("\n" + "="*60)
        print("STEP 3: Login + Import")
        print("="*60)
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

        # Import schedule
        await page.get_by_role("link", name="Today's Patients", exact=True).click()
        await asyncio.sleep(2)
        file_input = page.locator("input[type='file']")
        await file_input.wait_for(timeout=15000)
        await file_input.set_input_files(pediform_excel)
        await page.get_by_role("button", name="Import schedule").click()
        await asyncio.sleep(4)
        print("Schedule imported!")

        # Step 4: Send forms
        print("\n" + "="*60)
        print("STEP 4: Processing forms (SEND button ENABLED — forms will be sent)")
        print("="*60)
        sent = await send_forms_for_all(page, patients)

        print("\n" + "="*60)
        print(f"DONE — Forms processed for {len(sent)} / {len(patients)} patient(s)")
        for p in sent:
            print(f"  ✓ {p['full_name']} → {p['form_name']}")
        print("="*60)
        print("\nBrowser staying open. Press Ctrl+C to exit.")
        await asyncio.sleep(99999)


asyncio.run(main())
