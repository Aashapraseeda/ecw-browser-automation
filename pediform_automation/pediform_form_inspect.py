"""
pediform_form_inspect.py
------------------------
Opens the FIRST patient's View page after importing the schedule,
then prints all buttons and inputs on that page so we can confirm
the correct selectors for the form-send panel.

Does NOT send or submit anything.
All credentials and paths are read from .env.

Run:
    python pediform_form_inspect.py
"""

import asyncio
import os
from playwright.async_api import async_playwright

import config.settings as cfg
from utils.excel_reader import get_eligible_patients
from utils.ecw_to_pediform import build_pediform_excel


async def main():
    patients = get_eligible_patients(cfg.SCHEDULE_EXCEL_PATH)
    if not patients:
        print("No eligible patients — check your Excel file.")
        return

    os.makedirs(os.path.dirname(cfg.PEDIFORM_IMPORT_PATH), exist_ok=True)
    pediform_excel = build_pediform_excel(patients, cfg.PEDIFORM_IMPORT_PATH)

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

        # ── Login ──────────────────────────────────────────────────────────────
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

        # ── Import schedule ────────────────────────────────────────────────────
        await page.get_by_role("link", name="Today's Patients", exact=True).click()
        await asyncio.sleep(2)
        file_input = page.locator("input[type='file']")
        await file_input.wait_for(timeout=15000)
        await file_input.set_input_files(pediform_excel)
        await page.get_by_role("button", name="Import schedule").click()
        await asyncio.sleep(4)

        # ── Click the first patient's View button ──────────────────────────────
        print("\nLooking for first View button...")
        first_view = page.locator("table tbody tr").first.locator("a")
        await first_view.wait_for(state="visible", timeout=15000)
        patient_href = await first_view.get_attribute("href")
        print(f"Navigating to patient: {patient_href}")
        await first_view.click()
        await asyncio.sleep(3)

        # ── Inspect the patient page ───────────────────────────────────────────
        print("\n--- BUTTONS ON PATIENT PAGE ---")
        buttons = await page.locator("button").all()
        for btn in buttons:
            id_   = await btn.get_attribute("id") or ""
            cls   = await btn.get_attribute("class") or ""
            text  = (await btn.inner_text()).strip().replace("\n", " ")[:80]
            print(f"  <button> id='{id_}'  class='{cls[:60]}'  text='{text}'")

        print("\n--- INPUTS ON PATIENT PAGE ---")
        inputs = await page.locator("input").all()
        for el in inputs:
            id_   = await el.get_attribute("id") or ""
            name  = await el.get_attribute("name") or ""
            type_ = await el.get_attribute("type") or ""
            cls   = await el.get_attribute("class") or ""
            label = await el.get_attribute("aria-label") or ""
            print(f"  <input> id='{id_}'  name='{name}'  type='{type_}'  class='{cls[:50]}'  aria-label='{label}'")

        print("\n--- LINKS ON PATIENT PAGE ---")
        links = await page.locator("a").all()
        for a in links:
            href = await a.get_attribute("href") or ""
            text = (await a.inner_text()).strip().replace("\n", " ")[:60]
            print(f"  <a> href='{href}'  text='{text}'")

        print("\nDone! Browser staying open.")
        await asyncio.sleep(99999)


asyncio.run(main())
