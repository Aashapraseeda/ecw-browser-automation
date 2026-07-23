"""
ecw/chart_upload.py
----------------------
STEP 4 - eCW: upload completed forms into the patient's Chart Documents.

Ported unchanged from the reference project's ecw_upload_forms() and
_go_to_search(), using the shared ecw_login() helper.
"""

import asyncio
import glob
import json
import os

from playwright.async_api import async_playwright

from config import settings
from ecw.login import ecw_login
from utils.logger import get_logger

log = get_logger(__name__)


async def _go_to_search(page):
    await page.keyboard.press("Escape")
    await asyncio.sleep(1)
    await page.keyboard.press("Escape")
    await asyncio.sleep(1)
    try:
        await page.wait_for_selector("#patient-hubBtn1", timeout=10000)
        await page.locator("#patient-hubBtn1").click()
        await asyncio.sleep(1)
    except Exception:
        pass
    await page.wait_for_selector("#jellybean-panelLink65", timeout=30000)
    await page.locator("#jellybean-panelLink65").click(force=True)
    await page.get_by_role("textbox", name="Last Name, First Name").wait_for(timeout=30000)


async def run(patients):
    log.info("=" * 50)
    log.info("STEP 4 - ECW: UPLOADING FORMS")
    log.info("=" * 50)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=200)
        context = await browser.new_context()
        page = await context.new_page()

        await ecw_login(page)

        await page.wait_for_selector("#jellybean-panelLink65", timeout=30000)
        await page.locator("#jellybean-panelLink65").click(force=True)
        await page.get_by_role("textbox", name="Last Name, First Name").wait_for(timeout=30000)
        log.info("Patient search ready!")

        uploaded_ok = []

        for index, patient in enumerate(patients):
            log.info(f"Processing {index+1}/{len(patients)}: {patient['last_name']} {patient['first_name']}")
            folder_path = os.path.join(settings.DOC_FOLDER, patient["folder_name"])
            if not os.path.exists(folder_path):
                log.info("Folder not found - skipping")
                continue
            all_files = glob.glob(os.path.join(folder_path, "*"))
            all_files.sort(key=os.path.getmtime)
            if not all_files:
                log.info("No files found - skipping")
                continue
            try:
                search_box = page.get_by_role("textbox", name="Last Name, First Name")
                await search_box.wait_for(timeout=30000)
                await search_box.fill(patient['search_name'])
                await page.wait_for_selector("#patientLName1", timeout=30000)
                try:
                    await page.get_by_role("cell", name=patient['last_name'], exact=False).first.click()
                    await page.get_by_text(patient['last_name'], exact=False).first.click()
                    log.info("Patient selected!")
                except Exception:
                    log.info("Could not find patient - skipping")
                    await _go_to_search(page)
                    continue
                try:
                    if await page.locator("text=Please select a patient").is_visible():
                        await page.get_by_role("button", name="OK").click()
                        await _go_to_search(page)
                        continue
                except Exception:
                    pass
                await page.wait_for_selector('button:has-text("Patient Docs")', timeout=30000)
                await page.get_by_role("button", name="Patient Docs").click()
                await page.get_by_role("textbox", name="Quick Search").wait_for(timeout=30000)
                await page.get_by_role("textbox", name="Quick Search").fill("chart")
                await page.wait_for_selector('a:has-text("Chart Documents")', timeout=30000)
                await page.locator("a").filter(has_text="Chart Documents").nth(1).click()
                await asyncio.sleep(2)
                existing_docs = []
                try:
                    doc_links = await page.locator('a[id^="patientdocsTreeLink"]').all()
                    for link in doc_links:
                        doc_object = await link.get_attribute('document-object')
                        if doc_object:
                            doc_data = json.loads(doc_object)
                            label = doc_data.get('label', '').strip().lower()
                            if label:
                                existing_docs.append(label)
                except Exception:
                    pass
                new_files = []
                for f in all_files:
                    filename = os.path.splitext(os.path.basename(f))[0].lower()
                    if filename not in existing_docs:
                        new_files.append(f)
                    else:
                        log.info(f"Already uploaded: {filename}")
                if not new_files:
                    log.info("All files already uploaded!")
                    uploaded_ok.append(patient)
                    await _go_to_search(page)
                    continue
                for file_index, file_path in enumerate(new_files):
                    log.info(f"Uploading {file_index+1}/{len(new_files)}: {os.path.basename(file_path)}")
                    if file_index > 0:
                        await page.get_by_role("textbox", name="Quick Search").wait_for(timeout=30000)
                        await page.get_by_role("textbox", name="Quick Search").fill("chart")
                        await page.wait_for_selector('a:has-text("Chart Documents")', timeout=30000)
                        await page.locator("a").filter(has_text="Chart Documents").nth(1).click()
                        await asyncio.sleep(1)
                    async with page.expect_file_chooser() as fc_info:
                        await page.wait_for_selector("#patientdocsBtn4", timeout=30000)
                        await page.locator("#patientdocsBtn4").click()
                    file_chooser = await fc_info.value
                    await file_chooser.set_files(file_path)
                    await asyncio.sleep(1)
                    try:
                        if await page.locator("text=Please select a category").is_visible():
                            await page.get_by_role("button", name="OK").click()
                    except Exception:
                        pass
                    await page.wait_for_selector('button.commonButton:has-text("OK")', timeout=30000)
                    await page.locator('button.commonButton:has-text("OK")').click()
                    await page.wait_for_selector('#btnOk', timeout=30000)
                    await page.locator('#btnOk').click()
                    log.info(f"File {file_index+1} saved!")
                    await asyncio.sleep(1)
                log.info("All files uploaded!")
                uploaded_ok.append(patient)
                await _go_to_search(page)
            except Exception as e:
                log.info(f"Error: {e}")
                try:
                    await _go_to_search(page)
                except Exception:
                    pass
                continue

        await browser.close()
        log.info("All forms uploaded to eCW!")
        return uploaded_ok
