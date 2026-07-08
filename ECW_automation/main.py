import asyncio
import os
import re
import glob
import json
import openpyxl
from datetime import date, timedelta
from dotenv import load_dotenv
from playwright.async_api import async_playwright

load_dotenv()

# --- CREDENTIALS ---
ECW_USERNAME = os.getenv("ECW_USERNAME")
ECW_PASSWORD = os.getenv("ECW_PASSWORD")
PEDIFORMS_ORG = os.getenv("PEDIFORMS_ORG")
PEDIFORMS_EMAIL = os.getenv("PEDIFORMS_EMAIL")
PEDIFORMS_PASSWORD = os.getenv("PEDIFORMS_PASSWORD")
PCARELINK_EMAIL = os.getenv("PCARELINK_EMAIL")
PCARELINK_PASSWORD = os.getenv("PCARELINK_PASSWORD")
PCARELINK_PRACTICE = os.getenv("PCARELINK_PRACTICE")
PCARELINK_MESSAGE = os.getenv("PCARELINK_MESSAGE")
EXCEL_PATH = os.getenv("EXCEL_PATH")
FILTERED_EXCEL_PATH = os.path.join(os.path.dirname(EXCEL_PATH), "filtered_schedule.xlsx") if EXCEL_PATH else None
DOC_FOLDER = os.getenv("ECW_PATIENTS_DOC_FOLDER")

# --- SETTINGS ---
INITIAL_CHECK_INTERVAL = 180

LONG_CHECK_INTERVAL = 360

VISIT_TYPE_TO_FORM = {
    "9 MONTH WC": "ASQ9Mos",
    "12 MONTH WC": "ASQ 12 Months",
    "12 MONTHWC": "ASQ 12 Months",
    "1 YEAR WC": "ASQ 12 Months",
    "15 MONTH WC": "ASQ 18 Months",
    "15 MONTHWC": "ASQ 18 Months",
    "18 MONTH WC": "ASQ 18 Months",
    "18 MONTHWC": "ASQ 18 Months",
    "24 MONTH WC": "ASQ 24 Months",
    "24 MONTHWC": "ASQ 24 Months",
    "2 YEAR WC": "ASQ 24 Months",
    "30 MONTH WC": "ASQ30",
    "30 MONTHWC": "ASQ30",
    "3 YEAR WC": "ASQ 36 Months",
    "36 MONTH WC": "ASQ 36 Months",
    "36 MONTHWC": "ASQ 36 Months",
    "4 YEAR WC": "ASQ 48 Months",
    "48 MONTH WC": "ASQ 48 Months",
    "48 MONTHWC": "ASQ 48 Months",
}

VISIT_TYPE_TO_FORM_FILENAME = {
    "9 MONTH WC": "ASQ9Mos",
    "12 MONTH WC": "ASQ_12_Months",
    "12 MONTHWC": "ASQ_12_Months",
    "1 YEAR WC": "ASQ_12_Months",
    "15 MONTH WC": "ASQ_18_Months",
    "15 MONTHWC": "ASQ_18_Months",
    "18 MONTH WC": "ASQ_18_Months",
    "18 MONTHWC": "ASQ_18_Months",
    "24 MONTH WC": "ASQ_24_Months",
    "24 MONTHWC": "ASQ_24_Months",
    "2 YEAR WC": "ASQ_24_Months",
    "30 MONTH WC": "ASQ30",
    "30 MONTHWC": "ASQ30",
    "3 YEAR WC": "ASQ_36_Months",
    "36 MONTH WC": "ASQ_36_Months",
    "36 MONTHWC": "ASQ_36_Months",
    "4 YEAR WC": "ASQ_48_Months",
    "48 MONTH WC": "ASQ_48_Months",
    "48 MONTHWC": "ASQ_48_Months",
}

# ─────────────────────────────────────────
# SHARED HELPERS
# ─────────────────────────────────────────

def read_patients_from_excel():
    """
    Production filter: all patients with 9-48 month WC visit types.
    No visit reason filter — all ASQ-eligible patients included.
    """
    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True)
    ws = wb.active
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    col = {name: idx for idx, name in enumerate(headers)}
    patients = []
    filtered_rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        acct_no = row[col["Patient Acct No"]]
        last_name = row[col["Patient Last Name"]]
        first_name = row[col["Patient First Name"]]
        visit_type_raw = row[col["Visit Type"]]
        if not acct_no:
            continue
        visit_type_desc = str(visit_type_raw).split(":")[-1].strip().upper() if visit_type_raw else ""
        form_name = VISIT_TYPE_TO_FORM.get(visit_type_desc, None)
        form_filename = VISIT_TYPE_TO_FORM_FILENAME.get(visit_type_desc, "form")
        if not form_name:
            print(f"Skipping {last_name} {first_name} - no ASQ form for visit type: {visit_type_desc!r}")
            continue
        patients.append({
            "acct_no": str(acct_no).strip(),
            "last_name": str(last_name).strip() if last_name else "",
            "first_name": str(first_name).strip() if first_name else "",
            "folder_name": f"{last_name} {first_name}_doc".strip(),
            "search_name": f"{last_name},{first_name}".strip(),
            "visit_type": visit_type_desc,
            "form_name": form_name,
            "form_filename": form_filename,
        })
        filtered_rows.append(row)

    # Save filtered Excel for Pediforms import
    filtered_wb = openpyxl.Workbook()
    filtered_ws = filtered_wb.active
    filtered_ws.append(headers)
    for row in filtered_rows:
        filtered_ws.append(list(row))
    filtered_wb.save(FILTERED_EXCEL_PATH)
    print(f"Filtered schedule ({len(filtered_rows)} patients) saved to: {FILTERED_EXCEL_PATH}")
    return patients

def ensure_patient_folder(patient):
    folder_path = os.path.join(DOC_FOLDER, patient["folder_name"])
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"Created folder: {folder_path}")
    return folder_path

# ─────────────────────────────────────────
# STEP 0 — ECW: EXPORT SCHEDULE
# ─────────────────────────────────────────

async def click_calendar_option(iframe, day_str, descriptions_to_try, label):
    """Try clicking calendar option with multiple description variants then fallback."""
    for desc in descriptions_to_try:
        try:
            await iframe.get_by_role("option", name=day_str, description=desc, exact=True).click(timeout=5000)
            print(f"{label} date set! (description='{desc}')")
            return
        except:
            pass
    # Final fallback - no description
    try:
        await iframe.get_by_role("option", name=day_str, exact=True).first.click(timeout=5000)
        print(f"{label} date set! (no description fallback)")
        return
    except Exception as e:
        raise RuntimeError(f"Could not set {label} date to day {day_str}: {e}")

async def ecw_export_schedule():
    print("\n" + "="*50)
    print("STEP 0 — ECW: EXPORTING SCHEDULE (TODAY + 7 DAYS)")
    print("="*50)

    # Both 2-letter and 3-letter day description variants
    day_desc_2 = {0: "Mo", 1: "Tu", 2: "We", 3: "Th", 4: "Fr", 5: "Sa", 6: "Su"}
    day_desc_3 = {0: "Mon", 1: "Tue", 2: "Wed", 3: "Thu", 4: "Fri", 5: "Sat", 6: "Sun"}

    start_date = date.today()
    end_date = start_date + timedelta(days=7)
    start_day_str = str(start_date.day)
    end_day_str = str(end_date.day)
    start_descs = [day_desc_2[start_date.weekday()], day_desc_3[start_date.weekday()]]
    end_descs = [day_desc_2[end_date.weekday()], day_desc_3[end_date.weekday()]]
    print(f"Date range: {start_date} to {end_date}")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        # --- LOGIN ---
        print("Logging into eCW...")
        await page.goto(
            "https://txsnmbapp.ecwcloud.com/mobiledoc/jsp/webemr/login/newLogin.jsp",
            timeout=120000, wait_until="domcontentloaded"
        )
        await asyncio.sleep(3)
        await page.get_by_role("textbox", name="Enter username to continue").fill(ECW_USERNAME)
        await page.get_by_role("button", name="Next").click()
        await asyncio.sleep(8)
        await page.click('input[type="password"]')
        await page.keyboard.type(ECW_PASSWORD)
        await page.keyboard.press("Enter")
        print("Login submitted...")

        await page.wait_for_selector('#jellybean-panelLink33', timeout=120000)
        print("Home page loaded!")

        print("Waiting for eCW to fully load...")
        try:
            await page.wait_for_selector('#load', state='hidden', timeout=120000)
            print("eCW fully loaded!")
        except:
            print("Loading screen already hidden!")

        print("Checking for License Alert...")
        dismissed = False
        for i in range(20):
            try:
                if await page.locator("#providerLicense button.clsMyButton").is_visible():
                    await page.click("#providerLicense button.clsMyButton")
                    print("License Alert dismissed!")
                    dismissed = True
                    break
            except:
                pass
            await asyncio.sleep(1)
        if not dismissed:
            print("No License Alert, continuing...")
        await asyncio.sleep(2)

        # --- NAVIGATE TO EBO REPORTS ---
        print("Navigating to eBO Reports...")
        await page.locator("#jellybean-panelLink4").click()
        await asyncio.sleep(1)
        await page.get_by_text("Menu", exact=True).click()
        await asyncio.sleep(1)
        await page.locator("#pane6").get_by_text("Reports").click()
        await asyncio.sleep(1)

        async with page.expect_popup(timeout=60000) as page1_info:
            await page.get_by_text("eBO Reports CTRL + ALT + E").click()
        page1 = await page1_info.value
        await page1.goto(
            "https://txsnmbebo.ecwcloud.com/bi/?perspective=home",
            timeout=60000, wait_until="domcontentloaded"
        )
        await asyncio.sleep(5)
        print("eBO Reports opened!")

        # --- NAVIGATE TO ENCOUNTER PATIENT DOWNLOAD ---
        await page1.get_by_role("link", name="eCWEBO", exact=True).click()
        await asyncio.sleep(2)
        await page1.get_by_role("link", name="- Administrative Reports").click()
        await asyncio.sleep(2)
        await page1.get_by_role("link", name="- Encounter Patient Download").click()
        await asyncio.sleep(8)
        print("Encounter Patient Download opened!")

        # --- WAIT IF REPORT ALREADY RUNNING ---
        try:
            if await page1.get_by_text("Your report is running").is_visible():
                print("Report already running — waiting...")
                for i in range(90):
                    if not await page1.get_by_text("Your report is running").is_visible():
                        print("Report finished.")
                        break
                    print(f"Still running... ({i+1}/90)")
                    await asyncio.sleep(2)
        except:
            pass

        # --- WAIT FOR IFRAME ---
        print("Waiting for iframe to load...")
        iframe = page1.locator("iframe[name=\"iD6D96C5E47F347C9B95828AC68A2D69B\"]").content_frame
        await iframe.get_by_role("img").first.wait_for(timeout=60000)
        await asyncio.sleep(3)
        print("Iframe loaded!")

        # --- SET DATES ---
        print(f"Setting start date: {start_date}")
        await iframe.get_by_role("img").first.click()
        await asyncio.sleep(2)
        await click_calendar_option(iframe, start_day_str, start_descs, "Start")
        await asyncio.sleep(1)

        print(f"Setting end date: {end_date}")
        await iframe.get_by_role("img").nth(1).click()
        await asyncio.sleep(2)
        await click_calendar_option(iframe, end_day_str, end_descs, "End")
        await asyncio.sleep(1)

        # --- CLICK OK ---
        await iframe.get_by_role("button", name="OK").click()
        await asyncio.sleep(2)
        print(f"Date range confirmed: {start_date} to {end_date}")

        # --- WAIT FOR REPORT TO GENERATE ---
        print("Waiting for report to generate (2-3 minutes)...")
        for i in range(120):
            try:
                is_disabled = await page1.locator("button[aria-label='Select a format']").get_attribute("disabled")
                if is_disabled is None:
                    print("Report ready!")
                    break
                print(f"Report still running... ({i+1}/120)")
            except:
                pass
            await asyncio.sleep(2)

        await asyncio.sleep(1)

        # --- DOWNLOAD EXCEL ---
        print("Clicking Select a format...")
        await page1.get_by_role("button", name="Select a format").click()
        await asyncio.sleep(3)
        await page1.get_by_role("link", name="Excel data").wait_for(timeout=15000)

        print("Clicking Excel data — waiting for download (may take 2-3 minutes)...")
        async with page1.expect_download(timeout=300000) as download_info:
            await page1.get_by_role("link", name="Excel data").click()
        download = await download_info.value
        await download.save_as(EXCEL_PATH)
        print(f"Excel saved to: {EXCEL_PATH}")

        await browser.close()
        print("\nSchedule exported successfully!")

# ─────────────────────────────────────────
# STEP 1 — PEDIFORMS: SEND FORMS
# ─────────────────────────────────────────

async def pediforms_send_forms(patients):
    print("\n" + "="*50)
    print("STEP 1 — PEDIFORMS: SENDING FORMS")
    print("="*50)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=300)
        context = await browser.new_context()
        page = await context.new_page()

        print("Logging into Pediforms...")
        await page.goto("https://admin.pediformpro.com/staff/login", timeout=60000, wait_until="domcontentloaded")
        await asyncio.sleep(3)
        await page.get_by_role("textbox", name="Organization name").fill(PEDIFORMS_ORG)
        await page.get_by_role("textbox", name="Email").fill(PEDIFORMS_EMAIL)
        await page.get_by_role("textbox", name="Password").fill(PEDIFORMS_PASSWORD)
        await page.get_by_role("button", name="Sign in").click()
        await page.wait_for_load_state("networkidle")
        print("Logged in!")

        print("Uploading filtered schedule Excel...")
        await page.get_by_role("button", name="Choose File").set_input_files(FILTERED_EXCEL_PATH)
        await page.get_by_role("button", name="Import schedule").click()
        await page.wait_for_load_state("networkidle")
        print("Schedule imported!")

        await page.get_by_role("combobox").nth(4).select_option("week")
        await page.wait_for_timeout(1000)

        for patient in patients:
            try:
                print(f"\nSending form for {patient['acct_no']} ({patient['last_name']} {patient['first_name']})")
                search_box = page.get_by_role("textbox", name="Search…")
                await search_box.click()
                await search_box.fill(patient["acct_no"])
                await page.wait_for_timeout(1000)
                try:
                    await page.get_by_role("link", name="View").click(timeout=10000)
                except:
                    print(f"Patient {patient['acct_no']} not found - skipping")
                    continue
                await page.get_by_role("button", name="+ Send a form").click()
                await page.wait_for_timeout(500)
                try:
                    await page.locator("label").filter(has_text=patient["form_name"]).first.click(timeout=10000)
                except:
                    print(f"Form '{patient['form_name']}' not found - skipping")
                    await page.get_by_role("link", name="← Back to today's patients").click()
                    await page.wait_for_load_state("networkidle")
                    continue
                await page.get_by_role("button", name="Send form").click()
                print(f"Sent '{patient['form_name']}' successfully!")
                await page.get_by_role("link", name="← Back to today's patients").click()
                await page.wait_for_load_state("networkidle")
            except Exception as e:
                print(f"Error: {e}")
                try:
                    await page.get_by_role("link", name="← Back to today's patients").click()
                    await page.wait_for_load_state("networkidle")
                except:
                    pass
                continue

        await browser.close()
        print("\nForms sent!")

# ─────────────────────────────────────────
# STEP 2 — PCARELINK: SEND MESSAGES
# ─────────────────────────────────────────

async def pcarelink_send_messages(patients):
    print("\n" + "="*50)
    print("STEP 2 — PCARELINK: SENDING MESSAGES")
    print("="*50)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=300)
        context = await browser.new_context()
        page = await context.new_page()

        print("Logging into pcarelink...")
        await page.goto("https://app.pcarelink.com/login", timeout=60000, wait_until="domcontentloaded")
        await asyncio.sleep(3)
        await page.get_by_role("textbox", name="Enter email id").fill(PCARELINK_EMAIL)
        await page.get_by_role("textbox", name="Enter password").fill(PCARELINK_PASSWORD)
        await page.locator('[data-test-id="pcl-login-signInButton"]').click()
        await asyncio.sleep(5)
        print("Logged in!")

        await page.locator('[data-test-id="pcl-menuDropDownComponent"]').click()
        await page.locator('[data-test-id="pcl-dashboard-popOver1"]').click()
        await page.wait_for_load_state("networkidle")

        await page.get_by_role("button", name="Filter by Practice").click()
        await page.get_by_text(f"{PCARELINK_PRACTICE}Round Rock, us").click()
        await page.wait_for_timeout(2000)
        print(f"Filtered by practice: {PCARELINK_PRACTICE}")

        for patient in patients:
            try:
                print(f"\nSending message for {patient['acct_no']} ({patient['last_name']} {patient['first_name']})")
                search_box = page.get_by_role("searchbox", name="Enter patient first name or")
                await search_box.click()
                await search_box.fill(patient["acct_no"])
                await page.wait_for_timeout(2000)
                try:
                    await page.get_by_text(f"{patient['last_name'].upper()}, {patient['first_name'].upper()}").first.click(timeout=10000)
                except:
                    try:
                        await page.locator(".patient-result, .search-result").first.click(timeout=5000)
                    except:
                        print(f"Patient {patient['acct_no']} not found - skipping")
                        continue
                await page.wait_for_timeout(1000)
                await page.locator('[data-test-id="pcl-payments-sendMessageLinkGuarantorDrawer"]').click()
                await page.wait_for_timeout(1000)
                try:
                    await page.get_by_role("button", name=PCARELINK_PRACTICE).click(timeout=5000)
                    await page.wait_for_timeout(500)
                    await page.get_by_role("menuitem", name="Appointment Scheduling").get_by_role("radio").check(timeout=5000)
                    await page.locator("#menu- > div").first.click(timeout=5000)
                    await page.wait_for_timeout(500)
                except:
                    print("Message type selection skipped")
                message_box = page.get_by_role("textbox", name="Type your response and send")
                await message_box.click()
                await message_box.fill(PCARELINK_MESSAGE)
                print("Message typed!")
                await page.locator('[data-test-id="pcl-payments-sendMessageButton"]').click()
                print("Message sent!")
                await page.wait_for_timeout(1000)
                await page.locator('[data-test-id="pcl-appointments-closePatientsDetails"]').click()
                await page.wait_for_timeout(1000)
            except Exception as e:
                print(f"Error: {e}")
                try:
                    await page.locator('[data-test-id="pcl-appointments-closePatientsDetails"]').click()
                except:
                    pass
                continue

        await browser.close()
        print("\nMessages sent!")

# ─────────────────────────────────────────
# STEP 3 — PEDIFORMS: CHECK & DOWNLOAD
# ─────────────────────────────────────────

async def check_and_download_completed(page, patients, downloaded):
    """
    Search each patient without status filter, then check if they have
    a Completed submission. If multiple submissions exist, click the
    most recent Completed one and download it.
    """
    print(f"\nChecking for completed forms...")
    # Set week filter only — no status filter to avoid missing patients
    # with multiple submissions where overall status may show differently
    await page.get_by_role("combobox").nth(4).select_option("week")
    await page.wait_for_timeout(1000)

    for patient in patients:
        if patient["acct_no"] in downloaded:
            continue
        try:
            search_box = page.get_by_role("textbox", name="Search…")
            await search_box.click()
            await search_box.fill(patient["acct_no"])
            await page.wait_for_timeout(2000)

            view_count = await page.get_by_role("link", name="View").count()
            if view_count == 0:
                print(f"Patient {patient['acct_no']} not found this week")
                await search_box.fill("")
                continue

            # Check if Completed badge is visible in the row
            completed_visible = await page.get_by_text("Completed", exact=True).count()
            if completed_visible == 0:
                print(f"Patient {patient['acct_no']} not completed yet")
                await search_box.fill("")
                continue

            print(f"Patient {patient['acct_no']} has Completed form — opening...")
            await page.get_by_role("link", name="View").first.click()
            await page.wait_for_timeout(1000)

            # Handle multiple submissions — find the most recent Completed one
            # Look for Export PDF buttons next to (completed) submissions
            # The page shows submissions list; click Export PDF on first completed
            export_buttons = page.get_by_role("button", name="Export PDF")
            export_count = await export_buttons.count()

            if export_count == 0:
                print(f"No Export PDF button found for {patient['acct_no']} - skipping")
                await page.get_by_role("link", name="← Back to today's patients").click()
                await page.wait_for_timeout(1000)
                continue

            folder_path = ensure_patient_folder(patient)
            file_name = f"{patient['last_name']}_{patient['first_name']}_{patient['form_filename']}.pdf"
            save_path = os.path.join(folder_path, file_name)

            print(f"Downloading most recent completed form for {patient['acct_no']}...")
            # Click the FIRST Export PDF — Pediforms shows most recent first
            async with page.expect_download() as download_info:
                await export_buttons.first.click()
            download = await download_info.value
            await download.save_as(save_path)
            print(f"Saved: {save_path}")

            downloaded.add(patient["acct_no"])

            await page.get_by_role("link", name="← Back to today's patients").click()
            await page.wait_for_timeout(1000)

        except Exception as e:
            print(f"Error downloading for {patient['acct_no']}: {e}")
            try:
                await page.get_by_role("link", name="← Back to today's patients").click()
            except:
                pass
            continue

async def pediforms_check_and_download(patients):
    print("\n" + "="*50)
    print("STEP 3 — PEDIFORMS: CHECKING FOR COMPLETED FORMS")
    print("="*50)

    downloaded = set()
    total = len(patients)
    check_num = 0
    elapsed_minutes = 0

    while True:
        check_num += 1
        if elapsed_minutes < 720:
            interval = INITIAL_CHECK_INTERVAL
            phase = "Phase 1 (every 3 hours)"
        
        else:
            interval = LONG_CHECK_INTERVAL
            phase = "Phase 2(every 6 hours)"

        print(f"\nCheck #{check_num} — {len(downloaded)}/{total} forms — {phase}")

        try:
            async with async_playwright() as p:
                browser = await p.chromium.launch(headless=False, slow_mo=300)
                context = await browser.new_context()
                page = await context.new_page()
                print("Logging into Pediforms...")
                await page.goto("https://admin.pediformpro.com/staff/login", timeout=60000, wait_until="domcontentloaded")
                await asyncio.sleep(3)
                await page.get_by_role("textbox", name="Organization name").fill(PEDIFORMS_ORG)
                await page.get_by_role("textbox", name="Email").fill(PEDIFORMS_EMAIL)
                await page.get_by_role("textbox", name="Password").fill(PEDIFORMS_PASSWORD)
                await page.get_by_role("button", name="Sign in").click()
                await page.wait_for_load_state("networkidle")

                await check_and_download_completed(page, patients, downloaded)
                await browser.close()

            if len(downloaded) > 0:
                print(f"\n{len(downloaded)} forms downloaded — uploading to eCW now...")
                await ecw_upload_forms(patients)

            if len(downloaded) >= total:
                print(f"\nAll {total} forms downloaded and uploaded!")
                return True

        except Exception as e:
            print(f"\nCheck #{check_num} failed: {e}")
            print("Will retry on next scheduled check.")

        print(f"\nWaiting {interval} minutes before next check...")
        await asyncio.sleep(interval * 60)
        elapsed_minutes += interval

# ─────────────────────────────────────────
# STEP 4 — ECW: UPLOAD FORMS
# ─────────────────────────────────────────

async def ecw_upload_forms(patients):
    print("\n" + "="*50)
    print("STEP 4 — ECW: UPLOADING FORMS")
    print("="*50)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=200)
        context = await browser.new_context()
        page = await context.new_page()

        print("Opening eCW login page...")
        await page.goto(
            "https://txsnmbapp.ecwcloud.com/mobiledoc/jsp/webemr/login/newLogin.jsp",
            timeout=120000, wait_until="domcontentloaded"
        )
        await asyncio.sleep(3)
        await page.get_by_role("textbox", name="Enter username to continue").fill(ECW_USERNAME)
        await page.get_by_role("button", name="Next").click()
        print("Username entered!")
        await asyncio.sleep(8)
        await page.click('input[type="password"]')
        await page.keyboard.type(ECW_PASSWORD)
        print("Password entered!")
        await page.keyboard.press("Enter")
        print("Login submitted...")

        await page.wait_for_selector('#jellybean-panelLink33', timeout=120000)
        print("Home page loaded!")
        try:
            await page.wait_for_selector('#load', state='hidden', timeout=120000)
            print("eCW fully loaded!")
        except:
            pass

        print("Checking for License Alert...")
        dismissed = False
        for i in range(20):
            try:
                if await page.locator("#providerLicense button.clsMyButton").is_visible():
                    await page.click("#providerLicense button.clsMyButton")
                    print("License Alert dismissed!")
                    dismissed = True
                    break
            except:
                pass
            await asyncio.sleep(1)
        if not dismissed:
            print("No License Alert, continuing...")
        await asyncio.sleep(2)

        await page.wait_for_selector("#jellybean-panelLink65", timeout=30000)
        await page.locator("#jellybean-panelLink65").click()
        await page.get_by_role("textbox", name="Last Name, First Name").wait_for(timeout=30000)
        print("Patient search ready!")

        for index, patient in enumerate(patients):
            print(f"\nProcessing {index+1}/{len(patients)}: {patient['last_name']} {patient['first_name']}")
            folder_path = os.path.join(DOC_FOLDER, patient["folder_name"])
            if not os.path.exists(folder_path):
                print(f"Folder not found - skipping")
                continue
            all_files = glob.glob(os.path.join(folder_path, "*"))
            all_files.sort(key=os.path.getmtime)
            if not all_files:
                print(f"No files found - skipping")
                continue
            try:
                search_box = page.get_by_role("textbox", name="Last Name, First Name")
                await search_box.wait_for(timeout=30000)
                await search_box.fill(patient['search_name'])
                await page.wait_for_selector("#patientLName1", timeout=30000)
                try:
                    await page.get_by_role("cell", name=patient['last_name'], exact=False).first.click()
                    await page.get_by_text(patient['last_name'], exact=False).first.click()
                    print("Patient selected!")
                except:
                    print("Could not find patient - skipping")
                    await _go_to_search(page)
                    continue
                try:
                    if await page.locator("text=Please select a patient").is_visible():
                        await page.get_by_role("button", name="OK").click()
                        await _go_to_search(page)
                        continue
                except:
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
                except:
                    pass
                new_files = []
                for f in all_files:
                    filename = os.path.splitext(os.path.basename(f))[0].lower()
                    if filename not in existing_docs:
                        new_files.append(f)
                    else:
                        print(f"Already uploaded: {filename}")
                if not new_files:
                    print("All files already uploaded!")
                    await _go_to_search(page)
                    continue
                for file_index, file_path in enumerate(new_files):
                    print(f"Uploading {file_index+1}/{len(new_files)}: {os.path.basename(file_path)}")
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
                    except:
                        pass
                    await page.wait_for_selector('button.commonButton:has-text("OK")', timeout=30000)
                    await page.locator('button.commonButton:has-text("OK")').click()
                    await page.wait_for_selector('#btnOk', timeout=30000)
                    await page.locator('#btnOk').click()
                    print(f"File {file_index+1} saved!")
                    await asyncio.sleep(1)
                print(f"All files uploaded!")
                await _go_to_search(page)
            except Exception as e:
                print(f"Error: {e}")
                try:
                    await _go_to_search(page)
                except:
                    pass
                continue

        await browser.close()
        print("\nAll forms uploaded to eCW!")

async def _go_to_search(page):
    await page.keyboard.press("Escape")
    await asyncio.sleep(1)
    await page.keyboard.press("Escape")
    await asyncio.sleep(1)
    try:
        await page.wait_for_selector("#patient-hubBtn1", timeout=10000)
        await page.locator("#patient-hubBtn1").click()
        await asyncio.sleep(1)
    except:
        pass
    await page.wait_for_selector("#jellybean-panelLink65", timeout=30000)
    await page.locator("#jellybean-panelLink65").click()
    await page.get_by_role("textbox", name="Last Name, First Name").wait_for(timeout=30000)

# ─────────────────────────────────────────
# MAIN PIPELINE — PRODUCTION
# ─────────────────────────────────────────

async def main():
    print("="*50)
    print("PRODUCTION PIPELINE — ALL ASQ PATIENTS")
    print("="*50)

    await ecw_export_schedule()

    patients = read_patients_from_excel()
    print(f"\nFound {len(patients)} ASQ patients")

    if not patients:
        print("No ASQ patients found in Excel!")
        return

    await pediforms_send_forms(patients)
    await pcarelink_send_messages(patients)
    await pediforms_check_and_download(patients)

    print("\n" + "="*50)
    print("PIPELINE COMPLETE!")
    print("="*50)

if __name__ == "__main__":
    asyncio.run(main())

# import asyncio
# import os
# import glob
# import json
# import openpyxl
# from dotenv import load_dotenv
# from playwright.async_api import async_playwright

# load_dotenv()

# # --- CREDENTIALS ---
# ECW_USERNAME = os.getenv("ECW_USERNAME")
# ECW_PASSWORD = os.getenv("ECW_PASSWORD")
# PEDIFORMS_ORG = os.getenv("PEDIFORMS_ORG")
# PEDIFORMS_EMAIL = os.getenv("PEDIFORMS_EMAIL")
# PEDIFORMS_PASSWORD = os.getenv("PEDIFORMS_PASSWORD")
# PCARELINK_EMAIL = os.getenv("PCARELINK_EMAIL")
# PCARELINK_PASSWORD = os.getenv("PCARELINK_PASSWORD")
# PCARELINK_PRACTICE = os.getenv("PCARELINK_PRACTICE")
# PCARELINK_MESSAGE = os.getenv("PCARELINK_MESSAGE")
# EXCEL_PATH = os.getenv("EXCEL_PATH")
# DOC_FOLDER = os.getenv("ECW_PATIENTS_DOC_FOLDER")

# # --- SETTINGS ---
# CHECK_INTERVAL_MINUTES = 10  # Check Pediforms every 10 minutes
# MAX_WAIT_HOURS = 24  # Stop checking after 24 hours

# VISIT_TYPE_TO_FORM = {
#     "9 MONTH WC": "ASQ9Mos",
#     "12 MONTH WC": "ASQ 12 Months",
#     "1 YR WC": "ASQ 12 Months",
#     "18 MONTH WC": "ASQ 18 Months",
#     "24 MONTH WC": "ASQ 24 Months",
#     "2 YR WC": "ASQ 24 Months",
#     "30 MONTH WC": "ASQ30",
#     "3 YEAR WC": "ASQ 36 Months",
#     "36 MONTH WC": "ASQ 36 Months",
#     "4 YEAR WC": "ASQ 48 Months",
#     "48 MONTH WC": "ASQ 48 Months",
# }

# VISIT_TYPE_TO_FORM_FILENAME = {
#     "9 MONTH WC": "ASQ9Mos",
#     "12 MONTH WC": "ASQ_12_Months",
#     "1 YR WC": "ASQ_12_Months",
#     "18 MONTH WC": "ASQ_18_Months",
#     "24 MONTH WC": "ASQ_24_Months",
#     "2 YR WC": "ASQ_24_Months",
#     "30 MONTH WC": "ASQ30",
#     "3 YEAR WC": "ASQ_36_Months",
#     "36 MONTH WC": "ASQ_36_Months",
#     "4 YEAR WC": "ASQ_48_Months",
#     "48 MONTH WC": "ASQ_48_Months",
# }

# # ─────────────────────────────────────────
# # SHARED HELPERS
# # ─────────────────────────────────────────

# def read_patients_from_excel():
#     wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True)
#     ws = wb.active
#     headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
#     col = {name: idx for idx, name in enumerate(headers)}
#     patients = []
#     for row in ws.iter_rows(min_row=2, values_only=True):
#         acct_no = row[col["Patient Acct No"]]
#         last_name = row[col["Patient Last Name"]]
#         first_name = row[col["Patient First Name"]]
#         visit_type_raw = row[col["Visit Type"]]
#         if not acct_no:
#             continue
#         visit_type_desc = str(visit_type_raw).split(":")[-1].strip().upper() if visit_type_raw else ""
#         form_name = VISIT_TYPE_TO_FORM.get(visit_type_desc, None)
#         form_filename = VISIT_TYPE_TO_FORM_FILENAME.get(visit_type_desc, "form")
#         patients.append({
#             "acct_no": str(acct_no).strip(),
#             "last_name": str(last_name).strip() if last_name else "",
#             "first_name": str(first_name).strip() if first_name else "",
#             "folder_name": f"{last_name} {first_name}_doc".strip(),
#             "search_name": f"{last_name},{first_name}".strip(),
#             "visit_type": visit_type_desc,
#             "form_name": form_name,
#             "form_filename": form_filename,
#         })
#     return patients

# def ensure_patient_folder(patient):
#     folder_path = os.path.join(DOC_FOLDER, patient["folder_name"])
#     if not os.path.exists(folder_path):
#         os.makedirs(folder_path)
#         print(f"Created folder: {folder_path}")
#     return folder_path

# # ─────────────────────────────────────────
# # STEP 1 — PEDIFORMS: SEND FORMS
# # ─────────────────────────────────────────

# async def pediforms_send_forms(patients):
#     print("\n" + "="*50)
#     print("STEP 1 — PEDIFORMS: SENDING FORMS")
#     print("="*50)

#     async with async_playwright() as p:
#         browser = await p.chromium.launch(headless=False, slow_mo=300)
#         context = await browser.new_context()
#         page = await context.new_page()

#         print("Logging into Pediforms...")
#         await page.goto("https://admin.pediformpro.com/staff/login")
#         await page.get_by_role("textbox", name="Organization name").fill(PEDIFORMS_ORG)
#         await page.get_by_role("textbox", name="Email").fill(PEDIFORMS_EMAIL)
#         await page.get_by_role("textbox", name="Password").fill(PEDIFORMS_PASSWORD)
#         await page.get_by_role("button", name="Sign in").click()
#         await page.wait_for_load_state("networkidle")
#         print("Logged in!")

#         print("Uploading schedule Excel...")
#         await page.get_by_role("button", name="Choose File").set_input_files(EXCEL_PATH)
#         await page.get_by_role("button", name="Import schedule").click()
#         await page.wait_for_load_state("networkidle")
#         print("Schedule imported!")

#         # Set week filter
#         await page.get_by_role("combobox").nth(4).select_option("week")
#         await page.wait_for_timeout(1000)

#         for patient in patients:
#             try:
#                 if not patient["form_name"]:
#                     print(f"No form mapped for {patient['acct_no']} - skipping")
#                     continue

#                 print(f"\nSending form for {patient['acct_no']} ({patient['last_name']} {patient['first_name']})")

#                 search_box = page.get_by_role("textbox", name="Search…")
#                 await search_box.click()
#                 await search_box.fill(patient["acct_no"])
#                 await page.wait_for_timeout(1000)

#                 try:
#                     await page.get_by_role("link", name="View").click(timeout=10000)
#                 except:
#                     print(f"Patient {patient['acct_no']} not found - skipping")
#                     continue

#                 await page.get_by_role("button", name="+ Send a form").click()
#                 await page.wait_for_timeout(500)

#                 try:
#                     await page.locator("label").filter(has_text=patient["form_name"]).first.click(timeout=10000)
#                 except:
#                     print(f"Form '{patient['form_name']}' not found - skipping")
#                     await page.get_by_role("link", name="← Back to today's patients").click()
#                     await page.wait_for_load_state("networkidle")
#                     continue

#                 await page.get_by_role("button", name="Send form").click()
#                 print(f"Sent '{patient['form_name']}' successfully!")

#                 await page.get_by_role("link", name="← Back to today's patients").click()
#                 await page.wait_for_load_state("networkidle")

#             except Exception as e:
#                 print(f"Error: {e}")
#                 try:
#                     await page.get_by_role("link", name="← Back to today's patients").click()
#                     await page.wait_for_load_state("networkidle")
#                 except:
#                     pass
#                 continue

#         await browser.close()
#         print("\nForms sent to all patients!")

# # ─────────────────────────────────────────
# # STEP 2 — PCARELINK: SEND MESSAGES
# # ─────────────────────────────────────────

# async def pcarelink_send_messages(patients):
#     print("\n" + "="*50)
#     print("STEP 2 — PCARELINK: SENDING MESSAGES")
#     print("="*50)

#     async with async_playwright() as p:
#         browser = await p.chromium.launch(headless=False, slow_mo=300)
#         context = await browser.new_context()
#         page = await context.new_page()

#         print("Logging into pcarelink...")
#         await page.goto("https://app.pcarelink.com/login", timeout=60000, wait_until="domcontentloaded")
#         await asyncio.sleep(3)
#         print("Logged in!")

#         await page.get_by_role("textbox", name="Enter email id").fill(PCARELINK_EMAIL)
#         await page.get_by_role("textbox", name="Enter password").fill(PCARELINK_PASSWORD)
#         await page.locator('[data-test-id="pcl-login-signInButton"]').click()
#         await asyncio.sleep(5)
#         print("Login submitted!")

#         if "profile" in page.url:
#             print("Complete-your-profile popup detected - dismissing via ReachMyDr logo...")
#             await page.get_by_text("ReachMyDr", exact=False).first.click()
#             await asyncio.sleep(3)

#         await page.locator('[data-test-id="pcl-menuDropDownComponent"]').click()
#         await page.locator('[data-test-id="pcl-dashboard-popOver1"]').click()
#         await page.wait_for_load_state("networkidle")

#         await page.get_by_role("button", name="Filter by Practice").click()
#         await page.get_by_text(f"{PCARELINK_PRACTICE}Round Rock, us").click()
#         await page.wait_for_timeout(2000)
#         print(f"Filtered by practice: {PCARELINK_PRACTICE}")

#         for patient in patients:
#             try:
#                 print(f"\nSending message for {patient['acct_no']} ({patient['last_name']} {patient['first_name']})")

#                 search_box = page.get_by_role("searchbox", name="Enter patient first name or")
#                 await search_box.click()
#                 await search_box.fill(patient["acct_no"])
#                 await page.wait_for_timeout(2000)

#                 try:
#                     await page.get_by_text(f"{patient['last_name'].upper()}, {patient['first_name'].upper()}").first.click(timeout=10000)
#                 except:
#                     try:
#                         await page.locator(".patient-result, .search-result").first.click(timeout=5000)
#                     except:
#                         print(f"Patient {patient['acct_no']} not found - skipping")
#                         continue

#                 await page.wait_for_timeout(1000)
#                 await page.locator('[data-test-id="pcl-payments-sendMessageLinkGuarantorDrawer"]').click()
#                 await page.wait_for_timeout(1000)

#                 try:
#                     await page.get_by_role("button", name=PCARELINK_PRACTICE).click(timeout=5000)
#                     await page.wait_for_timeout(500)
#                     await page.get_by_role("menuitem", name="Appointment Scheduling").get_by_role("radio").check(timeout=5000)
#                     await page.locator("#menu- > div").first.click(timeout=5000)
#                     await page.wait_for_timeout(500)
#                 except:
#                     print("Message type selection skipped")

#                 message_box = page.get_by_role("textbox", name="Type your response and send")
#                 await message_box.click()
#                 await message_box.fill(PCARELINK_MESSAGE)
#                 print("Message typed!")

#                 await page.locator('[data-test-id="pcl-payments-sendMessageButton"]').click()
#                 print(f"Message sent successfully!")
#                 await page.wait_for_timeout(1000)

#                 await page.locator('[data-test-id="pcl-appointments-closePatientsDetails"]').click()
#                 await page.wait_for_timeout(1000)

#             except Exception as e:
#                 print(f"Error: {e}")
#                 try:
#                     await page.locator('[data-test-id="pcl-appointments-closePatientsDetails"]').click()
#                 except:
#                     pass
#                 continue

#         await browser.close()
#         print("\nMessages sent to all patients!")

# # ─────────────────────────────────────────
# # STEP 3 — PEDIFORMS: CHECK & DOWNLOAD
# # ─────────────────────────────────────────

# async def pediforms_check_and_download(patients):
#     print("\n" + "="*50)
#     print("STEP 3 — PEDIFORMS: CHECKING FOR COMPLETED FORMS")
#     print("="*50)

#     # Track which patients have been downloaded
#     downloaded = set()
#     total = len([p for p in patients if p["form_name"]])
#     max_checks = (MAX_WAIT_HOURS * 60) // CHECK_INTERVAL_MINUTES

#     for check_num in range(max_checks):
#         print(f"\nCheck {check_num+1}/{max_checks} — {len(downloaded)}/{total} forms downloaded")

#         async with async_playwright() as p:
#             browser = await p.chromium.launch(headless=False, slow_mo=300)
#             context = await browser.new_context()
#             page = await context.new_page()

#             print("Logging into Pediforms...")
#             await page.goto("https://admin.pediformpro.com/staff/login")
#             await page.get_by_role("textbox", name="Organization name").fill(PEDIFORMS_ORG)
#             await page.get_by_role("textbox", name="Email").fill(PEDIFORMS_EMAIL)
#             await page.get_by_role("textbox", name="Password").fill(PEDIFORMS_PASSWORD)
#             await page.get_by_role("button", name="Sign in").click()
#             await page.wait_for_load_state("networkidle")

#             # Set completed filter
#             await page.get_by_role("combobox").nth(1).select_option("completed")
#             await page.wait_for_timeout(1000)
#             await page.get_by_role("combobox").nth(4).select_option("week")
#             await page.wait_for_timeout(1000)

#             for patient in patients:
#                 if patient["acct_no"] in downloaded:
#                     continue
#                 if not patient["form_name"]:
#                     continue

#                 try:
#                     search_box = page.get_by_role("textbox", name="Search…")
#                     await search_box.click()
#                     await search_box.fill(patient["acct_no"])
#                     await page.wait_for_timeout(2000)

#                     view_count = await page.get_by_role("link", name="View").count()
#                     if view_count == 0:
#                         print(f"Patient {patient['acct_no']} form not completed yet")
#                         await search_box.fill("")
#                         continue

#                     await page.get_by_role("link", name="View").first.click()
#                     await page.wait_for_timeout(1000)

#                     folder_path = ensure_patient_folder(patient)
#                     file_name = f"{patient['last_name']}_{patient['first_name']}_{patient['form_filename']}.pdf"
#                     save_path = os.path.join(folder_path, file_name)

#                     print(f"Downloading form for {patient['acct_no']}...")
#                     async with page.expect_download() as download_info:
#                         await page.get_by_role("button", name="Export PDF").click()
#                     download = await download_info.value
#                     await download.save_as(save_path)
#                     print(f"Saved: {save_path}")

#                     downloaded.add(patient["acct_no"])

#                     await page.get_by_role("link", name="← Back to today's patients").click()
#                     await page.wait_for_timeout(1000)

#                 except Exception as e:
#                     print(f"Error downloading for {patient['acct_no']}: {e}")
#                     try:
#                         await page.get_by_role("link", name="← Back to today's patients").click()
#                     except:
#                         pass
#                     continue

#             await browser.close()

#         # Check if all done
#         if len(downloaded) >= total:
#             print(f"\nAll {total} forms downloaded!")
#             return True

#         print(f"\n{len(downloaded)}/{total} forms downloaded. Waiting {CHECK_INTERVAL_MINUTES} minutes before next check...")
#         await asyncio.sleep(CHECK_INTERVAL_MINUTES * 60)

#     print(f"\nMax wait time reached. {len(downloaded)}/{total} forms downloaded.")
#     return len(downloaded) > 0

# # ─────────────────────────────────────────
# # STEP 4 — ECW: UPLOAD FORMS
# # ─────────────────────────────────────────

# async def ecw_upload_forms(patients):
#     print("\n" + "="*50)
#     print("STEP 4 — ECW: UPLOADING FORMS")
#     print("="*50)

#     async with async_playwright() as p:
#         browser = await p.chromium.launch(headless=False, slow_mo=200)
#         context = await browser.new_context()
#         page = await context.new_page()

#         print("Opening eCW login page...")
#         await page.goto("https://txsnmbapp.ecwcloud.com/mobiledoc/jsp/webemr/login/newLogin.jsp")
#         await page.get_by_role("textbox", name="Enter username to continue").fill(ECW_USERNAME)
#         await page.get_by_role("button", name="Next").click()
#         await page.get_by_role("textbox", name="Enter Password to continue").fill(ECW_PASSWORD)
#         await page.get_by_role("button", name="Log In").click()
#         print("Login submitted...")

#         await page.wait_for_selector('#jellybean-panelLink33', timeout=120000)
#         print("Home page loaded!")

#         try:
#             await page.wait_for_selector('#load', state='hidden', timeout=120000)
#         except:
#             pass

#         # Handle License Alert
#         for i in range(20):
#             try:
#                 if await page.locator("#providerLicense button.clsMyButton").is_visible():
#                     await page.click("#providerLicense button.clsMyButton")
#                     print(f"License Alert dismissed!")
#                     break
#             except:
#                 pass
#             await asyncio.sleep(1)

#         # Open patient search
#         await page.wait_for_selector("#jellybean-panelLink65", timeout=30000)
#         await page.locator("#jellybean-panelLink65").click()
#         await page.get_by_role("textbox", name="Last Name, First Name").wait_for(timeout=30000)
#         print("Patient search ready!")

#         for index, patient in enumerate(patients):
#             print(f"\nProcessing {index+1}/{len(patients)}: {patient['last_name']} {patient['first_name']}")

#             # Get files from patient folder
#             folder_path = os.path.join(DOC_FOLDER, patient["folder_name"])
#             if not os.path.exists(folder_path):
#                 print(f"Folder not found - skipping")
#                 continue

#             all_files = glob.glob(os.path.join(folder_path, "*"))
#             all_files.sort(key=os.path.getmtime)

#             if not all_files:
#                 print(f"No files found - skipping")
#                 continue

#             try:
#                 # Search patient
#                 search_box = page.get_by_role("textbox", name="Last Name, First Name")
#                 await search_box.wait_for(timeout=30000)
#                 await search_box.fill(patient['search_name'])
#                 await page.wait_for_selector("#patientLName1", timeout=30000)

#                 # Select patient
#                 try:
#                     await page.get_by_role("cell", name=patient['last_name'], exact=False).first.click()
#                     await page.get_by_text(patient['last_name'], exact=False).first.click()
#                     print("Patient selected!")
#                 except:
#                     print("Could not find patient - skipping")
#                     await _go_to_search(page)
#                     continue

#                 # Handle popup
#                 try:
#                     if await page.locator("text=Please select a patient").is_visible():
#                         await page.get_by_role("button", name="OK").click()
#                         await _go_to_search(page)
#                         continue
#                 except:
#                     pass

#                 # Open Patient Docs
#                 await page.wait_for_selector('button:has-text("Patient Docs")', timeout=30000)
#                 await page.get_by_role("button", name="Patient Docs").click()

#                 # Search Chart Documents
#                 await page.get_by_role("textbox", name="Quick Search").wait_for(timeout=30000)
#                 await page.get_by_role("textbox", name="Quick Search").fill("chart")
#                 await page.wait_for_selector('a:has-text("Chart Documents")', timeout=30000)
#                 await page.locator("a").filter(has_text="Chart Documents").nth(1).click()
#                 await asyncio.sleep(2)

#                 # Get existing docs
#                 existing_docs = []
#                 try:
#                     doc_links = await page.locator('a[id^="patientdocsTreeLink"]').all()
#                     for link in doc_links:
#                         doc_object = await link.get_attribute('document-object')
#                         if doc_object:
#                             doc_data = json.loads(doc_object)
#                             label = doc_data.get('label', '').strip().lower()
#                             if label:
#                                 existing_docs.append(label)
#                 except:
#                     pass

#                 # Filter new files
#                 new_files = []
#                 for f in all_files:
#                     filename = os.path.splitext(os.path.basename(f))[0].lower()
#                     if filename not in existing_docs:
#                         new_files.append(f)
#                     else:
#                         print(f"Already uploaded: {filename}")

#                 if not new_files:
#                     print("All files already uploaded!")
#                     await _go_to_search(page)
#                     continue

#                 # Upload files
#                 for file_index, file_path in enumerate(new_files):
#                     print(f"Uploading {file_index+1}/{len(new_files)}: {os.path.basename(file_path)}")

#                     if file_index > 0:
#                         await page.get_by_role("textbox", name="Quick Search").wait_for(timeout=30000)
#                         await page.get_by_role("textbox", name="Quick Search").fill("chart")
#                         await page.wait_for_selector('a:has-text("Chart Documents")', timeout=30000)
#                         await page.locator("a").filter(has_text="Chart Documents").nth(1).click()
#                         await asyncio.sleep(1)

#                     async with page.expect_file_chooser() as fc_info:
#                         await page.wait_for_selector("#patientdocsBtn4", timeout=30000)
#                         await page.locator("#patientdocsBtn4").click()
#                     file_chooser = await fc_info.value
#                     await file_chooser.set_files(file_path)
#                     await asyncio.sleep(1)

#                     try:
#                         if await page.locator("text=Please select a category").is_visible():
#                             await page.get_by_role("button", name="OK").click()
#                     except:
#                         pass

#                     await page.wait_for_selector('button.commonButton:has-text("OK")', timeout=30000)
#                     await page.locator('button.commonButton:has-text("OK")').click()
#                     await page.wait_for_selector('#btnOk', timeout=30000)
#                     await page.locator('#btnOk').click()
#                     print(f"File {file_index+1} saved!")
#                     await asyncio.sleep(1)

#                 print(f"All files uploaded for {patient['last_name']} {patient['first_name']}!")
#                 await _go_to_search(page)

#             except Exception as e:
#                 print(f"Error: {e}")
#                 try:
#                     await _go_to_search(page)
#                 except:
#                     pass
#                 continue

#         await browser.close()
#         print("\nAll forms uploaded to eCW!")

# async def _go_to_search(page):
#     await page.keyboard.press("Escape")
#     await asyncio.sleep(1)
#     await page.keyboard.press("Escape")
#     await asyncio.sleep(1)
#     try:
#         await page.wait_for_selector("#patient-hubBtn1", timeout=10000)
#         await page.locator("#patient-hubBtn1").click()
#         await asyncio.sleep(1)
#     except:
#         pass
#     await page.wait_for_selector("#jellybean-panelLink65", timeout=30000)
#     await page.locator("#jellybean-panelLink65").click()
#     await page.get_by_role("textbox", name="Last Name, First Name").wait_for(timeout=30000)

# # ─────────────────────────────────────────
# # MAIN PIPELINE
# # ─────────────────────────────────────────

# async def main():
#     print("="*50)
#     print("COMPLETE PEDIFORMS PIPELINE")
#     print("="*50)

#     patients = read_patients_from_excel()
#     print(f"Found {len(patients)} patients in Excel")

#     # STEP 1 - Send forms via Pediforms
#     await pediforms_send_forms(patients)

#     # STEP 2 - Send messages via pcarelink
#     await pcarelink_send_messages(patients)

#     # STEP 3 - Wait and check for completed forms
#     has_downloads = await pediforms_check_and_download(patients)

#     # STEP 4 - Upload to eCW
#     if has_downloads:
#         await ecw_upload_forms(patients)
#     else:
#         print("No forms downloaded - skipping eCW upload")

#     print("\n" + "="*50)
#     print("PIPELINE COMPLETE!")
#     print("="*50)

# if __name__ == "__main__":
#     asyncio.run(main())