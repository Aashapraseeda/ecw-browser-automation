
import asyncio
import os
from dotenv import load_dotenv
from playwright.async_api import async_playwright

load_dotenv()

USERNAME = os.getenv("ECW_USERNAME")
PASSWORD = os.getenv("ECW_PASSWORD")
print("USERNAME =", USERNAME)
print("PASSWORD =", PASSWORD)

async def login_ecw():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=500)
        context = await browser.new_context()
        page = await context.new_page()

        print("Opening eCW login page...")
        await page.goto("https://txsnmbapp.ecwcloud.com/mobiledoc/jsp/webemr/login/newLogin.jsp")
        print("Login page loaded")

        # Step 1 - Enter username and click Next
        await page.get_by_role("textbox", name="Enter username to continue").fill(USERNAME)
        await page.get_by_role("button", name="Next").click()
        print("Username entered and Next clicked!")

        # Step 2 - Enter password and click Log In
        await page.get_by_role("textbox", name="Enter Password to continue").fill(PASSWORD)
        print("Password entered!")
        await page.get_by_role("button", name="Log In").click()
        print("Log In clicked - waiting for home page...")

        # Wait for home page to fully load
        await page.wait_for_selector('#jellybean-panelLink33', timeout=120000)
        print("Home page loaded!")

        # Wait for loading overlay to hide
        try:
            await page.wait_for_selector('#load', state='hidden', timeout=120000)
            print("ECW fully loaded!")
        except:
            print("Loading overlay already hidden!")

        # --- HANDLE LICENSE ALERT ---
        print("Checking for License Alert...")
        try:
            await page.wait_for_selector('#providerLicense', state='visible', timeout=15000)
            print("License Alert detected!")
            await page.click('#providerLicense button.clsMyButton')
            print("License Alert dismissed!")
            await asyncio.sleep(2)
        except:
            print("No License Alert, continuing...")

        # --- NAVIGATE TO REFERRALS INCOMING ---
        print("Navigating to Referrals...")
        await page.click('#jellybean-panelLink33')
        await asyncio.sleep(4)
        print("Referrals page loaded!")

        # --- CLEAR AND SET ASSIGNED TO FIELD ---
        print("Setting Assigned To - All Providers and Staff...")
        assigned_input = page.locator('#staff-lookup_staff_Ipt1')
        await assigned_input.click()
        await page.keyboard.press('Control+A')
        await page.keyboard.press('Backspace')
        await assigned_input.fill('All Providers and Staff')
        await page.keyboard.press('Enter')
        print("Selected All Providers and Staff!")
        await asyncio.sleep(2)

        # --- CLICK FILTER BUTTON ---
        print("Applying filter...")
        await page.click('#referralBtn7')
        print("Filter applied!")

        await asyncio.sleep(5)
        print("All done! Browser staying open...")

        await asyncio.sleep(99999)

asyncio.run(login_ecw())