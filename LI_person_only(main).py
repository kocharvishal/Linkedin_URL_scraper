import os
import time
import random
import pandas as pd
from playwright.sync_api import sync_playwright

# CONFIGURATION
EXCEL_PATH = r"D:\Tasks\linkedin_automation\18june.xlsx"
LINKEDIN_EMAIL = "Siddharth.sanghvi@proton.me"
LINKEDIN_PASSWORD = "Generic1!"


def linkedin_login(page):
    page.goto("https://www.linkedin.com/login")
    page.fill('input#username', LINKEDIN_EMAIL)
    page.fill('input#password', LINKEDIN_PASSWORD)
    page.click('button[type=\"submit\"]')
    page.wait_for_selector('input[placeholder*="Search"]', timeout=20000)

def search_personal_linkedin_url(page, first_name, last_name, company):
    search_query = f'{first_name} {last_name} {company}'.strip()
    print(f"[DEBUG] Search query: {search_query}")
    search_input = page.wait_for_selector('input[placeholder*="Search"]', timeout=15000)
    search_input.fill(search_query)
    search_input.press('Enter')
    page.wait_for_selector('div.search-results-container, .search-results__list', timeout=10000)
    # Click "People" filter
    try:
        people_button = page.wait_for_selector('//button[contains(@aria-label, "People")]', timeout=4000)
        people_button.click()
        page.wait_for_selector('div.search-results-container, .search-results__list', timeout=6000)
    except Exception as e:
        print(f"[DEBUG] People filter not found or error: {e}")
        pass
    # Get the first profile link
    try:
        profile_link = page.wait_for_selector('//a[contains(@href, "/in/")]', timeout=4000)
        profile_url = profile_link.get_attribute('href').split('?')[0]
        print(f"[DEBUG] Found profile URL: {profile_url}")
        return profile_url
    except Exception as e:
        print(f"[DEBUG] No profile link found or error: {e}")
        return ""

def main():
    print(f"[DEBUG] Checking if Excel file exists at {EXCEL_PATH}")
    if not os.path.isfile(EXCEL_PATH):
        print(f"ERROR: Excel file not found at {EXCEL_PATH}")
        return

    df = pd.read_excel(EXCEL_PATH, engine='openpyxl')
    print(f"[DEBUG] DataFrame loaded with shape: {df.shape}")
    print(f"[DEBUG] DataFrame columns: {df.columns.tolist()}")
    if 'Found LinkedIn URL' not in df.columns:
        df['Found LinkedIn URL'] = ""
        print("[DEBUG] Added 'Found LinkedIn URL' column to DataFrame.")

    print("[DEBUG] Starting Playwright...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        linkedin_login(page)

        print("[DEBUG] Entering row processing loop...")
        for idx, row in df.iterrows():
            print(f"[DEBUG] Processing row {idx}: {row.to_dict()}")
            url_field = row.get('Found LinkedIn URL', '')
            if pd.isna(url_field) or str(url_field).strip() == "" or "here" in str(url_field).lower():
                first_name = str(row['First name']).strip()
                last_name = str(row['Last name']).strip()
                company = str(row.get('Company name', '')).strip()
                print(f"[Personal] Searching for: {first_name} {last_name} {company}")
                url = search_personal_linkedin_url(page, first_name, last_name, company)
                if url:
                    print(f"[Personal] URL found: {url}")
                    df.at[idx, 'Found LinkedIn URL'] = url
                else:
                    print("[Personal] URL not found.")
                    df.at[idx, 'Found LinkedIn URL'] = ""
                time.sleep(random.uniform(0.5, 1.2))
            else:
                print(f"[DEBUG] Skipping row {idx} as URL already present: {url_field}")

        context.close()
        browser.close()

    print("[DEBUG] Saving DataFrame to Excel...")
    df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
    print("Done! Results saved to Excel.")

if __name__ == "__main__":
    main() 