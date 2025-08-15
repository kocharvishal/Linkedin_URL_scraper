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
    page.click('button[type="submit"]')
    page.wait_for_selector('input[placeholder*="Search"]', timeout=20000)

def search_personal_linkedin_url(page, first_name, last_name, company):
    search_query = f'{first_name} {last_name} {company}'.strip()
    search_input = page.wait_for_selector('input[placeholder*="Search"]', timeout=15000)
    search_input.fill(search_query)
    search_input.press('Enter')
    page.wait_for_selector('div.search-results-container, .search-results__list', timeout=10000)
    # Click "People" filter
    try:
        people_button = page.wait_for_selector('//button[contains(@aria-label, "People")]', timeout=4000)
        people_button.click()
        page.wait_for_selector('div.search-results-container, .search-results__list', timeout=6000)
    except Exception:
        pass
    # Get the first profile link
    try:
        profile_link = page.wait_for_selector('//a[contains(@href, "/in/")]', timeout=4000)
        profile_url = profile_link.get_attribute('href').split('?')[0]
        return profile_url
    except Exception:
        return ""

def search_company_linkedin_url(page, company_name):
    search_input = page.wait_for_selector('input[placeholder*="Search"]', timeout=15000)
    search_input.fill(company_name)
    search_input.press('Enter')
    page.wait_for_selector('div.search-results-container, .search-results__list', timeout=10000)
    # Click "Companies" filter
    try:
        companies_button = page.wait_for_selector('//button[contains(@aria-label, "Companies")]', timeout=4000)
        companies_button.click()
        page.wait_for_selector('div.search-results-container, .search-results__list', timeout=6000)
    except Exception:
        pass
    # Get the first company page link
    try:
        company_link = page.wait_for_selector('//a[contains(@href, "/company/")]', timeout=4000)
        url = company_link.get_attribute('href').split('?')[0]
        if not url.startswith('https://'):
            url = 'https://www.linkedin.com' + url
        return url
    except Exception:
        return ""

def main():
    if not os.path.isfile(EXCEL_PATH):
        print(f"ERROR: Excel file not found at {EXCEL_PATH}")
        return

    df = pd.read_excel(EXCEL_PATH, engine='openpyxl')
    if 'Found LinkedIn URL' not in df.columns:
        df['Found LinkedIn URL'] = ""
    if 'Founded Company LinkedIn URL' not in df.columns:
        df['Founded Company LinkedIn URL'] = ""

    # Prepare company cache
    company_names = df['Company name'].astype(str).str.strip().unique()
    company_cache = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        linkedin_login(page)

        # --- Personal LinkedIn URL Extraction ---
        for idx, row in df.iterrows():
            url_field = str(row.get('Found LinkedIn URL', '')).lower()
            if pd.isna(url_field) or url_field.strip() == "" or "here" in url_field:
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

        # --- Company LinkedIn URL Extraction ---
        for company in company_names:
            if not company or pd.isna(company):
                company_cache[company] = ""
                continue
            print(f"[Company] Searching for company: {company}")
            url = search_company_linkedin_url(page, company)
            company_cache[company] = url
            print(f"[Company] URL found: {url}")
            time.sleep(0.5)

        context.close()
        browser.close()

    # Fill the company column in the DataFrame
    df['Founded Company LinkedIn URL'] = df['Company name'].astype(str).str.strip().map(company_cache)
    df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
    print("Done! Results saved to Excel.")

if __name__ == "__main__":
    main() 