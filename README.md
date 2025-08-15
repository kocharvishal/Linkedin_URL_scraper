# LinkedIn URL Scraper (All-in-One)

This script automates the process of finding **LinkedIn profile URLs** and **company LinkedIn page URLs** using Playwright and an Excel input file.

It performs:
- Login to LinkedIn
- Profile URL scraping for people (first name + last name + company)
- Company page URL scraping
- Writing results back into the Excel file

---

## 🔧 Requirements

- Python 3.7+
- [Playwright](https://playwright.dev/python/)
- pandas
- openpyxl

Install dependencies:

```bash
pip install pandas openpyxl
pip install playwright
playwright install
```

---

## 📁 Input File Format

The input Excel file (e.g. `18june.xlsx`) must contain at least:

- `First name`
- `Last name`
- `Company name`

The script creates or updates the following columns:
- `Found LinkedIn URL` – personal profile URL
- `Founded Company LinkedIn URL` – official company page URL

---

## 🚀 How to Run

Update the config section in `LI_all_in_one.py` with:
```python
EXCEL_PATH = r"path_to_your_excel_file.xlsx"
LINKEDIN_EMAIL = "your_email"
LINKEDIN_PASSWORD = "your_password"
```

Then run:
```bash
python LI_all_in_one.py
```

Make sure LinkedIn credentials are valid and 2FA is disabled or handled manually.

---

## 🔐 Notes

- **Browser opens in non-headless mode** to bypass bot detection and allow manual steps.
- Avoid scraping too fast – delays are added between searches.
- Do not exceed LinkedIn’s rate limits or TOS.

---

## ✅ Output

After execution, your Excel file will contain the filled-in LinkedIn URLs for each person and their company.

---

## 📄 License

This project is for educational use only. Scraping LinkedIn is subject to their [Terms of Service](https://www.linkedin.com/legal/user-agreement). Use responsibly.
