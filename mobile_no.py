import os
import re
import sys
import time
import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# ==========================================
# CREDENTIALS REFERENCE (Copy & Paste here)
# ==========================================
# Portal Login (saccess.nic.in):
#   User: debashis.shivamangal@odisha.gov.in
#   Pass: Debashis@123
#
# LRMS Portal Login:
#   User: 1104-ri11
#   Pass: Sdpur@432
# ==========================================

# 1. File & Delay Configurations
EXCEL_PATH = "cases.xlsx"
SHEET_NAME = "Sheet2"      # Target sheet containing your case list
DELAY_AFTER_TYPING = 1.5   # Delay (seconds) after entering case number
DELAY_AFTER_CLICK = 2.5    # Delay (seconds) for page rendering after clicking Go

def save_dataframe_to_sheet(df, file_path, sheet_name):
    """Safely updates a specific sheet inside an Excel workbook."""
    with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

# 2. Configure Chrome
chrome_options = Options()
chrome_options.add_argument("--disable-infobars")
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--start-maximized")

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 15)

try:
    # 3. Open portal and authenticate manually
    driver.get("https://saccess.nic.in")

    print("\n" + "=" * 65)
    print("ACTION REQUIRED:")
    print("1. Login to saccess.nic.in using:")
    print("   User: debashis.shivamangal@odisha.gov.in | Pass: Debashis@123")
    print("2. Open and log in to LRMS Odisha using:")
    print("   User: 1104-ri11 | Pass: Sdpur@432")
    print("3. Navigate to the 'View Order Sheet' page.")
    print("=" * 65)
    input("\nOnce you are on the Order Sheet page, press Enter here to start...")

    # 4. Locate and focus the LRMS tab
    lrms_handle = None
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        if "lrmsodisha" in driver.current_url.lower():
            lrms_handle = handle
            break

    if not lrms_handle:
        lrms_handle = driver.window_handles[-1]
        driver.switch_to.window(lrms_handle)

    print(f"\nLocked onto target tab: {driver.current_url}")

    # 5. Read target sheet from Excel
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
    df.columns = df.columns.astype(str).str.strip()

    # Identify Case Number column
    case_col = None
    for col in df.columns:
        clean_col = re.sub(r"[^a-zA-Z0-9]", "", col).lower()
        if clean_col in ["caseno", "casenumber", "case"]:
            case_col = col
            break

    if not case_col:
        print(f"\n[ERROR] Could not find a Case Number column in '{SHEET_NAME}'.")
        print(f"Available columns found: {list(df.columns)}")
        sys.exit(1)

    print(f"Using column '{case_col}' for Case Numbers.")

    # Identify or initialize Mobile Number column
    mobile_col = None
    for col in df.columns:
        clean_col = re.sub(r"[^a-zA-Z0-9]", "", col).lower()
        if clean_col in ["mobileno", "mobilenumber", "mobile", "contactno"]:
            mobile_col = col
            break

    if not mobile_col:
        mobile_col = "Mobile No."
        df[mobile_col] = ""

    df[mobile_col] = df[mobile_col].astype(str)

    # 6. Process each Case Number
    for idx, row in df.iterrows():
        case_no = str(row[case_col]).strip()

        # Skip empty rows
        if not case_no or case_no.lower() == "nan":
            continue

        existing_mobile = str(row.get(mobile_col, "")).strip()
        if existing_mobile and existing_mobile.lower() not in ["nan", "", "not found", "error"]:
            print(f"[{idx + 1}/{len(df)}] Skipping {case_no} (already has {existing_mobile})")
            continue

        print(f"\n[{idx + 1}/{len(df)}] Processing Case: {case_no}")

        try:
            # Step A: Enter Case Number
            input_box = wait.until(
                EC.presence_of_element_located(
                    (By.ID, "ctl00_ContentPlaceHolder1_txtIp1")
                )
            )
            input_box.clear()
            input_box.send_keys(case_no)

            time.sleep(DELAY_AFTER_TYPING)

            # Step B: Click Go Button
            search_btn = wait.until(
                EC.element_to_be_clickable(
                    (By.ID, "ctl00_ContentPlaceHolder1_btnSearch")
                )
            )
            search_btn.click()

            time.sleep(DELAY_AFTER_CLICK)

            # Step C: Extract Mobile Number
            mobile_elem = wait.until(
                EC.presence_of_element_located(
                    (By.ID, "ctl00_ContentPlaceHolder1_lblContactNo")
                )
            )
            mobile_no = mobile_elem.text.strip()

            if mobile_no:
                df.at[idx, mobile_col] = mobile_no
                print(f"  --> Found Mobile: {mobile_no}")
            else:
                df.at[idx, mobile_col] = "Not Found"
                print("  --> No mobile number listed on page.")

        except Exception as e:
            print(f"  --> Error fetching Case '{case_no}': {e}")
            df.at[idx, mobile_col] = "Error"

        # Auto-save progress back to Sheet2 every 5 cases
        if (idx + 1) % 5 == 0:
            save_dataframe_to_sheet(df, EXCEL_PATH, SHEET_NAME)
            print("  [Auto-saved progress to Excel]")

    # Final save
    save_dataframe_to_sheet(df, EXCEL_PATH, SHEET_NAME)
    print(f"\nAll cases processed! File saved to '{SHEET_NAME}' in {EXCEL_PATH}")

except SystemExit:
    pass

finally:
    try:
        input("\nPress Enter to close browser session...")
    except Exception:
        pass
    driver.quit()