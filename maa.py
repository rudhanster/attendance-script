import sys
import os
import time
import pandas as pd
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# =============================
# Config
# =============================
file_path = "./OS V attendance.xlsx"  # Excel file name

HOME_URL  = "https://maheslcmtech.lightning.force.com/lightning/page/home"
BASE_URL  = "https://maheslcmtech.lightning.force.com"
LOGIN_URL = "https://maheslcm.manipal.edu/login"

# =============================
# Fast helpers
# =============================
def js_click(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    driver.execute_script("arguments[0].click();", el)

def click_calendar_date_fast(driver, day_number: str):
    """
    Click the mini-calendar date using direct JS (much faster than repeated waits).
    Requires Calendar sidebar to be present.
    """
    js = """
    const wrap = document.querySelector('#calendarSidebar');
    if (!wrap) return false;
    const dayNodes = wrap.querySelectorAll('table.datepicker .slds-day, .slds-day');
    for (const n of dayNodes) {
        const txt = (n.textContent || '').trim();
        const disabled = n.getAttribute('aria-disabled') === 'true' || (n.className || '').includes('disabled');
        if (!disabled && txt === arguments[0]) {
            n.scrollIntoView({block:'center'});
            n.click();
            return true;
        }
    }
    return false;
    """
    ok = driver.execute_script(js, day_number)
    if not ok:
        raise RuntimeError(f"❌ Could not click mini calendar date {day_number}")
    print(f"✅ Clicked calendar date (fast): {day_number}")

def click_attendance_tab_fast(driver):
    """
    Open the Attendance tab via direct JS (fast), with a short Selenium fallback.
    """
    js = """
    let el = document.querySelector("a[data-label='Attendance']");
    if (!el) {
        const span = Array.from(document.querySelectorAll('span.title'))
            .find(s => (s.textContent || '').trim() === 'Attendance');
        if (span) el = span.closest('a, button, [role=\"tab\"]') || span;
    }
    if (el) {
        el.scrollIntoView({block:'center'});
        el.click();
        return true;
    }
    return false;
    """
    ok = driver.execute_script(js)
    if ok:
        print("✅ Opened Attendance tab (fast)")
        return
    # short fallback if JS missed due to render timing
    att_tab = WebDriverWait(driver, 8).until(
        EC.element_to_be_clickable((By.XPATH, "//a[@data-label='Attendance'] | //span[@class='title' and normalize-space()='Attendance']"))
    )
    js_click(driver, att_tab)
    print("✅ Opened Attendance tab (fallback)")

def ready(driver):
    try:
        return driver.execute_script("return document.readyState") == "complete"
    except Exception:
        return False

def close_blank_tabs(driver):
    handles = driver.window_handles[:]
    for h in handles:
        driver.switch_to.window(h)
        url = driver.current_url
        if url.startswith(("about:blank", "chrome://newtab", "chrome://")):
            try: driver.close()
            except Exception: pass
    if driver.window_handles:
        driver.switch_to.window(driver.window_handles[-1])

def hard_nav(driver, url, attempts=4):
    for _ in range(attempts):
        try:
            driver.get(url); time.sleep(0.5)
            if ready(driver) and driver.current_url.startswith("http"):
                return True
        except Exception:
            pass
        try:
            driver.execute_script("window.location.href = arguments[0];", url); time.sleep(0.5)
            if ready(driver) and driver.current_url.startswith("http"):
                return True
        except Exception:
            pass
        try:
            driver.switch_to.new_window('tab')
            driver.get(url); time.sleep(0.7)
            if ready(driver) and driver.current_url.startswith("http"):
                close_blank_tabs(driver)
                return True
        except Exception:
            pass
        time.sleep(0.3)
    close_blank_tabs(driver)
    return False

# =============================
# 1) Parse date argument (d/m/Y) or use today's date
# =============================
if len(sys.argv) > 1:
    selected_date = datetime.strptime(sys.argv[1], "%d/%m/%Y").date()
    print(f"📅 Using date: {selected_date} (from argument)")
else:
    selected_date = datetime.today().date()
    print(f"📅 Using date: {selected_date} (today)")

# =============================
# 2) Load Excel file (OS V attendance.xlsx)
# =============================
attendance_df = pd.read_excel(file_path, sheet_name="Attendance", header=1)
setup_df = pd.read_excel(file_path, sheet_name="Initial Setup", header=None)
subject_code = str(setup_df.iloc[1, 1]).strip()

def find_date_column(columns, target_date):
    for col in columns:
        if isinstance(col, datetime) and col.date() == target_date:
            return col
        if isinstance(col, str):
            try:
                if datetime.strptime(col, "%m/%d/%Y").date() == target_date:
                    return col
            except:
                pass
    return None

date_col = find_date_column(attendance_df.columns, selected_date)
if date_col is None:
    raise ValueError("❌ No column found for the specified date")
print(f"✅ Using date column in sheet: {date_col}")

# Extract absentees
reg_no_col = "Reg. No. "
absentees = (
    attendance_df[attendance_df[date_col].astype(str).str.lower() == "ab"][reg_no_col]
    .astype(str)
    .str.split(".")
    .str[0]  # drop any decimals like ".0"
    .tolist()
)
print("Absentees (IDs to untick):", absentees)

# =============================
# 3) Selenium with webdriver-manager (auto ChromeDriver)
# =============================
# You can switch to your real Chrome profile by replacing the profile dir below
PROFILE_DIR = os.path.abspath("./slcm_automation_profile")  # dedicated reusable profile
os.makedirs(PROFILE_DIR, exist_ok=True)

options = webdriver.ChromeOptions()
options.add_argument(f"--user-data-dir={PROFILE_DIR}")
options.add_argument("--no-first-run")
options.add_argument("--no-default-browser-check")
# options.add_argument("--headless=new")  # uncomment to run headless if needed

# 👉 webdriver-manager auto-installs the matching ChromeDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# --- Bootstrap: force navigation & handle SSO once ---
if not hard_nav(driver, HOME_URL):
    hard_nav(driver, BASE_URL)
    hard_nav(driver, HOME_URL)

cur = driver.current_url.lower()
print("🌐 After bootstrap:", cur)

if ("login.microsoftonline.com" in cur) or ("saml" in cur) or ("manipal.edu" in cur and "/login" in cur):
    print("🔐 SSO/login detected. Complete it in the opened Chrome window.")
    input("Press Enter here AFTER you reach Salesforce Home... ")
    hard_nav(driver, HOME_URL)

# Verify we’re on Lightning Home
WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//a[@title='Calendar']")))
print("✅ Logged in & on Lightning Home")

# =============================
# 4) Calendar → date → subject → Attendance
# =============================
# Open Calendar tab
cal_tab = WebDriverWait(driver, 40).until(
    EC.element_to_be_clickable((By.XPATH, "//a[@title='Calendar']"))
)
js_click(driver, cal_tab)

# Wait for sidebar calendar (short) and click date FAST
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "calendarSidebar")))
time.sleep(0.15)
day_number = str(selected_date.day).lstrip("0")
click_calendar_date_fast(driver, day_number)

# Click subject event tile
event_xpath = f"//a[contains(@role,'button') and contains(., '{subject_code}')]"
event_tile = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, event_xpath))
)
js_click(driver, event_tile)

# Click "More Details" if popover appears; otherwise Lightning may navigate directly
try:
    more_details = WebDriverWait(driver, 6).until(
        EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='More Details']"))
    )
    js_click(driver, more_details)
except Exception:
    pass  # direct navigation case

# Open Attendance tab FAST
click_attendance_tab_fast(driver)

# =============================
# 5) Untick Absentees + console summary
# =============================
print("🔎 Searching for each absentee ID on page...")
unticked_ids = []
not_found = []

for ab in absentees:
    try:
        cell = WebDriverWait(driver, 6).until(
            EC.presence_of_element_located((By.XPATH, f"//lightning-base-formatted-text[normalize-space()='{ab}']"))
        )
        row = cell.find_element(By.XPATH, "./ancestor::tr")
        checkbox = row.find_element(By.XPATH, ".//input[@type='checkbox']")
        if checkbox.is_selected():
            js_click(driver, checkbox)
            print(f"✔️ Unticked absentee: {ab}")
            unticked_ids.append(ab)
        else:
            print(f"ℹ️ Already unticked: {ab}")
    except Exception:
        not_found.append(ab)

# --- Final summary in console ---
print("\n📊 Attendance Summary")
print(f"✔️ Successfully unticked: {len(unticked_ids)}")
print(f"❌ Not unticked (not found on page): {len(not_found)}")
if not_found:
    print("👉 IDs not unticked:")
    for nf in not_found:
        print(f"   - {nf}")

# =============================
# 6) Submit & Confirm (robust)
# =============================
try:
    submit_btn = WebDriverWait(driver, 12).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Submit Attendance')]"))
    )
    js_click(driver, submit_btn)
    print("✅ Clicked Submit Attendance")

    # Wait for modal to render & be visible
    modal = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'modal-container') or contains(@class,'uiModal') or contains(@class,'slds-modal')]"))
    )
    WebDriverWait(driver, 8).until(EC.visibility_of(modal))
    print("✅ Confirmation modal visible")

    candidate_xpaths = [
        ".//button[normalize-space()='Confirm Submission']",
        ".//button[.//span[normalize-space()='Confirm Submission']]",
        ".//button[contains(.,'Confirm Submission')]",
        ".//footer//*[self::button or self::*[contains(@class,'slds-button')]][contains(.,'Confirm') and contains(@class,'slds-button_brand')]",
        ".//button[contains(.,'Confirm') and contains(@class,'slds-button_brand')]",
    ]

    confirm_clicked = False
    for xp in candidate_xpaths:
        try:
            confirm_btn = WebDriverWait(modal, 4).until(EC.element_to_be_clickable((By.XPATH, xp)))
            js_click(driver, confirm_btn)
            print(f"✅ Clicked Confirm via locator: {xp}")
            confirm_clicked = True
            break
        except Exception:
            continue

    if not confirm_clicked:
        confirm_btn = driver.execute_script(
            """
            const modal = document.querySelector('.modal-container, .uiModal, .slds-modal');
            if (!modal) return null;
            const btns = Array.from(modal.querySelectorAll('button, .slds-button'));
            const norm = t => (t || '').trim().toLowerCase();
            return btns.find(b => {
              const txt = norm(b.innerText || b.textContent);
              return txt === 'confirm submission' || txt === 'confirm' || txt.includes('confirm submission');
            }) || null;
            """
        )
        if confirm_btn:
            js_click(driver, confirm_btn)
            print("✅ Clicked Confirm via JS fallback")
        else:
            try:
                modal.send_keys(Keys.ENTER)
                print("↩️ Sent ENTER to modal (fallback)")
            except Exception:
                print("⚠️ Could not click Confirm Submission. Please click it manually.")

except Exception as e:
    print("⚠️ Could not confirm submission:", e)

# =============================
# 7) Done + developer credit
# =============================
print("🎉 Attendance marking complete!")
time.sleep(1.5)
driver.quit()

print("\n====================================================")
print("👨‍💻 Developed by: Anirudhan Adukkathayar C, SCE, MIT")
print("====================================================\n")

