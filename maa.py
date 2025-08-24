import sys
import os
import time
import re
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import (
    StaleElementReferenceException,
    TimeoutException,
)
import json
# =============================
# Excel path resolver (with UI picker + persisted config)
# =============================
import json

# Support both normal script runs and interactive runs with no __file__
BASE_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
CONFIG_FILE = os.path.join(BASE_DIR, "attendance_config.json")

def load_saved_excel_path():
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                p = data.get("excel_path")
                if p and os.path.exists(p):
                    return p
    except Exception:
        pass
    return None

def save_excel_path(path):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"excel_path": path}, f, ensure_ascii=False, indent=2)
        print(f"💾 Saved Excel path to {CONFIG_FILE}")
    except Exception as e:
        print(f"⚠️ Could not save config: {e}")

def pick_excel_via_ui():
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)  # bring dialog to front
        while True:
            path = filedialog.askopenfilename(
                title="Select Attendance Excel (.xlsx)",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            if not path:
                messagebox.showinfo("Cancelled", "No file selected. Exiting.")
                return None
            if not path.lower().endswith(".xlsx"):
                messagebox.showerror("Invalid file", "Please select a .xlsx file.")
                continue
            # quick validation: required sheets
            try:
                xl = pd.ExcelFile(path)
                sheets = set(xl.sheet_names)
                needed = {"Attendance", "Initial Setup"}
                if not needed.issubset(sheets):
                    messagebox.showerror(
                        "Missing Sheets",
                        "Selected file must contain sheets: 'Attendance' and 'Initial Setup'."
                    )
                    continue
            except Exception as e:
                messagebox.showerror("Error", f"Could not read Excel file:\n{e}")
                continue
            return path
    except Exception:
        # Tk not available – fallback to console
        print("🪟 Tkinter UI not available. Enter full path to the Excel (.xlsx) file:")
        p = input("> ").strip('"').strip()
        if p and os.path.exists(p) and p.lower().endswith(".xlsx"):
            return p
        return None

def resolve_excel_path(default_name="./attendance.xlsx"):
    # 1) default in working dir
    if os.path.exists(default_name):
        return default_name
    # 2) previously saved
    saved = load_saved_excel_path()
    if saved:
        print(f"📂 Using saved Excel path: {saved}")
        return saved
    # 3) UI picker
    print(f"❌ Excel not found at: {default_name}")
    print("📁 Please select your attendance Excel file…")
    picked = pick_excel_via_ui()
    if not picked:
        print("❌ No Excel selected. Exiting.")
        sys.exit(1)
    save_excel_path(picked)
    return picked

# ------ NOW resolve and use the path ------
file_path = resolve_excel_path("./attendance.xlsx")
print(f"🗂️  Excel file: {file_path}")


# =============================
# Config
# =============================


HOME_URL  = "https://maheslcmtech.lightning.force.com/lightning/page/home"
BASE_URL  = "https://maheslcmtech.lightning.force.com"
LOGIN_URL = "https://maheslcm.manipal.edu/login"

# =============================
# Helpers
# =============================
def js_click(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    driver.execute_script("arguments[0].click();", el)

def click_calendar_date_fast(driver, day_number: str):
    """
    Click the mini-calendar date using direct JS (fast).
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
        if (span) el = span.closest('a, button, [role="tab"]') || span;
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

def _norm(s: str) -> str:
    return " ".join((s or "").split())

def matches_event_text(txt: str, code: str, sem: str, sec: str, sess: str) -> bool:
    """
    Example event text:
    "CSE 3142 - CSE 3142 - OPERATING SYSTEMS LAB - 905 - Semester V: Program Sec B-1."
    We match:
      - course code -> "CSE 3142"
      - semester    -> "Semester V"
      - section     -> "Sec B" or "Section B"
      - session     -> optional: "B-1" or "Session 1"
    """
    T = _norm(txt).upper()
    ok = True
    if code:
        ok = ok and (code.upper() in T)
    if sem:
        ok = ok and (f"SEMESTER {sem.upper()}" in T)
    if sec:
        # allow "SEC B" or "SECTION B"
        ok = ok and (f"SEC {sec.upper()}" in T or f"SECTION {sec.upper()}" in T)
    if sess and sess.strip():
        sess = sess.strip()
        sess_ok = False
        # "Sec B-1" style
        if sec and f"{sec.upper()}-{sess}".upper() in T:
            sess_ok = True
        # "Session 1" style
        if not sess_ok and f"SESSION {sess}".upper() in T:
            sess_ok = True
        ok = ok and sess_ok
    return ok

def click_event_candidate(driver, candidate):
    """
    Switch into candidate's iframe (if any), re-locate by text/title, then click.
    'candidate' is a dict: {'frame_index': None or int, 'text': '...'}
    """
    driver.switch_to.default_content()
    if candidate['frame_index'] is not None:
        ifr = driver.find_elements(By.TAG_NAME, "iframe")
        if candidate['frame_index'] >= len(ifr):
            return False
        driver.switch_to.frame(ifr[candidate['frame_index']])

    txt = _norm(candidate['text'])
    probe = txt[:80]  # keep shorter for XPath contains()

    locators = [
        (By.XPATH, f"//a[contains(@role,'button') and contains(normalize-space(.), {repr(probe)})]"),
        (By.XPATH, f"//*[@title and contains(normalize-space(@title), {repr(probe)})]"),
        (By.XPATH, f"//*[contains(@class,'slds-truncate') and contains(normalize-space(.), {repr(probe)})]"),
        (By.XPATH, f"//button[contains(normalize-space(.), {repr(probe)})]"),
        (By.XPATH, f"//*[contains(normalize-space(.), {repr(probe)})]"),
    ]
    for by, xp in locators:
        try:
            el = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((by, xp)))
            js_click(driver, el)
            driver.switch_to.default_content()
            return True
        except Exception:
            continue
    driver.switch_to.default_content()
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
# 2) Load Excel file + Initial Setup fields
# =============================
attendance_df = pd.read_excel(file_path, sheet_name="Attendance", header=1)
setup_df = pd.read_excel(file_path, sheet_name="Initial Setup", header=None)

# Extract values from Initial Setup (Column B values on rows 1..5)
course_name   = str(setup_df.iloc[0, 1]).strip() if len(setup_df) > 0 else ""
course_code   = str(setup_df.iloc[1, 1]).strip() if len(setup_df) > 1 else ""
semester      = str(setup_df.iloc[2, 1]).strip() if len(setup_df) > 2 else ""
class_section = str(setup_df.iloc[3, 1]).strip() if len(setup_df) > 3 else ""
session_no    = str(setup_df.iloc[4, 1]).strip() if len(setup_df) > 4 else ""

print("\n📘 Course Details from Initial Setup:")
print(f"   Course Name   : {course_name}")
print(f"   Course Code   : {course_code}")
print(f"   Semester      : {semester}")
print(f"   Class Section : {class_section}")
print(f"   Session       : {session_no}\n")

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
PROFILE_DIR = os.path.abspath("./slcm_automation_profile")  # dedicated reusable profile
os.makedirs(PROFILE_DIR, exist_ok=True)

options = webdriver.ChromeOptions()
options.add_argument(f"--user-data-dir={PROFILE_DIR}")
options.add_argument("--no-first-run")
options.add_argument("--no-default-browser-check")
# options.add_argument("--headless=new")  # keep visible for SSO/Lightning

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# --- Bootstrap / SSO ---
if not hard_nav(driver, HOME_URL):
    hard_nav(driver, BASE_URL)
    hard_nav(driver, HOME_URL)

cur = driver.current_url.lower()
print("🌐 After bootstrap:", cur)

if ("login.microsoftonline.com" in cur) or ("saml" in cur) or ("manipal.edu" in cur and "/login" in cur):
    print("🔐 SSO/login detected. Complete it in the opened Chrome window.")
    input("Press Enter here AFTER you reach Salesforce Home... ")
    hard_nav(driver, HOME_URL)

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

# Click date (fast)
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "calendarSidebar")))
time.sleep(0.15)
day_number = str(selected_date.day).lstrip("0")
click_calendar_date_fast(driver, day_number)

# --- Search all iframes for subject tiles by Course Code + Semester + Section + Session ---
driver.switch_to.default_content()
candidates = []
iframes = driver.find_elements(By.TAG_NAME, "iframe")

def scan_in_context(ctx_driver, frame_index):
    found = []
    # broader sweep: role=button links, buttons, nodes with title, or truncation class
    nodes = ctx_driver.find_elements(By.XPATH, "//a[contains(@role,'button')] | //button | //*[@title] | //*[contains(@class,'slds-truncate')]")
    for el in nodes:
        try:
            txt = (el.get_attribute("innerText") or el.get_attribute("title") or el.text or "").strip()
            if not txt:
                continue
            if matches_event_text(txt, course_code, semester, class_section, session_no):
                found.append({"frame_index": frame_index, "text": txt})
        except StaleElementReferenceException:
            continue
    return found

# top document first
candidates.extend(scan_in_context(driver, None))
# then each iframe
for idx in range(len(iframes)):
    try:
        driver.switch_to.frame(iframes[idx])
        candidates.extend(scan_in_context(driver, idx))
    except Exception:
        pass
    finally:
        driver.switch_to.default_content()

if not candidates:
    # fallback: accept course code only (if strict filters missed due to wording)
    for idx in [None] + list(range(len(iframes))):
        try:
            driver.switch_to.default_content()
            if idx is not None:
                driver.switch_to.frame(iframes[idx])
            nodes = driver.find_elements(By.XPATH, "//*[normalize-space(@title) or normalize-space(text())]")
            for el in nodes:
                try:
                    txt = (el.get_attribute("innerText") or el.get_attribute("title") or el.text or "").strip()
                    if not txt:
                        continue
                    if course_code and course_code.upper() in _norm(txt).upper():
                        candidates.append({"frame_index": idx, "text": txt})
                except StaleElementReferenceException:
                    continue
        except Exception:
            pass

driver.switch_to.default_content()

if not candidates:
    # brief dump to help diagnose
    print("⚠️ No event matched the criteria. Could not find a tile for the selected date.")
    driver.quit()
    sys.exit(1)

print("🎯 Found candidate(s):")
for c in candidates:
    where = "top" if c['frame_index'] is None else f"iframe#{c['frame_index']}"
    print(f" - [{where}] {c['text']}")

# Click first candidate safely
clicked = False
for cand in candidates:
    if click_event_candidate(driver, cand):
        clicked = True
        break
    time.sleep(0.25)
if not clicked:
    driver.quit()
    raise RuntimeError("❌ Could not click any candidate event tile")

# "More Details" if a popover appears; otherwise Lightning may navigate directly
try:
    more_details = WebDriverWait(driver, 6).until(
        EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='More Details']"))
    )
    js_click(driver, more_details)
except Exception:
    pass  # direct navigation case

# Open Attendance tab
click_attendance_tab_fast(driver)

# =============================
# 5) Untick Absentees (stale-safe per absentee) + summary
# =============================
print("🔎 Searching for each absentee ID on page...")
unticked_ids = []
not_found = []

def untick_absentee_once(ab):
    cell = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, f"//lightning-base-formatted-text[normalize-space()='{ab}']"))
    )
    row = cell.find_element(By.XPATH, "./ancestor::tr")
    checkbox = row.find_element(By.XPATH, ".//input[@type='checkbox']")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", checkbox)
    if checkbox.is_selected():
        js_click(driver, checkbox)
        return True
    return False

for ab in absentees:
    success = False
    attempts = 0
    while attempts < 4 and not success:
        try:
            if untick_absentee_once(ab):
                print(f"✔️ Unticked absentee: {ab}")
                unticked_ids.append(ab)
            else:
                print(f"ℹ️ Already unticked: {ab}")
            success = True
        except (StaleElementReferenceException, TimeoutException):
            attempts += 1
            time.sleep(0.3)
        except Exception:
            break
    if not success:
        print(f"❌ Not found on page: {ab}")
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
    submit_btn = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Submit Attendance')]"))
    )
    js_click(driver, submit_btn)
    print("✅ Clicked Submit Attendance")

    # Wait for modal to render & be visible
    modal = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'modal-container') or contains(@class,'uiModal') or contains(@class,'slds-modal')]"))
    )
    WebDriverWait(driver, 10).until(EC.visibility_of(modal))
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
            confirm_btn = WebDriverWait(modal, 5).until(EC.element_to_be_clickable((By.XPATH, xp)))
            js_click(driver, confirm_btn)
            print("✅ Clicked Confirm")
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
# 7) Done + credit
# =============================
print("🎉 Attendance marking complete!")
time.sleep(1.5)
driver.quit()

print("\n====================================================")
print("👨‍💻 Developed by: Anirudhan Adukkathayar C, SCE, MIT")
print("====================================================\n")



