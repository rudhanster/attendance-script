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
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException

# =============================
# Config
# =============================
file_path = "./attendance.xlsx"  # Excel file name

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

# Keep subject_code for fallback search/compat
subject_code = course_code

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
# options.add_argument("--headless=new")  # uncomment to run headless if needed

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

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

# ---------- SMART EVENT TILE MATCHING (stale-safe) ----------
def _norm(s: str) -> str:
    return " ".join((s or "").split())

def matches_event_text(txt: str, code: str, sem: str, sec: str, sess: str) -> bool:
    """
    Example event text:
    "CSE 3142 - CSE 3142 - OPERATING SYSTEMS LAB - 905 - Semester V: Program Sec B-1."
    Match on:
      - course code -> "CSE 3142"
      - semester    -> "Semester V"
      - section     -> "Sec B" or "Section B"
      - session     -> optional: "B-1" (Sec-section style) or "Session 1"
    """
    t = _norm(txt)
    T = t.upper()

    ok = True
    if code:
        ok = ok and (code.upper() in T)
    if sem:
        ok = ok and bool(re.search(rf"\bSEMESTER\s*{re.escape(sem)}\b", T, flags=re.I))
    if sec:
        ok = ok and bool(re.search(rf"\bSEC(?:TION)?\s*{re.escape(sec)}\b", T, flags=re.I))
    if sess and sess.strip():
        sess = sess.strip()
        sess_ok = False
        # "B-1" (or "B - 1")
        if sec and re.search(rf"\b{re.escape(sec)}\s*-\s*{re.escape(sess)}\b", T, flags=re.I):
            sess_ok = True
        # "Session 1"
        if not sess_ok and re.search(rf"\bSESSION\s*{re.escape(sess)}\b", T, flags=re.I):
            sess_ok = True
        ok = ok and sess_ok
    return ok

def collect_tiles_pairs():
    """Return list of (webelement, text) pairs; ignore stale items."""
    pairs = []
    els = driver.find_elements(By.XPATH, "//a[contains(@role,'button')]")
    for el in els:
        try:
            txt = el.get_attribute("innerText") or el.text or ""
            pairs.append((el, txt))
        except StaleElementReferenceException:
            continue
    return pairs

# Try multiple passes to avoid stale references right after the date render
candidates = []
for attempt in range(5):
    pairs = collect_tiles_pairs()
    # strict match first
    candidates = [el for (el, txt) in pairs if matches_event_text(txt, course_code, semester, class_section, session_no)]
    if candidates:
        break
    # fallback to course code only
    if not candidates and course_code:
        for (el, txt) in pairs:
            if (course_code or "").upper() in (txt or "").upper():
                candidates.append(el)
        if candidates:
            break
    time.sleep(0.3)  # tiny backoff to let DOM settle

if not candidates:
    print("⚠️ No event matched the criteria. Available tiles were:")
    for (_, txt) in collect_tiles_pairs():
        print(" -", _norm(txt))
    raise RuntimeError("❌ Could not find a matching subject tile for the selected date.")

# Click the first matched candidate; refetch if it goes stale during click
clicked = False
for _ in range(3):
    try:
        event_tile = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(candidates[0]))
        js_click(driver, event_tile)
        clicked = True
        break
    except StaleElementReferenceException:
        fresh = []
        for (el, txt) in collect_tiles_pairs():
            if matches_event_text(txt, course_code, semester, class_section, session_no) or (
                course_code and (course_code.upper() in (txt or "").upper())
            ):
                fresh.append(el)
        if fresh:
            candidates[0] = fresh[0]
        time.sleep(0.2)

if not clicked:
    raise RuntimeError("❌ Found a matching tile but failed to click it due to staleness.")

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
# 5) Untick Absentees + console summary (stale-safe per absentee)
# =============================
print("🔎 Searching for each absentee ID on page...")
unticked_ids = []
not_found = []

def untick_absentee_once(ab):
    """Find row+checkbox fresh and untick if selected. Returns True if unticked, False if already unticked."""
    cell = WebDriverWait(driver, 8).until(
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
            # Table likely re-rendered; retry with fresh lookup
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
# 7) Done + developer credit
# =============================
print("🎉 Attendance marking complete!")
time.sleep(1.5)
driver.quit()

print("\n====================================================")
print("👨‍💻 Developed by: Anirudhan Adukkathayar C, SCE, MIT")
print("====================================================\n")

