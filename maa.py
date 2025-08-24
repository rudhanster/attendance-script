import sys
import os
import time
import json
import tempfile
import shutil
import re
import pandas as pd
from datetime import datetime, date
from pathlib import Path

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
    SessionNotCreatedException,
)

# =============================
# Excel path resolver (UI picker + persisted config)
# =============================
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

# ------ Resolve and use the path ------
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
    We match within the event text:
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
        ok = ok and (f"SEC {sec.upper()}" in T or f"SECTION {sec.upper()}" in T)
    if sess and str(sess).strip() and str(sess).strip().lower() != "nan":
        sess = str(sess).strip()
        sess_ok = False
        if sec and f"{sec.upper()}-{sess}".upper() in T:
            sess_ok = True
        if not sess_ok and f"SESSION {sess}".upper() in T:
            sess_ok = True
        ok = ok and sess_ok
    return ok

# --- NEW: down-only scroll & strict day-panel selection ---
def _day_heading_variants(d: date):
    # Salesforce day headings like "Saturday, August 23"
    try:
        return list({
            d.strftime("%A, %B %-d"),
            d.strftime("%A, %B %#d"),
            d.strftime("%A, %B %d"),
            d.strftime("%A, %B %-d, %Y"),
            d.strftime("%A, %B %#d, %Y"),
            d.strftime("%A, %B %d, %Y"),
        })
    except Exception:
        return [d.strftime("%A, %B %d"), d.strftime("%A, %B %d, %Y")]

def disable_auto_scroll(driver):
    driver.execute_script("""
    (function(){
      if (!window.__origScrollIntoView) {
        window.__origScrollIntoView = Element.prototype.scrollIntoView;
        Element.prototype.scrollIntoView = function(){};
      }
      try {
        document.documentElement.style.scrollBehavior = 'auto';
        document.body.style.scrollBehavior = 'auto';
      } catch(e){}
    })();
    """)

def enable_auto_scroll(driver):
    driver.execute_script("""
    (function(){
      if (window.__origScrollIntoView) {
        Element.prototype.scrollIntoView = window.__origScrollIntoView;
        delete window.__origScrollIntoView;
      }
    })();
    """)

def scroll_to_day_panel(driver, target_date, timeout=40):
    labels = _day_heading_variants(target_date)
    container = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".calendarRow.slds-scrollable_y"))
    )
    driver.execute_script("arguments[0].setAttribute('tabindex','0'); arguments[0].focus();", container)

    start = time.time()
    last_top = driver.execute_script("return arguments[0].scrollTop;", container)
    max_top = driver.execute_script("return arguments[0].scrollHeight - arguments[0].clientHeight;", container)
    step = max(220, int(driver.execute_script("return Math.floor(arguments[0].clientHeight*0.9);", container)))

    while time.time() - start < timeout:
        res = driver.execute_script("""
        const cont = arguments[0], labels = arguments[1];
        function t(n){return (n && (n.innerText||n.textContent)||'').trim();}
        function topWithin(node, anc){let y=0, el=node; while(el && el!==anc){y+=el.offsetTop; el=el.offsetParent;} return y;}
        const heads = Array.from(cont.querySelectorAll('h2.slds-assistive-text'));
        let H=null; for(const h of heads){ if(labels.includes(t(h))){H=h; break;} }
        if(!H) return {found:false, cur:cont.scrollTop, max:(cont.scrollHeight-cont.clientHeight)};
        let sib = H.nextElementSibling;
        while(sib && !sib.classList.contains('calendarDay')) sib = sib.nextElementSibling;
        if(!sib) return {found:false, cur:cont.scrollTop, max:(cont.scrollHeight-cont.clientHeight)};
        const targetTop = Math.max(0, topWithin(sib, cont) - 80);
        return {found:true, targetTop:targetTop, cur:cont.scrollTop, max:(cont.scrollHeight-cont.clientHeight)};
        """, container, labels)

        cur = res.get("cur", 0)
        max_top = res.get("max", max_top)

        if res.get("found"):
            tgt = min(res["targetTop"], max_top)
            if tgt < cur:
                tgt = cur   # enforce DOWN-ONLY
            driver.execute_script("arguments[0].scrollTop = arguments[1];", container, tgt)
            time.sleep(0.25)
            cur2 = driver.execute_script("return arguments[0].scrollTop;", container)
            if abs(cur2 - tgt) <= step:
                return True

        # keep scrolling down
        next_top = min(max_top, max(cur, last_top) + step)
        driver.execute_script("arguments[0].scrollTop = arguments[1];", container, next_top)
        time.sleep(0.25)
        new_top = driver.execute_script("return arguments[0].scrollTop;", container)

        if new_top < last_top:
            driver.execute_script("arguments[0].scrollTop = arguments[1];", container, last_top + step)
            time.sleep(0.2)
            new_top = driver.execute_script("return arguments[0].scrollTop;", container)
            step = min(step + 100, 1400)

        last_top = new_top
        if last_top >= max_top - 2:
            break
    return False

def get_day_panel_webelement(driver, target_date):
    labels = _day_heading_variants(target_date)
    container = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".calendarRow.slds-scrollable_y"))
    )
    panel = driver.execute_script("""
      const cont = arguments[0], labels = arguments[1];
      function text(n){ return (n && (n.innerText || n.textContent) || '').trim(); }
      const heads = Array.from(cont.querySelectorAll('h2.slds-assistive-text'));
      let H=null; for(const h of heads){ if(labels.includes(text(h))){ H=h; break; } }
      if(!H) return null;
      let sib = H.nextElementSibling;
      while(sib && !sib.classList.contains('calendarDay')) { sib = sib.nextElementSibling; }
      return sib || null;
    """, container, labels)
    return panel  # WebElement or None

def open_event_from_day_panel(driver, target_date, code, sem, sec, sess):
    panel = get_day_panel_webelement(driver, target_date)
    if not panel:
        return False

    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'start'});", panel)
        time.sleep(0.2)
    except Exception:
        pass

    links = panel.find_elements(By.CSS_SELECTOR, "a.subject-link, a[data-id='subject-link']")
    if not links:
        return False

    best, best_href = None, None
    for el in links:
        try:
            txt = (el.text or el.get_attribute("innerText") or "").strip()
            if not txt:
                continue
            if matches_event_text(txt, code, sem, sec, sess):
                href = el.get_attribute("href") or ""
                if href.startswith("/lightning/r/") or "lightning/r/" in href:
                    best, best_href = el, href
                    break
                if best is None:
                    best, best_href = el, href
        except StaleElementReferenceException:
            continue

    if not best:
        return False

    # Click the exact element we found in THIS day panel
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", best)
        time.sleep(0.1)
        driver.execute_script("arguments[0].click();", best)
    except Exception:
        try:
            best.click()
        except Exception:
            return False

    if not best_href or best_href.strip() in ("", "javascript:void(0)", "#"):
        try:
            more_details = WebDriverWait(driver, 8).until(
                EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='More Details']"))
            )
            js_click(driver, more_details)
        except Exception:
            try:
                preview_icon = WebDriverWait(driver, 4).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "a.previewIcon, a[aria-label='Preview icon']"))
                )
                js_click(driver, preview_icon)
                more_details = WebDriverWait(driver, 6).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='More Details']"))
                )
                js_click(driver, more_details)
            except Exception:
                pass

    return True

# =============================
# 1) Parse date argument (d/m/Y) or use today's date
# =============================
def parse_date_any(s: str) -> date:
    s = s.strip()
    fmts = ["%d/%m/%Y","%d-%m-%Y","%Y-%m-%d","%d-%b-%y","%d-%b-%Y",
            "%A, %d %B %Y","%A, %d %B %Y at %I:%M:%S %p"]
    for f in fmts:
        try: return datetime.strptime(s, f).date()
        except Exception: pass
    return pd.to_datetime(s, dayfirst=True).date()

if len(sys.argv) > 1:
    selected_date = parse_date_any(sys.argv[1])
    print(f"📅 Using date: {selected_date} (from argument)")
else:
    selected_date = datetime.today().date()
    print(f"📅 Using date: {selected_date} (today)")

# =============================
# 2) Load Excel file + Initial Setup fields (Session optional)
# =============================
try:
    attendance_df = pd.read_excel(file_path, sheet_name="Attendance", header=1)
    setup_df = pd.read_excel(file_path, sheet_name="Initial Setup", header=None)
except FileNotFoundError:
    print(f"❌ Excel file not found: {file_path}")
    sys.exit(1)
except Exception as e:
    print(f"❌ Failed to read Excel: {e}")
    sys.exit(1)

def val_or_empty(x):
    s = str(x).strip()
    return "" if s.lower() in ("nan", "none", "null") else s

# Extract values from Initial Setup (Column B values on rows 1..5)
course_name   = val_or_empty(setup_df.iloc[0, 1]) if len(setup_df) > 0 else ""
course_code   = val_or_empty(setup_df.iloc[1, 1]) if len(setup_df) > 1 else ""
semester      = val_or_empty(setup_df.iloc[2, 1]) if len(setup_df) > 2 else ""
class_section = val_or_empty(setup_df.iloc[3, 1]) if len(setup_df) > 3 else ""
session_no    = val_or_empty(setup_df.iloc[4, 1]) if len(setup_df) > 4 else ""  # optional

print("\n📘 Course Details from Initial Setup:")
print(f"   Course Name   : {course_name or '(blank)'}")
print(f"   Course Code   : {course_code or '(blank)'}")
print(f"   Semester      : {semester or '(blank)'}")
print(f"   Class Section : {class_section or '(blank)'}")
print(f"   Session       : {session_no or '(blank/optional)'}\n")

# Validate required fields (session optional)
missing = []
if not course_code:   missing.append("Course Code (B2)")
if not semester:      missing.append("Semester (B3)")
if not class_section: missing.append("Class Section (B4)")
if missing:
    print("⚠️ Initial Setup is incomplete. Required fields missing:")
    for m in missing:
        print(f"   - {m}")
    print("Please fill these in 'Initial Setup' sheet and re-run.")
    sys.exit(1)

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
    print("❌ No column found for the specified date in the 'Attendance' sheet.")
    sys.exit(1)
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
# 3) Selenium with webdriver-manager (auto ChromeDriver) + Profile fallback
# =============================
PROFILE_DIR = os.path.abspath("./slcm_automation_profile")  # dedicated reusable profile
os.makedirs(PROFILE_DIR, exist_ok=True)

# Optional: clear leftover Chrome lock files if the profile isn't actually open
for name in os.listdir(PROFILE_DIR):
    if name.startswith("Singleton"):
        try:
            os.remove(os.path.join(PROFILE_DIR, name))
        except Exception:
            pass

TEMP_PROFILE_DIR = None  # set when we fall back

def build_options(user_data_dir):
    opts = webdriver.ChromeOptions()
    opts.add_argument(f"--user-data-dir={user_data_dir}")
    opts.add_argument("--no-first-run")
    opts.add_argument("--no-default-browser-check")
    # opts.add_argument("--headless=new")  # keep visible for SSO/Lightning
    return opts

def start_driver_with_fallback():
    """Try dedicated profile; if locked, fall back to a fresh temp profile."""
    global TEMP_PROFILE_DIR
    try:
        options = build_options(PROFILE_DIR)
        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=options)
    except SessionNotCreatedException:
        print("⚠️ Dedicated profile is in use/locked. Falling back to a fresh temp profile...")
        TEMP_PROFILE_DIR = tempfile.mkdtemp(prefix="slcm_profile_")
        options = build_options(TEMP_PROFILE_DIR)
        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=options)

driver = start_driver_with_fallback()

# --- Bootstrap / SSO ---
if not hard_nav(driver, HOME_URL):
    hard_nav(driver, BASE_URL)
    hard_nav(driver, HOME_URL)

cur = driver.current_url.lower()
print("🌐 After bootstrap:", cur)

if ("login.microsoftonline.com" in cur) or ("saml" in cur) or ("manipal.edu" in cur and "/login" in cur):
    print("🔐 SSO/login detected. Complete it in the opened Chrome window.")
    try:
        input("Press Enter here AFTER you reach Salesforce Home... ")
    except EOFError:
        print("⏳ Waiting 60s for manual login (no console input available)...")
        time.sleep(60)
    hard_nav(driver, HOME_URL)

WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//a[@title='Calendar']")))
print("✅ Logged in & on Lightning Home")

# =============================
# 4) Calendar → date → down-only scroll to the day's panel → open event
# =============================
# Open Calendar tab
cal_tab = WebDriverWait(driver, 40).until(
    EC.element_to_be_clickable((By.XPATH, "//a[@title='Calendar']"))
)
js_click(driver, cal_tab)

# Click date (mini calendar)
WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "calendarSidebar")))
time.sleep(0.15)
day_number = str(selected_date.day).lstrip("0")
ok = driver.execute_script("""
const wrap = document.querySelector('#calendarSidebar');
if (!wrap) return false;
const dayNodes = wrap.querySelectorAll('table.datepicker .slds-day, .slds-day');
for (const n of dayNodes) {
  const txt = (n.textContent || '').trim();
  const disabled = n.getAttribute('aria-disabled') === 'true' || (n.className || '').includes('disabled');
  if (!disabled && txt === arguments[0]) {
    n.scrollIntoView({block:'center'}); n.click(); return true;
  }
}
return false;
""", day_number)
if not ok:
    raise RuntimeError(f"❌ Could not click mini calendar date {day_number}")
print(f"✅ Clicked calendar date (fast): {day_number}")

# Ensure day list exists
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".calendarRow.slds-scrollable_y")))

# Down-only scroll to the correct day's panel
disable_auto_scroll(driver)
try:
    ok_scroll = scroll_to_day_panel(driver, selected_date, timeout=40)
finally:
    enable_auto_scroll(driver)

if not ok_scroll:
    print("⚠️ Could not scroll down to the selected day's panel.")
    driver.quit()
    if TEMP_PROFILE_DIR:
        try: shutil.rmtree(TEMP_PROFILE_DIR, ignore_errors=True)
        except Exception: pass
    sys.exit(1)

# Open event strictly from that day panel
if not open_event_from_day_panel(driver, selected_date, course_code, semester, class_section, session_no):
    driver.quit()
    if TEMP_PROFILE_DIR:
        try: shutil.rmtree(TEMP_PROFILE_DIR, ignore_errors=True)
        except Exception: pass
    raise RuntimeError("❌ Could not open any candidate event tile for the selected date.")

# "More Details" if a popover appears; otherwise Lightning may navigate directly
try:
    more_details = WebDriverWait(driver, 6).until(
        EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='More Details']"))
    )
    js_click(driver, more_details)
except Exception:
    pass  # direct navigation case

# Open Attendance tab
def click_attendance_tab_fast(driver):
    js = """
    let el = document.querySelector("a[data-label='Attendance']");
    if (!el) {
        const span = Array.from(document.querySelectorAll('span.title'))
            .find(s => (s.textContent || '').trim() === 'Attendance');
        if (span) el = span.closest('a, button, [role="tab"]') || span;
    }
    if (el) { el.click(); return true; }
    return false;
    """
    ok = driver.execute_script(js)
    if not ok:
        att_tab = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[@data-label='Attendance'] | //span[@class='title' and normalize-space()='Attendance']"))
        )
        js_click(driver, att_tab)

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
# 7) Done + credit + temp profile cleanup
# =============================
print("🎉 Attendance marking complete!")
time.sleep(1.5)
driver.quit()

# Clean up temp profile if we used one
if 'TEMP_PROFILE_DIR' in globals() and TEMP_PROFILE_DIR:
    try: shutil.rmtree(TEMP_PROFILE_DIR, ignore_errors=True)
    except Exception: pass

print("\n====================================================")
print("👨‍💻 Developed by: Anirudhan Adukkathayar C, SCE, MIT")
print("====================================================\n")

