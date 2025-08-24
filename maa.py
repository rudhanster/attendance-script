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

