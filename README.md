# Automated Attendance Marker (MAHE SLCM) — Enhanced Version

This tool automates attendance marking in the **MAHE SLCM portal** using Selenium and your Excel register.  

It now supports **automatic Excel file selection** (via UI picker if not found) and strict validation of course metadata from the **Initial Setup** sheet to ensure correct subject matching.

---

## 🚀 Features
- Reads absentees from your Excel file (`attendance.xlsx` or selected file).
- Saves Excel location in a config file for future runs.
- Automatically logs in via your saved Chrome profile (no need to log in every time).
- Opens Calendar → selects date → picks the right subject tile (Course Code + Semester + Section + optional Session).
- Unticks absent students in the SLCM portal.
- Submits and confirms attendance.
- Prints summary of unticked and not-found students in the terminal.

---

## ⚙️ Requirements

### 1. Python
- Install **Python 3.11+**.
- **Windows**:
  - Open **Microsoft Store**, search for **Python**, install it.
- Verify:
  ```bash
  python --version
  ```

### 2. Dependencies
Install in one line:
```bash
pip install selenium pandas openpyxl webdriver-manager
```

### 3. Google Chrome
Install the latest [Google Chrome](https://www.google.com/chrome/).  

---

## 📂 Excel Setup

Your Excel workbook **must contain** two sheets:

### 1. Attendance Sheet
- Name: **Attendance**  
- Columns:  
  - **Reg. No.** (first column)  
  - Dates as columns (dd/mm/yyyy or Excel date format)  
  - Mark **`ab`** for absentees.  

### 2. Initial Setup Sheet
- Name: **Initial Setup**  
- Required fields in **Column B**:

| Field         | Example Value              | Notes                                                                 |
|---------------|----------------------------|-----------------------------------------------------------------------|
| Course Name   | Operating Systems Lab      | Free text, for readability only                                       |
| Course Code   | CSE 3142                   | **Must exactly match** SLCM Calendar subject code                     |
| Semester      | V                          | Roman numeral or value exactly as shown in SLCM                       |
| Class Section | B                          | Section identifier (A, B, C …)                                        |
| Session       | 1                          | Optional. Use for **labs/batches** (1, 2, …). Leave blank for theory. |

⚠️ **Important**  
- The **Course Code, Semester, Class Section, Session** must **exactly match** the subject event in SLCM Calendar.  
- Example event in SLCM:  
  ```
  CSE 3142 - CSE 3142 - OPERATING SYSTEMS LAB - 905 - Semester V: Program Sec B-1
  ```
  → Course Code = `CSE 3142`  
  → Semester = `V`  
  → Class Section = `B`  
  → Session = `1`  

---

## ▶️ Running the Script

1. Place your Excel file in the project folder, or select it via the **UI prompt** on first run.  
   - The chosen path is saved in `attendance_config.json`.  
2. Open terminal/command prompt in the project folder.  
3. Run:
   ```bash
   python maa.py
   ```
   Or, specify a date:
   ```bash
   python maa.py 30/07/2025
   ```
4. Complete **SSO/OTP login** in the Chrome window if prompted.  
5. Script proceeds to mark absentees automatically.  

---

## 📊 Output
At the end of execution, you will see:

- ✅ Number of absentees successfully unticked.  
- ❌ Students not found in the portal.  
- 🎉 Confirmation of attendance submission.  

---

## ⚠️ **DISCLAIMER**

This script is **NOT** part of official MAHE SLCM.  

- It may break if the SLCM website UI changes.  
- Wrong configuration in Excel may cause incorrect submission.  
- ✅ Always **manually verify attendance** after running.  
- Developer assumes **no responsibility** for misuse or errors.  
- Use strictly at your **own risk**.  

---

## 👨‍💻 Developer
**Developed by:** Anirudhan Adukkathayar C, SCE, MIT  

