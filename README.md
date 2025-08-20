# 📌 Automated Attendance Marker (MAHE SLCM)

This project automates the attendance marking process on **MAHE SLCM (Salesforce Lightning)** using **Python + Selenium**.  
It extracts **absentees from an Excel sheet** (`attendance.xlsx`) and unticks them on the attendance portal, then submits and confirms the attendance automatically.  

---

## ✨ Features
- ✅ Reads absentees list directly from **Excel** (`attendance.xlsx`).  
- ✅ Fast navigation (optimized calendar & Attendance tab click).  
- ✅ Supports **manual login once per day** (profile reuse keeps session).  
- ✅ Cross-platform: **Windows / macOS / Linux**.  
- ✅ Logs output to terminal:  
  - Students **unticked (absentees handled)**.  
  - Students **not found**.  
  - Final **summary counts**.  
- ✅ Developer credit footer at end of run.  

---

## 🛠 Requirements

- **Python 3.10+** must be installed on your system.  
  👉 [Download Python](https://www.python.org/downloads/)  

- **Google Chrome** (latest version).  

- Python dependencies are listed in `requirements.txt`:  
  ```
  selenium
  pandas
  openpyxl
  webdriver-manager
  ```

> 🔹 `webdriver-manager` automatically manages the correct ChromeDriver for your installed Chrome version. No manual downloads needed.  

---

## 📂 Files in this Project

- `maa.py` → Main automation script  
- `attendance.xlsx` → Excel file with attendance data  
- `requirements.txt` → Python dependencies  
- `README.md` → Documentation (this file)  

---

## ⚙️ Installation & Setup

### 1. Install Python (if not installed)
- [Download Python](https://www.python.org/downloads/)  
- During installation on **Windows**, check **“Add Python to PATH”**.  

### 2. Install project dependencies
Open a terminal/command prompt in the project folder and run:  
```bash
pip install -r requirements.txt
```

If you face issues with `pip`, try:
```bash
python -m pip install -r requirements.txt
```

### 3. Verify dependencies
Run this in terminal:  
```bash
python -m pip show selenium pandas openpyxl webdriver-manager
```
If all 4 packages are listed, setup is complete. ✅  

---

## ▶️ Usage

### Run for **today’s date**
```bash
python maa.py
```

### Run for a **specific date** (format: `DD/MM/YYYY`)
```bash
python maa.py 20/08/2025
```

---

## 🚀 How It Works

1. Opens Chrome using a dedicated profile (`slcm_automation_profile`).  
   - First time: You must log in (SSO/OTP manually).  
   - Next runs: The login session is reused.  

2. Navigates to:  
   - **Calendar → Selected Date → Subject → Attendance Tab**  

3. Unticks all **absentees** found in Excel.  

4. Clicks **Submit Attendance** → **Confirm Submission**.  

5. Prints **summary of results** in terminal.  

---

## 📊 Example Output
```
📅 Using date: 2025-08-20 (today)
✅ Using date column in sheet: 20/08/2025
Absentees (IDs to untick): ['230905023', '230905098', '230905108']
🌐 After bootstrap: https://maheslcmtech.lightning.force.com/...
✅ Logged in & on Lightning Home
✅ Clicked calendar date (fast): 20
✅ Opened Attendance tab (fast)
🔎 Searching for each absentee ID on page...
✔️ Unticked absentee: 230905023
✔️ Unticked absentee: 230905098
❓ Not found (1): ['230905108']
✔️ Unticked (absentees): 2
✅ Clicked Submit Attendance
✅ Confirmation modal visible
✅ Clicked Confirm via locator: .//button[normalize-space()='Confirm Submission']
🎉 Attendance marking complete!

=================================================
👨‍💻 Developed by: Anirudhan Adukkathayar C, SCE, MIT
=================================================
```

---

## 🖥️ Notes

- First run may be slower due to login/OTP. Subsequent runs are faster.  
- If you face **profile in use errors**, close all Chrome windows before running.  
- If UI changes in SLCM portal, XPath selectors may need updates.  
- On **Windows**, run scripts using `python` in **Command Prompt** or **PowerShell**.  
- On **macOS/Linux**, run in **Terminal**.  

---

## 👨‍💻 Developer

**Anirudhan Adukkathayar C**  
*SCE, MIT*  
