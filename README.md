# Automated Attendance Marker (MAHE SLCM)

This project automates attendance marking in the **MAHE SLCM portal** using Selenium and an Excel file.

---

## 🚀 Features
- Reads absentees from an Excel file (`attendance.xlsx`).
- Automatically logs in via your saved Chrome profile (no need to log in every time).
- Opens Calendar, selects the correct date and subject.
- Unticks absent students in the SLCM portal.
- Submits and confirms attendance.
- Prints summary of unticked and not-found students in the terminal.

---

## ⚙️ Requirements

1. **Python Installation**
   - Install **Python 3.11+**.
   - On **Windows**, the easiest way:
     - Open Microsoft Store.
     - Search for **Python**.
     - Click **Get** and install.
   - Verify installation:
     ```bash
     python --version
     ```
     or
     ```bash
     py --version
     ```

2. **Install Dependencies**
   - Open **Command Prompt** (Windows) or **Terminal** (macOS/Linux).
  
   - Run:
     ```bash
     pip install selenium pandas openpyxl webdriver-manager

     ```

3. **Google Chrome**
   - Install the latest **Google Chrome** browser from [Google Chrome Download](https://www.google.com/chrome/).




---

## 📂 Project Setup

1. Clone or download this repository.
2. Place your **Excel file** in the project directory.
   - The file **must be named exactly**:
     ```
     attendance.xlsx
     ```
3. Ensure the Excel format:
   - Sheet **"Attendance"** contains registration numbers and attendance (`ab` for absentees).
   - Sheet **"Initial Setup"**, 


This sheet provides metadata that the script uses to correctly match the course in the SLCM Calendar.  

### Required Fields

| Field         | Example Value              | Notes                                                       |
|---------------|----------------------------|-------------------------------------------------------------|
| Course Name   | Operating Systems Lab      | Free text, used for human readability only                  |
| Course Code   | CSE 3142                   | **Must exactly match** the course code in SLCM              |
| Semester      | V                          | Roman numeral or value as displayed in SLCM                 |
| Class Section | B                          | Section identifier (A, B, C …)                              |
| Session       |                            | Session number as shown in the SLCM  (1, 2, …) for lab[##Keep it blank for theory|

---

### ⚠️ Important

These values must **exactly match** the subject event text in the SLCM Calendar.  
Otherwise, the script will not be able to locate the correct event for attendance.

### ⚠️ Important
SLCM calender should be set to default view , that is #week view.

### ⚠️ Important
Session number in excel sheet should be blank for theory and 1 or 2 for lab[batch number|

---

## ▶️ Running the Script

1. Open terminal/command prompt in the project folder.
2. Run the script:
   ```bash
   python maa.py
   ```
   Or, specify a date:
   ```bash
   python maa.py 20/08/2025
   ```

3. The script will:
   - Launch Chrome with your automation profile.
   - Navigate to the MAHE SLCM portal.
   - If SSO/OTP login is required, complete it in Chrome.
   - Automatically proceed with attendance marking.

---

## 📊 Output
At the end of execution, the script prints:
- ✅ Number of absentees successfully unticked.
- ❓ Students not found in the portal.
- 🎉 Confirmation that attendance was submitted.



## ⚠️ **DISCLAIMER**

This script is **NOT** an official part of MAHE SLCM.  
It automates browser actions to help with attendance marking, but:  

- It may fail if the website UI changes.  
- Incorrect automation may result in **wrong attendance submission**.  
- ✅ Always **manually verify attendance** after running the script.  
- The developer holds **no responsibility** for misuse or errors.  
- Use at your **own risk**. 

---

## 👨‍💻 Developer
**Developed by:** Anirudhan Adukkathayar C, SCE, MIT
