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
     - Search for **Python 3.11** (or later).
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
   - Navigate to the project folder where `requirements.txt` is located.
   - Run:
     ```bash
     pip install -r requirements.txt
     ```

3. **Google Chrome**
   - Install the latest **Google Chrome** browser from [Google Chrome Download](https://www.google.com/chrome/).

4. **ChromeDriver**
   - The script uses `webdriver_manager` to download the correct version of ChromeDriver automatically.
   - Ensure Chrome is installed and updated.

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
   - Sheet **"Initial Setup"**, cell **B2**, contains the **course code**.  
     ⚠️ This course code must exactly match the subject code shown in SLCM Calendar events.

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

---

## 👨‍💻 Developer
**Developed by:** Anirudhan Adukkathayar C, SCE, MIT
