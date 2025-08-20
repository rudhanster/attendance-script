# 📌 Automated Attendance Marker (MAHE SLCM)

[![Python](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/)
[![Selenium](https://img.shields.io/badge/selenium-latest-green.svg)](https://www.selenium.dev/)
[![License](https://img.shields.io/badge/license-MIT-lightgrey.svg)](LICENSE)

This project automates the process of **marking attendance** on the **MAHE SLCM portal** using Selenium and an Excel file.  
It reads the list of absentees from `OS V attendance.xlsx` and automates navigation + marking directly in the portal.  

---

## 🚀 Features
- 🔑 **Chrome profile reuse** → no login required each time.  
- 📊 **Excel-based absentees** → auto-read from `Attendance` sheet.  
- ⚡ **Fast navigation** → skips delays in Calendar/Attendance tab.  
- ✅ Automatically **unticks absentees** and submits.  
- 🖥 Works on **Windows, macOS, Linux**.  
- 📜 Prints a **summary** in the terminal:
  - ✔️ Number of students unticked  
  - ❓ Students not found  

---

## 📂 Excel Format

File: **`OS V attendance.xlsx`**

### Sheets
- **Attendance**
  - Column A → `Reg. No.`  
  - Column B → `Name`  
  - Column G onward → Dates (d/m/Y format)  
  - Value `ab` → absentee  
- **Initial Setup**
  - Cell `B2` → Subject Code  

Example:

| Reg. No. | Name       | 19/08/2025 | 20/08/2025 |
|----------|------------|------------|------------|
| 230905001 | Student A |            | ab         |
| 230905002 | Student B | ab         |            |

---

## 🛠 Installation

Clone repo:
```bash
git clone https://github.com/your-username/mahe-attendance-automation.git
cd mahe-attendance-automation
```

Install dependencies:
```bash
pip install -r requirements.txt
```

`requirements.txt`
```
selenium
pandas
openpyxl
chromedriver-autoinstaller
```

---

## ▶️ Usage

### Windows
```bat
python maa.py
```

### macOS / Linux
```bash
python3 maa.py
```

With a specific date:
```bash
python maa.py 20/08/2025
```

---

## ⚙️ How It Works
1. Opens Chrome with your profile.  
2. Loads SLCM portal → waits for login (SSO/OTP).  
3. Selects **date** in Calendar.  
4. Opens **subject** → Attendance tab.  
5. Unticks absentees listed in Excel.  
6. Submits & confirms.  
7. Prints a **summary** in the terminal.  

---

## 📸 Example Terminal Output
```
📅 Using date: 20/08/2025 (today)
✅ Using date column in sheet: 20/08/2025
Absentees (IDs to untick): ['230905001', '230905002']

🔎 Searching for each absentee ID on page...
✔️ Unticked absentee: 230905001
❓ Not found (1): ['230905002']

✔️ Unticked (absentees): 1
❓ Not unticked: 1

✅ Clicked Submit Attendance
✅ Confirmed submission

🎉 Attendance marking complete!

=================================================
👨‍💻 Developed by: Anirudhan Adukkathayar C, SCE, MIT
=================================================
```

---

## 👨‍💻 Developer
**Anirudhan Adukkathayar C**  
📍 SCE, MIT  

---

## 📜 License
MIT License – feel free to use and modify.
