# 🖨️ ChangeBarTenderPrinter

A tool for batch-updating the default printer in BarTender `.btw` label files based on filename prefixes.  
Executed via `.vbs → PowerShell → Python`, fully configurable through `config.ini`.

---

## 🔧 What It Does

- Loads label directories from `config.ini`
- Detects prefix in each `.btw` filename
- Uses the prefix to determine the correct printer (based on a config mapping)
- Saves the updated `.btw` file

---

## 🚀 How It Works

### 1. `ChangeBarTenderPrinter.vbs`
Runs the whole process silently in the background:

```vbscript
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File ""...ChangeBarTenderPrinter.ps1""", 0
Set WshShell = Nothing
```

### 2. `ChangeBarTenderPrinter.ps1`
- Loads config paths from `config.ini`
- Launches the Python script with `Start-Process`
- Optionally shows a MessageBox when configuration is missing or incorrect

### 3. `ChangeBarTenderPrinter.py`
- Uses the BarTender COM interface
- Based on .btw filename prefixes, it selects the appropriate printer
- Applies changes and logs results to a `.log` file

---

## 📁 Project Structure

```
Win_ChangeBarTenderPrinter/
│
├── ChangeBarTenderPrinter.vbs         # VBScript launcher
├── ChangeBarTenderPrinter.ps1         # PowerShell script
├── ChangeBarTenderPrinter.py          # Python core
├── config.ini                         # Configuration file
├── [ico]/                             # Optional icons (e.g., window icon)
├── [log]/                             # Directory for logs
└── ...
```

---

## ⚙️ Example `config.ini`

```ini
[Paths]
python_path = C:/Users/.../python.exe
python_script_path = C:/Users/.../ChangeBarTenderPrinter.py
bartender_path = C:/Program Files/BarTender Suite/bartender.exe
labels_folders = C:/Labels/Folder1; C:/Labels/Folder2
log_file_path = C:/Logs/ChangePrinter.log

[PrinterMapping]
ABC = Zebra GX430t
DEF = Brother QL-800
```

---

## ▶️ Running the Script

- Customize your `config.ini`
- Double-click `ChangeBarTenderPrinter.vbs` (or run the `.ps1` file manually for debugging)

---

## 📒 Logging

Activity is logged to the file specified in `config.ini` — includes both successes and errors.

---

## 🛠️ Requirements

- 🪟 Windows

- 🐍 Python 3.x + `pywin32` (`pip install pywin32`)

- 📦 BarTender installed

- PowerShell

---

## 🪪 License

MIT

---

## ✨ Author

Miloslav Hradecky 📧 [miloslavhradecky76@gmail.com]
