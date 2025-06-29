# 🖨️ ChangeBarTenderPrinter

Nástroj pro hromadnou změnu výchozí tiskárny v BarTender `.btw` souborech na základě jejich názvu.  
Spouštěno přes `.vbs → PowerShell → Python`, plně konfigurovatelné přes `config.ini`.

---

## 🔧 Co to dělá

- Načte složky s etiketami z `config.ini`
- Pro každý `.btw` soubor detekuje prefix v názvu
- Na základě prefixu vybere správnou tiskárnu (dle mapování v configu)
- Uloží aktualizovaný `.btw` soubor

---

## 🚀 Jak to funguje

### 1. `ChangeBarTenderPrinter.vbs`
Spustí celý proces nenápadně na pozadí:

```vbscript
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File ""...ChangeBarTenderPrinter.ps1""", 0
Set WshShell = Nothing
```

### 2. `ChangeBarTenderPrinter.vbs`
- Spouštěcí skript pro uživatele (dvojklikem)
- Spustí PowerShell s vypnutým oknem

### 3. `ChangeBarTenderPrinter.ps1`
- Načte `config.ini` s cestami
- Spustí Python skript (`Start-Process`)
- Volitelně zobrazí MessageBox s chybou při chybějící konfiguraci

### 4. `ChangeBarTenderPrinter.py`
- Používá COM rozhraní BarTender
- Na základě názvu souboru `.btw` vybere tiskárnu (prefix → printer)
- Uloží změny a zaznamená průběh do `.log` souboru

---

## 📁 Struktura projektu

```
Win_ChangeBarTenderPrinter/
│
├── ChangeBarTenderPrinter.vbs         # VBScript spouštěč
├── ChangeBarTenderPrinter.ps1         # PowerShell skript
├── ChangeBarTenderPrinter.py          # Python jádro
├── config.ini                         # Konfigurace (viz níž)
├── [ico]/                             # Ikony (např. .ico pro okno)
├── [log]/                             # Složka s logy
└── ...
```

---

## ⚙️ config.ini ukázka

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

## ▶️ Spuštění

- Uprav config.ini
- Dvojklik na ChangeBarTenderPrinter.vbs
- (nebo spusť ChangeBarTenderPrinter.ps1 ručně při ladění)

---

## 📒 Záznamy

Logy se zapisují do .log souboru určeného v config.ini, včetně úspěšných změn i případných chyb.

---

## 🛠️ Závislosti

- 🪟 Windows

- 🐍 Python 3.x + pywin32 (pip install pywin32)

- 📦 Nainstalovaný BarTender

- PowerShell

---

## 🪪 Licence

MIT

---

## ✨ Autor

Miloslav Hradecky 📧 Pro dotazy nebo vylepšení: [miloslavhradecky76@gmail.com]
