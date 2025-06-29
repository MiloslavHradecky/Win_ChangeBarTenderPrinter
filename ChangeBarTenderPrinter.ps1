# 📌 PowerShell version check (Kontrola verze PowerShellu)
if ($PSVersionTable.PSVersion.Major -lt 7) {
    [System.Windows.Forms.MessageBox]::Show("❌ Tento skript vyžaduje PowerShell 7 nebo vyšší!", "Neplatná verze PowerShellu", 0)
    exit 1
}

# 📌 Add support for MessageBox (Pridani podpory pro MessageBox)
Add-Type -AssemblyName System.Windows.Forms

# 📌 Get path to the current script directory (Ziskani cesty k aktualnimu adresari skriptu)
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -Path $scriptDir

# 📌 Load configuration from config.ini (Nacteni konfigurace ze souboru config.ini)
$configFile = "$scriptDir\config.ini"
$configData = @{}  # ✅ Initialize an empty hashtable (Inicializujeme prazdny hash table pro data)

foreach ($line in Get-Content $configFile) {
    # 📌 Skip empty lines and lines without '=' (Vynechame prazdne radky a radky bez '=')
    if ($line -match "^\s*([^=]+)\s*=\s*(.+)\s*$") {
        $configData[$matches[1].Trim()] = $matches[2].Trim()
    }
}

# 📌 Verify that required values are present (Overeni, ze hodnoty nejsou null)
if (-not $configData.ContainsKey("python_path") -or -not $configData.ContainsKey("python_script_path")) {
    [System.Windows.Forms.MessageBox]::Show("❌ Chyba: Soubor config.ini neobsahuje správně zadané hodnoty!", "Chyba konfigurace", 0)
    exit 1
}

# 📌 Load paths from config and convert slashes (Nacteni cest z konfigurace a prepocitani lomitek)
$pythonPath = $configData["python_path"] -replace '/', '\'
$scriptPath = $configData["python_script_path"] -replace '/', '\'

# 📌 Start the Python script (Spusteni Python skriptu)
Start-Process -FilePath $pythonPath -ArgumentList $scriptPath -NoNewWindow

# 📌 Debug: Show loaded paths in MessageBox (Vypis hodnot do MessageBoxu - Debug vystup)
# [System.Windows.Forms.MessageBox]::Show("✅ Cesta k Pythonu: $pythonPath`n✅ Cesta ke skriptu: $scriptPath", "Načtené cesty", 0)
