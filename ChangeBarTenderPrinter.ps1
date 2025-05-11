# 📌 Ziskani cesty k aktualnimu adresari skriptu
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -Path $scriptDir

# 📌 Nacteni konfigurace ze souboru config.ini
$configFile = "$scriptDir\config.ini"
$configData = @{}  # ✅ Inicializujeme prázdný hash table pro data

foreach ($line in Get-Content $configFile) {
    # 📌 Vynechame prazdne radky a radky bez '='
    if ($line -match "^\s*([^=]+)\s*=\s*(.+)\s*$") {
        $configData[$matches[1].Trim()] = $matches[2].Trim()
    }
}

# 📌 Overeni, ze hodnoty nejsou null
if (-not $configData.ContainsKey("python_path") -or -not $configData.ContainsKey("python_script_path")) {
    Write-Host "❌ Chyba: Soubor config.ini neobsahuje správně zadané hodnoty!"
    exit 1
}

# 📌 Nacteni cest z konfigurace a prepocitani lomitek
$pythonPath = $configData["python_path"] -replace '/', '\'
$scriptPath = $configData["python_script_path"] -replace '/', '\'

# 📌 Spusteni Python skriptu
# Start-Process -FilePath $pythonPath -ArgumentList $scriptPath -NoNewWindow

Write-Host($pythonPath)
Write-Host($scriptPath)
