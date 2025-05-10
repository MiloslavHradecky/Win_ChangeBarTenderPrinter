# Ziskani cesty k aktualnimu adresari skriptu
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -Path $scriptDir

# Zadej cestu k tvemu Pythonu a skriptu
$pythonPath = "C:\Users\Tester\AppData\Local\Programs\Python\Python313\python3.exe"
$scriptPath = "C:\DfpScripts\ChangeBarTenderPrinter\script\ChangeBarTenderPrinterDfP.py"

# Spusteni Python skriptu
Start-Process -FilePath $pythonPath -ArgumentList $scriptPath -NoNewWindow
