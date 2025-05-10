Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File ""C:\DfpScripts\ChangeBarTenderPrinter\script\ChangePrinterDfP.ps1""", 0
Set WshShell = Nothing