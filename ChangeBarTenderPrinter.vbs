Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File ""C:\Users\Home\Documents\Coding\Windows\Win_ChangeBarTenderPrinter\ChangeBarTenderPrinter.ps1""", 0
Set WshShell = Nothing