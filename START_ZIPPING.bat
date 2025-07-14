@echo off
cls
echo Starting the process... Please wait for the PowerShell window to open.

:: This line launches PowerShell, tells it to bypass security for THIS RUN ONLY,
:: and executes the .ps1 script that is in the same folder as this .bat file.
powershell.exe -ExecutionPolicy Bypass -File "%~dp0create_zips.ps1"

echo.
echo The script has finished. This window will close automatically.
timeout /t 5 >nul