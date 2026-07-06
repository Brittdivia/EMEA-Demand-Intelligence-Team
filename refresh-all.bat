@echo off
echo === EMEA Campaign Insights - Full Data Refresh ===
powershell.exe -ExecutionPolicy Bypass -File "C:\Users\I572929\campaign-calendar-site\refresh-all.ps1"
echo.
echo === Done! Refresh your browser (Ctrl+Shift+R) ===
pause
