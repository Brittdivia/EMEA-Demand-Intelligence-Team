# refresh-all.ps1
# Run this whenever you update any source Excel/CSV files.
# Refreshes: campaign calendar, pipeline, profiling, outreach data

$outDir = "C:\Users\I572929\campaign-calendar-site"

Write-Host "=== EMEA Campaign Insights - Full Data Refresh ===" -ForegroundColor Cyan
Write-Host ""

# 1. Campaign Calendar
Write-Host "[1/4] Refreshing Campaign Calendar..." -ForegroundColor Yellow
powershell.exe -ExecutionPolicy Bypass -File "$outDir\import-pbi.ps1"

# 2. Profiling Requests
Write-Host "[2/4] Refreshing Profiling Requests..." -ForegroundColor Yellow
powershell.exe -ExecutionPolicy Bypass -File "$outDir\import-profiling.ps1"

# 3. Pipeline
Write-Host "[3/4] Refreshing Pipeline..." -ForegroundColor Yellow
powershell.exe -ExecutionPolicy Bypass -File "$outDir\import-pipeline.ps1"

# 4. Outreach Data (Sequence Stats)
Write-Host "[4/4] Refreshing Outreach Sequence Data..." -ForegroundColor Yellow
powershell.exe -ExecutionPolicy Bypass -File "$outDir\import-outreach.ps1"

Write-Host ""
Write-Host "=== All done! Refresh the browser to see updated data. ===" -ForegroundColor Green
