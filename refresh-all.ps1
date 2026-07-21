# refresh-all.ps1
# Run this whenever you update any source Excel/CSV files.
# Refreshes: campaign calendar, pipeline, profiling, outreach data, sequence states tags
#
# NOTE: Sequence States is an occasional export (filename has a date).
#       Update the path in import-seqstates.ps1 when you get a new export.
#       Sequence Stats is refreshed regularly via import-outreach.ps1.

$outDir = "C:\Users\I572929\campaign-calendar-site"

Write-Host "=== EMEA Campaign Insights - Full Data Refresh ===" -ForegroundColor Cyan
Write-Host ""

# 1. Campaign Calendar
Write-Host "[1/5] Refreshing Campaign Calendar..." -ForegroundColor Yellow
powershell.exe -ExecutionPolicy Bypass -File "$outDir\import-pbi.ps1"

# 2. Profiling Requests
Write-Host "[2/5] Refreshing Profiling Requests..." -ForegroundColor Yellow
powershell.exe -ExecutionPolicy Bypass -File "$outDir\import-profiling.ps1"

# 3. Pipeline
Write-Host "[3/5] Refreshing Pipeline..." -ForegroundColor Yellow
powershell.exe -ExecutionPolicy Bypass -File "$outDir\import-pipeline.ps1"

# 4. Outreach Data (Sequence Stats - refreshed regularly)
Write-Host "[4/5] Refreshing Outreach Sequence Data..." -ForegroundColor Yellow
powershell.exe -ExecutionPolicy Bypass -File "$outDir\import-outreach.ps1"

# 5. Sequence States Tags (occasional export - update path in import-seqstates.ps1 when new file arrives)
Write-Host "[5/5] Refreshing Sequence States Tags..." -ForegroundColor Yellow
powershell.exe -ExecutionPolicy Bypass -File "$outDir\import-seqstates.ps1"

Write-Host ""
Write-Host "=== All done! Refresh the browser to see updated data. ===" -ForegroundColor Green
