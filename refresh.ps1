# ============================================================
# refresh.ps1 — Weekly data refresh for Campaign Insights site
# Run this script each week after saving new source files
# ============================================================

$ErrorActionPreference = "Stop"
$site = "C:\Users\I572929\campaign-calendar-site"
$calFile  = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Calendars\Campaign Calendars combined.xlsx"
$pipeFile = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Pipeline\Week 22 DL.xlsx"
$seqFile  = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Outreach\Sequence_Stats_2026-06-02.csv"
$expFile  = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Outreach\export-reports-table-1780391465252.csv"

function ExcelDateToISO($serial) {
    if (-not $serial -or $serial -eq "") { return "" }
    try { return ([DateTime]::FromOADate([double]$serial)).ToString("yyyy-MM-dd") } catch { return "" }
}

Write-Host "=== Campaign Insights Weekly Refresh ===" -ForegroundColor Cyan
$startTime = Get-Date

# ── 1. Campaign Calendar data ─────────────────────────────────────────────
Write-Host "`n[1/4] Reading campaign calendar..." -ForegroundColor Yellow
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($calFile, 0, $true)
$ws = $wb.Sheets.Item(1)
$lastRow = $ws.UsedRange.Rows.Count

# Map headers — scan all columns, skip blanks but don't stop at first blank
$headers = @{}
for ($col = 1; $col -le $ws.UsedRange.Columns.Count; $col++) {
    $val = $ws.Cells.Item(1, $col).Value2
    if ($val) {
        # Normalise known caps variants
        if ($val -eq "DG PROGRAM") { $val = "DG Program" }
        $headers[$val] = $col
    }
}

$campRecords = @()
for ($r = 2; $r -le $lastRow; $r++) {
    $include = $ws.Cells.Item($r, $headers["Include in EMEA 'Campaign Insights'?"]).Value2
    if ($include -ne "Yes") { continue }
    $obj = [ordered]@{}
    foreach ($h in $headers.Keys) {
        $val = $ws.Cells.Item($r, $headers[$h]).Value2
        # Convert Excel date serials for date columns
        if ($h -match "Date" -and $val -and $val -is [double]) { $val = ExcelDateToISO $val }
        $obj[$h] = if ($null -eq $val) { $null } else { "$val" }
    }
    $campRecords += [PSCustomObject]$obj
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "  $($campRecords.Count) campaigns loaded"

$campJson = $campRecords | ConvertTo-Json -Compress -Depth 3
"window.CAMP_DATA=$campJson;" | Set-Content "$site\data-camp.js" -Encoding UTF8

# ── 2. Pipeline data (via CSV for speed) ─────────────────────────────────
Write-Host "[2/4] Reading pipeline data..." -ForegroundColor Yellow
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($pipeFile, 0, $true)
$csvTemp = "$site\_pipe_temp.csv"
$wb.SaveAs($csvTemp, 6)
$wb.Close($false); $excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$keep = @("Opp Campaign ID","Opportunity ID","Eur","DRM Category","SDE Handover Date","Create Date","Create Quarter","Closing Qtr","TCP Pipeline Source Desc","MM Identifier","RBC","SAP Mastercode","IAC (Engagement Model)","Solution Area (L1)","Sub-Solution Area (L2)","Account Name","Opp Description")
function ToISO($d){if(-not $d -or $d.Trim() -eq ""){return ""}try{return[DateTime]::Parse($d.Trim()).ToString("yyyy-MM-dd")}catch{return ""}}

$pipeCsv = Import-Csv $csvTemp
$pipeRecords = $pipeCsv | ForEach-Object {
    $obj = [ordered]@{}
    foreach ($h in $keep) {
        $val = if ($_.PSObject.Properties[$h]) { $_.$h } else { "" }
        if ($h -match "Date") { $val = ToISO $val }
        elseif ($h -eq "Eur") { try { $val = [double]($val -replace '[^0-9.\-]','') } catch { $val = 0 } }
        $obj[$h] = $val
    }
    [PSCustomObject]$obj
}
Remove-Item $csvTemp -Force
Write-Host "  $($pipeRecords.Count) pipeline rows loaded"

$pipeJson = $pipeRecords | ConvertTo-Json -Compress -Depth 3
"window.PIPE_DATA=$pipeJson;" | Set-Content "$site\data-pipe.js" -Encoding UTF8

# ── 3. Outreach data ──────────────────────────────────────────────────────
Write-Host "[3/4] Reading outreach data..." -ForegroundColor Yellow
$seqStats = Import-Csv $seqFile
$expReports = Import-Csv $expFile

# Build Last used at lookup from export-reports
$lastUsedMap = @{}
$expReports | ForEach-Object {
    $sid = $_."Sequence ID".Trim()
    $luat = $_."Last used at".Trim()
    if ($sid -and $luat) { $lastUsedMap[$sid] = $luat.Substring(0,10) }
}

# We need the campaign mapping - load outreach.json as base
$outBase = Get-Content "$site\outreach.json" -Raw | ConvertFrom-Json

# Merge sequence stats into outreach base
$seqStatsMap = @{}
$seqStats | ForEach-Object { $seqStatsMap[$_."Sequence ID".Trim()] = $_ }

$outRecords = $outBase | ForEach-Object {
    $sid = "$($_.'Sequence ID')".Trim()
    $stats = $seqStatsMap[$sid]
    $obj = [ordered]@{}
    $_.PSObject.Properties | ForEach-Object { $obj[$_.Name] = $_.Value }
    # Update stats from latest export
    if ($stats) {
        $obj["Meetings Booked"]      = $stats."Meetings Booked"
        $obj["Prospects - Total"]     = $stats."Prospects - Total"
        $obj["Prospects - Delivered"] = $stats."Prospects - Delivered"
        $obj["Prospects - Finished"]  = $stats."Prospects - Finished"
        $obj["Prospects - Opened"]    = $stats."Prospects - Opened"
        $obj["Prospects - Replied"]   = $stats."Prospects - Replied"
        $obj["Emails - Deliveries"]   = $stats."Emails - Deliveries"
        $obj["Emails - Opens"]        = $stats."Emails - Opens"
        $obj["Emails - Replies"]      = $stats."Emails - Replies"
    }
    $obj["Last used at"] = if ($lastUsedMap[$sid]) { $lastUsedMap[$sid] } else { "" }
    [PSCustomObject]$obj
}
Write-Host "  $($outRecords.Count) outreach rows, $($lastUsedMap.Count) with Last used at"

$outJson = $outRecords | ConvertTo-Json -Compress -Depth 3
"window.OUT_DATA=$outJson;" | Set-Content "$site\data-out.js" -Encoding UTF8

# ── 4. Update embedded data in index.html ────────────────────────────────
Write-Host "[4/4] Updating index.html..." -ForegroundColor Yellow
$html = Get-Content "$site\index.html" -Raw

# Update d-camp
$html = [regex]::Replace($html,
    '(<script type="application/json" id="d-camp">)\[.*?\](</script>)',
    "`$1$campJson`$2",
    [System.Text.RegularExpressions.RegexOptions]::Singleline)

# Update d-pipe (use external file, just clear embedded to save size)
$html = [regex]::Replace($html,
    '(<script type="application/json" id="d-pipe">)\[.*?\](</script>)',
    '$1[]$2',
    [System.Text.RegularExpressions.RegexOptions]::Singleline)

# Update d-out
$html = [regex]::Replace($html,
    '(<script type="application/json" id="d-out">)\[.*?\](</script>)',
    "`$1$outJson`$2",
    [System.Text.RegularExpressions.RegexOptions]::Singleline)

Set-Content "$site\index.html" -Value $html -Encoding UTF8

# ── 5. Backup ─────────────────────────────────────────────────────────────
$backupName = "index_$(Get-Date -Format 'yyyy-MM-dd_HHmm').html"
Copy-Item "$site\index.html" "$site\backups\$backupName"

$elapsed = [Math]::Round(((Get-Date) - $startTime).TotalSeconds)
Write-Host "`n=== Done in ${elapsed}s — backup saved as $backupName ===" -ForegroundColor Green
Write-Host "Refresh http://localhost:8080 to see the updated site." -ForegroundColor Cyan
