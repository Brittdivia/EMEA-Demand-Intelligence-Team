# Rebuild data-camp.js from Campaign Calendars combined.xlsx
$calFile = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Calendars\Campaign Calendars combined.xlsx"
$site = "C:\Users\I572929\campaign-calendar-site"

function ExcelDateToISO($serial) {
    if (-not $serial -or $serial -eq "") { return "" }
    try { return ([DateTime]::FromOADate([double]$serial)).ToString("yyyy-MM-dd") } catch { return "" }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($calFile, 0, $true)
$ws = $wb.Sheets.Item(1)
$lastRow = $ws.UsedRange.Rows.Count
$lastCol = $ws.UsedRange.Columns.Count
Write-Host "Columns: $lastCol  Rows: $lastRow"

$headers = @{}
for ($col = 1; $col -le $lastCol; $col++) {
    $val = $ws.Cells.Item(1, $col).Value2
    if ($val) {
        if ($val -eq "DG PROGRAM") { $val = "DG Program" }
        $headers[$val] = $col
    }
}
Write-Host "Headers found: $($headers.Count)"
Write-Host "DG Program col: $($headers['DG Program'])"
Write-Host "Digital Assets col: $($headers['Digital Assets team support'])"

$campRecords = @()
$includeCol = $headers["Include in EMEA 'Campaign Insights'?"]
for ($r = 2; $r -le $lastRow; $r++) {
    $include = $ws.Cells.Item($r, $includeCol).Value2
    if ($include -ne "Yes") { continue }
    $obj = [ordered]@{}
    foreach ($h in $headers.Keys) {
        $val = $ws.Cells.Item($r, $headers[$h]).Value2
        if ($h -match "Date" -and $val -and $val -is [double]) { $val = ExcelDateToISO $val }
        # For tag fields, use Text to preserve full multi-line content
        if ($h -match "Tag|tag") {
            $textVal = $ws.Cells.Item($r, $headers[$h]).Text
            if ($textVal -and "$textVal" -ne "$val") { $val = $textVal }
        }
        $obj[$h] = if ($null -eq $val) { $null } else { "$val" }
    }
    $campRecords += [PSCustomObject]$obj
}
$wb.Close($false); $excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Campaigns loaded: $($campRecords.Count)"
$campJson = $campRecords | ConvertTo-Json -Compress -Depth 3
"window.CAMP_DATA=$campJson;" | Set-Content "$site\data-camp.js" -Encoding UTF8
$kb = [Math]::Round((Get-Item "$site\data-camp.js").Length / 1KB)
Write-Host "data-camp.js updated - ${kb}KB"

# Spot-check DG Program
$check = ($campRecords | Where-Object { $_."DG Program" -and $_."DG Program" -ne "" } | Measure-Object).Count
Write-Host "Records with DG Program: $check"
$check2 = ($campRecords | Where-Object { $_."Digital Assets team support" -and $_."Digital Assets team support" -ne "" } | Measure-Object).Count
Write-Host "Records with Digital Assets team support: $check2"
