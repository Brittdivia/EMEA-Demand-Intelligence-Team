# Rebuild CAMP_PROF_META and CAMP_PROF_TAGS from Profiling request.xlsx
# Run this whenever new rows are added to the profiling request spreadsheet

$xlFile = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Profiling request.xlsx"
$outFile = "C:\Users\I572929\campaign-calendar-site\data-out.js"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($xlFile, 0, $true)
$ws = $wb.Sheets.Item(1)

# Map headers
$headers = @{}
for ($col = 1; $col -le 50; $col++) {
    $val = $ws.Cells.Item(1, $col).Value2
    if ($val) { $headers[$val] = $col }
}

$profMeta = [ordered]@{}
$profTags = [ordered]@{}

$lastRow = $ws.UsedRange.Rows.Count
for ($r = 2; $r -le $lastRow; $r++) {
    $id    = $ws.Cells.Item($r, $headers["ID"]).Value2
    $tag   = $ws.Cells.Item($r, $headers["Tag for Outreach"]).Value2
    $title = $ws.Cells.Item($r, $headers["Title"]).Value2
    $ddm1  = $ws.Cells.Item($r, $headers["DDM1"]).Value2
    $camp  = $ws.Cells.Item($r, $headers["Campaign Code"]).Value2

    if (-not $id) { continue }
    $idStr = "$([int]$id)"

    # Clean strings
    $title = if ($title) { "$title".Trim() } else { "" }
    $camp  = if ($camp)  { "$camp".Trim()  } else { "" }

    # Convert "Last, First" -> "First Last" to match campaign data format
    $ddm1 = if ($ddm1) {
        $d = "$ddm1".Trim()
        if ($d -match '^([^,]+),\s*(.+)$') { "$($Matches[2].Trim()) $($Matches[1].Trim())" } else { $d }
    } else { "" }

    # Split tag field on commas — each becomes its own entry so it can match independently
    $rawTag = if ($tag) { "$tag".Trim() } else { "" }
    $tags = if ($rawTag) {
        $rawTag -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    } else { @("") }

    $tagIdx = 0
    foreach ($t in $tags) {
        $key = if ($tags.Count -gt 1) { "${idStr}_${tagIdx}" } else { $idStr }
        $profMeta[$key] = @{ tag = $t; title = $title; ddm1 = $ddm1; camp = $camp }

        if ($t -and $camp -and $camp -ne "" -and $camp -ne "TBD" -and $camp -ne "no") {
            $profTags[$camp] = $t
        }
        $tagIdx++
    }
}

$wb.Close($false); $excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Profiling requests loaded: $($profMeta.Count)"
Write-Host "Camp->Tag mappings: $($profTags.Count)"

$metaJson = $profMeta | ConvertTo-Json -Compress -Depth 3
$tagsJson = $profTags | ConvertTo-Json -Compress -Depth 2

# Read existing data-out.js and replace lines 2 and 3
$lines = [System.IO.File]::ReadAllLines($outFile, [System.Text.Encoding]::UTF8)

# Lines are 0-indexed; line 2 = index 1 (CAMP_PROF_TAGS), line 3 = index 2 (CAMP_PROF_META)
$lines[1] = "window.CAMP_PROF_TAGS=$tagsJson;"
$lines[2] = "window.CAMP_PROF_META=$metaJson;"

[System.IO.File]::WriteAllLines($outFile, $lines, [System.Text.Encoding]::UTF8)

$kb = [Math]::Round((Get-Item $outFile).Length / 1KB)
Write-Host "data-out.js updated - ${kb}KB"
