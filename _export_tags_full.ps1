$campPath = "C:\Users\I572929\campaign-calendar-site\data-camp.js"
$profPath = "C:\Users\I572929\campaign-calendar-site\data-profiling-req-20261005.js"
$outPath  = "C:\Users\I572929\campaign-calendar-site\tag-matching.xlsx"

# ── Profiling tag normaliser (mirrors _splitDtrT in JS) ────────────────────
function Split-Tags($raw) {
    if (-not $raw) { return @() }
    return ([string]$raw) -split '[\r\n,;|\/\(\)\[\]]+|\\r\\n|\\n' | ForEach-Object {
        $t = $_.Trim()
        $t = $t -ireplace '^(?:New Prospects\s+Tag|Existing\s+Tag|New\s+Tag|Tag|Existing|New)(?::\s*|\s+)', ''
        $t = $t -ireplace '^(?:EX|NP)-', ''
        $t = $t -ireplace '[\s-]+(?:existing|new|ex|np)$', ''
        $t.Trim().ToLower()
    } | Where-Object { $_ -and $_.Length -gt 1 -and $_ -ne '-' -and $_ -ne '--' }
}

# ── Calendar tag splitter — raw values, comma-split only ───────────────────
function Split-CalTags($raw) {
    if (-not $raw) { return @() }
    return ([string]$raw) -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_.Length -gt 0 }
}

# ── Parse profiling data ───────────────────────────────────────────────────
Write-Host "Parsing profiling data..."
$profJs = [System.IO.File]::ReadAllText($profPath, [System.Text.Encoding]::UTF8)
$profStart = $profJs.IndexOf('{')
$profEnd   = $profJs.IndexOf('};' + [char]10 + 'window.')
if ($profEnd -lt 0) { $profEnd = $profJs.LastIndexOf('};') }
$profJs    = $profJs.Substring($profStart, $profEnd - $profStart + 1)
$profData  = $profJs | ConvertFrom-Json

$profTagRows = [System.Collections.Generic.List[object]]::new()
$profTagSet  = [System.Collections.Generic.HashSet[string]]::new()

foreach ($entry in $profData.PSObject.Properties.Value) {
    $id    = $entry.id
    $title = $entry.title
    $ddm1  = $entry.ddm1
    $wbs   = $entry.wbs
    foreach ($field in @('tag','tagOut')) {
        $raw = [string]($entry.$field)
        foreach ($t in (Split-Tags $raw)) {
            $profTagSet.Add($t) | Out-Null
            $profTagRows.Add([PSCustomObject]@{
                Tag       = $t
                RawValue  = $raw.Trim()
                Field     = $field
                RequestID = $id
                Title     = $title
                DDM1      = $ddm1
                WBS       = $wbs
                Matched   = ""   # filled after calendar parse
            })
        }
    }
}
Write-Host "  Profiling: $($profTagRows.Count) tag rows, $($profTagSet.Count) unique tags"

# ── Parse calendar data ────────────────────────────────────────────────────
Write-Host "Parsing calendar data..."
$campJs = [System.IO.File]::ReadAllText($campPath, [System.Text.Encoding]::UTF8)
$campJs = $campJs -replace '^window\.\w+=', '' -replace ';?\s*$', ''
$campData = $campJs | ConvertFrom-Json

$calTagRows = [System.Collections.Generic.List[object]]::new()
$calTagSet  = [System.Collections.Generic.HashSet[string]]::new()
$calTagCols = if ($campData.Count -gt 0) {
    $campData[0].PSObject.Properties.Name | Where-Object { $_ -ilike '*tag*' }
} else { @('Profiling to Outreach Tag','Tag per SDE','Tag 2 per SDE','Outreach TAGS') }
Write-Host ("  Tag columns detected: " + ($calTagCols -join ", "))

foreach ($row in $campData) {
    $wbs  = [string]$row.'Campaign/WBS Code'
    $name = [string]$row.'Campaign Name'
    $dm   = [string]$row.'Demand Manager'
    foreach ($col in $calTagCols) {
        $raw = [string]$row.$col
        foreach ($t in (Split-CalTags $raw)) {
            $calTagSet.Add($t) | Out-Null
            $calTagRows.Add([PSCustomObject]@{
                Tag          = $t
                RawValue     = $raw.Trim()
                Column       = $col
                WBS          = $wbs
                CampaignName = $name
                DCM          = $dm
                Matched      = ""   # filled below
            })
        }
    }
}
Write-Host "  Calendar: $($calTagRows.Count) tag rows, $($calTagSet.Count) unique tags"

# ── Fill Matched column ────────────────────────────────────────────────────
$calTagSetLower = [System.Collections.Generic.HashSet[string]]::new()
foreach ($t in $calTagSet) { $calTagSetLower.Add($t.ToLower()) | Out-Null }

foreach ($r in $profTagRows) { $r.Matched = if ($calTagSetLower.Contains($r.Tag)) { "Yes" } else { "No" } }
foreach ($r in $calTagRows)  { $r.Matched = if ($profTagSet.Contains($r.Tag.ToLower())) { "Yes" } else { "No" } }

$profMatched   = ($profTagRows | Where-Object { $_.Matched -eq "Yes" } | Select-Object -ExpandProperty Tag | Sort-Object -Unique).Count
$calMatched    = ($calTagRows  | Where-Object { $_.Matched -eq "Yes" } | Select-Object -ExpandProperty Tag | Sort-Object -Unique).Count
Write-Host "  Profiling unique tags matched: $profMatched / $($profTagSet.Count)"
Write-Host "  Calendar  unique tags matched: $calMatched / $($calTagSet.Count)"

# ── Sort: matched first, then alphabetical ─────────────────────────────────
$profSorted = $profTagRows | Sort-Object @{e={if($_.Matched -eq "Yes"){0}else{1}}}, Tag, RequestID
$calSorted  = $calTagRows  | Sort-Object @{e={if($_.Matched -eq "Yes"){0}else{1}}}, Tag, WBS

# ── Write Excel via COM ────────────────────────────────────────────────────
Write-Host "Creating Excel..."
$xl  = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb  = $xl.Workbooks.Add()

$navy  = 0x4A1400   # BGR for #00144A (dark navy)
$white = 16777215
$green = 0x00C000   # matched highlight (green text)
$grey  = 0x808080   # unmatched (grey text)

function Write-Sheet($wb, $sheetIndex, $sheetName, $data, $headers) {
    if ($sheetIndex -le $wb.Sheets.Count) { $ws = $wb.Sheets.Item($sheetIndex) }
    else { $ws = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count)) }
    $ws.Name = $sheetName

    # Header row
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $cell = $ws.Cells.Item(1, $c+1)
        $cell.Value2 = $headers[$c]
        $cell.Font.Bold = $true
        $cell.Interior.Color = $navy
        $cell.Font.Color = $white
    }

    # Data rows
    $matchedCol = $headers.IndexOf("Matched")
    $r = 2
    foreach ($row in $data) {
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $cell = $ws.Cells.Item($r, $c+1)
            $cell.Value2 = [string]($row.($headers[$c]))
            # Colour the Matched cell
            if ($c -eq $matchedCol) {
                if ($row.Matched -eq "Yes") {
                    $cell.Font.Color = $green; $cell.Font.Bold = $true
                } else {
                    $cell.Font.Color = $grey
                }
            }
        }
        $r++
    }
    $ws.Columns.AutoFit() | Out-Null
    return $ws
}

$hdProf = [System.Collections.ArrayList]@('Tag','Matched','RawValue','Field','RequestID','Title','DDM1','WBS')
$hdCal  = [System.Collections.ArrayList]@('Tag','Matched','RawValue','Column','WBS','CampaignName','DCM')

Write-Sheet $wb 1 'Profiling Tags' $profSorted $hdProf | Out-Null
$wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count)) | Out-Null
Write-Sheet $wb 2 'Calendar Tags'  $calSorted  $hdCal  | Out-Null

while ($wb.Sheets.Count -gt 2) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }

$wb.SaveAs($outPath, 51)
$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null

Write-Host "Saved: $outPath"
