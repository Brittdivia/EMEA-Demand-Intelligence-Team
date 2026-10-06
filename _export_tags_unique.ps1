$campPath = "C:\Users\I572929\campaign-calendar-site\data-camp.js"
$profPath = "C:\Users\I572929\campaign-calendar-site\data-profiling-req-20261005.js"
$outPath  = "C:\Users\I572929\campaign-calendar-site\tag-overview.xlsx"

# ── Profiling tag normaliser ───────────────────────────────────────────────
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

# ── Calendar tag splitter — raw, comma-split only ──────────────────────────
function Split-CalTags($raw) {
    if (-not $raw) { return @() }
    return ([string]$raw) -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_.Length -gt 0 }
}

# ── Parse profiling ────────────────────────────────────────────────────────
Write-Host "Parsing profiling data..."
$profJs = [System.IO.File]::ReadAllText($profPath, [System.Text.Encoding]::UTF8)
$profStart = $profJs.IndexOf('{')
$profEnd   = $profJs.IndexOf('};' + [char]10 + 'window.')
if ($profEnd -lt 0) { $profEnd = $profJs.LastIndexOf('};') }
$profJs    = $profJs.Substring($profStart, $profEnd - $profStart + 1)
$profData  = $profJs | ConvertFrom-Json

$profTagSet = [System.Collections.Generic.HashSet[string]]::new()
foreach ($entry in $profData.PSObject.Properties.Value) {
    foreach ($field in @('tag','tagOut','tagEnr','tagWave2')) {
        foreach ($t in (Split-Tags ([string]($entry.$field)))) {
            $profTagSet.Add($t) | Out-Null
        }
    }
}
Write-Host "  Profiling: $($profTagSet.Count) unique tags"

# ── Parse calendar ─────────────────────────────────────────────────────────
Write-Host "Parsing calendar data..."
$campJs = [System.IO.File]::ReadAllText($campPath, [System.Text.Encoding]::UTF8)
$campJs = $campJs -replace '^window\.\w+=', '' -replace ';?\s*$', ''
$campData = $campJs | ConvertFrom-Json

$calTagCols = if ($campData.Count -gt 0) {
    $campData[0].PSObject.Properties.Name | Where-Object { $_ -ilike '*tag*' }
} else { @('Profiling to Outreach Tag','Tag per SDE','Tag 2 per SDE') }
Write-Host "  Tag columns: $($calTagCols -join ', ')"

$calTagSet = [System.Collections.Generic.HashSet[string]]::new()
foreach ($row in $campData) {
    foreach ($col in $calTagCols) {
        foreach ($t in (Split-CalTags ([string]$row.$col))) {
            $calTagSet.Add($t) | Out-Null
        }
    }
}
Write-Host "  Calendar: $($calTagSet.Count) unique tags"

# ── Build unique tag rows ──────────────────────────────────────────────────
$calTagSetLower = [System.Collections.Generic.HashSet[string]]::new()
foreach ($t in $calTagSet) { $calTagSetLower.Add($t.ToLower()) | Out-Null }

$calRows  = $calTagSet  | Sort-Object | ForEach-Object {
    [PSCustomObject]@{ Tag = $_; Matched = if ($profTagSet.Contains($_.ToLower())) { "Yes" } else { "No" } }
} | Sort-Object @{e={if($_.Matched -eq "Yes"){0}else{1}}}, Tag

$profRows = $profTagSet | Sort-Object | ForEach-Object {
    [PSCustomObject]@{ Tag = $_; Matched = if ($calTagSetLower.Contains($_)) { "Yes" } else { "No" } }
} | Sort-Object @{e={if($_.Matched -eq "Yes"){0}else{1}}}, Tag

$calMatched  = ($calRows  | Where-Object { $_.Matched -eq "Yes" }).Count
$profMatched = ($profRows | Where-Object { $_.Matched -eq "Yes" }).Count
Write-Host "  Calendar matched: $calMatched / $($calTagSet.Count)"
Write-Host "  Profiling matched: $profMatched / $($profTagSet.Count)"

# ── Write Excel ────────────────────────────────────────────────────────────
Write-Host "Creating Excel..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()

$navy  = 0x4A1400
$white = 16777215
$green = 0x00C000
$grey  = 0x808080

function Write-Sheet($ws, $name, $rows) {
    $ws.Name = $name
    # Header
    foreach ($ci in @(1,2)) {
        $h = $ws.Cells.Item(1, $ci)
        $h.Value2 = @("Tag","Matched")[$ci-1]
        $h.Font.Bold = $true
        $h.Interior.Color = $navy
        $h.Font.Color = $white
    }
    $r = 2
    foreach ($row in $rows) {
        $ws.Cells.Item($r,1).Value2 = $row.Tag
        $cell = $ws.Cells.Item($r,2)
        $cell.Value2 = $row.Matched
        if ($row.Matched -eq "Yes") { $cell.Font.Color = $green; $cell.Font.Bold = $true }
        else                        { $cell.Font.Color = $grey }
        $r++
    }
    $ws.Columns.AutoFit() | Out-Null
}

$ws1 = $wb.Sheets.Item(1)
Write-Sheet $ws1 "Calendar Tags" $calRows

$ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $ws1)
Write-Sheet $ws2 "Profiling Tags" $profRows

while ($wb.Sheets.Count -gt 2) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }

$wb.SaveAs($outPath, 51)
$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null

Write-Host "Saved: $outPath"
