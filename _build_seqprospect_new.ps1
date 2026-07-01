$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open("C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Outreach\Prospects_Export.xlsx")
$ws = $wb.Sheets.Item(1)
$lastRow = $ws.UsedRange.Rows.Count
$lastCol = $ws.UsedRange.Columns.Count

# Map headers to column indices
$colMap = @{}
for ($c = 1; $c -le $lastCol; $c++) {
    $h = $ws.Cells.Item(1, $c).Text
    $colMap[$h] = $c
}
Write-Host "Headers found: $($colMap.Keys -join ' | ')"

# Build name->ID map from Sequence Stats CSV
$seqStats = Import-Csv "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Outreach\Sequence_Stats_2026-06-02.csv"
$nameToId = @{}
foreach ($row in $seqStats) {
    $sid = $row."Sequence ID".Trim()
    $sname = $row."Sequence Name".Trim()
    if ($sid -and $sname -and -not $nameToId.ContainsKey($sname)) {
        $nameToId[$sname] = $sid
    }
}
Write-Host "Sequence Stats name->ID map: $($nameToId.Count) entries"

# Filter to only sequences tracked in outreach.json (campaign-matched)
$outJson = Get-Content "C:\Users\I572929\campaign-calendar-site\outreach.json" -Raw | ConvertFrom-Json
$trackedSids = New-Object System.Collections.Generic.HashSet[string]
$outJson | ForEach-Object { [void]$trackedSids.Add("$($_.'Sequence ID')".Trim()) }
Write-Host "Tracked sequences in outreach.json: $($trackedSids.Count)"

# Column indices
$colId       = $colMap["ID"]
$colCompany  = $colMap["Company"]
$colTouched  = $colMap["Touched At"]
$colActive   = $colMap["Active Sequences"]
$colFinished = $colMap["Finished Sequences"]
$colAcctId   = $colMap["Account ID"]
$colSource   = $colMap["Source"]
$colCoType   = $colMap["Company Type"]

if (-not $colId)       { Write-Host "ERROR: 'ID' column not found"; $wb.Close($false); $excel.Quit(); exit 1 }
if (-not $colActive)   { Write-Host "ERROR: 'Active Sequences' column not found"; $wb.Close($false); $excel.Quit(); exit 1 }
if (-not $colFinished) { Write-Host "ERROR: 'Finished Sequences' column not found"; $wb.Close($false); $excel.Quit(); exit 1 }

# Build prospect entries - one entry per prospect per matched sequence
$entries = [System.Collections.Generic.List[string]]::new()
$seenPid = New-Object System.Collections.Generic.HashSet[string]
$totalRows = 0; $matchedRows = 0

for ($r = 2; $r -le $lastRow; $r++) {
    $prospId     = $ws.Cells.Item($r, $colId).Text.Trim()
    $company = $ws.Cells.Item($r, $colCompany).Text.Trim().Replace('\','\\').Replace('"','\"')
    $touched = $ws.Cells.Item($r, $colTouched).Text.Trim()
    $active  = $ws.Cells.Item($r, $colActive).Text.Trim()
    $finished = $ws.Cells.Item($r, $colFinished).Text.Trim()
    $acctId  = $ws.Cells.Item($r, $colAcctId).Text.Trim()
    $source  = if ($colSource) { $ws.Cells.Item($r, $colSource).Text.Trim().Replace('\','\\').Replace('"','\"') } else { "" }
    $coType  = if ($colCoType) { $ws.Cells.Item($r, $colCoType).Text.Trim().Replace('\','\\').Replace('"','\"') } else { "" }

    if (-not $prospId) { continue }
    $totalRows++

    # Collect all sequence names from both columns (comma or semicolon separated)
    $allSeqs = @()
    if ($active)   { $allSeqs += $active   -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ } }
    if ($finished) { $allSeqs += $finished -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ } }

    foreach ($seqName in $allSeqs) {
        if ($nameToId.ContainsKey($seqName)) {
            $sid = $nameToId[$seqName]
            if (-not $trackedSids.Contains($sid)) { continue }
            $key = "$prospId|$sid"
            if (-not $seenPid.Contains($key)) {
                $seenPid.Add($key) | Out-Null
                $entries.Add("{""sid"":""$sid"",""pid"":""$prospId"",""co"":""$company"",""aid"":""$acctId"",""touched"":""$touched"",""src"":""$source"",""ct"":""$coType""}")
                $matchedRows++
            }
        }
    }

    if ($r % 5000 -eq 0) { Write-Host "Processed $r / $lastRow rows, $matchedRows entries so far" }
}

Write-Host "Total prospect rows: $totalRows"
Write-Host "Matched entries: $matchedRows"

$js = "window.SEQ_PROSPECT_DATA=[" + ($entries -join ",") + "];"
[System.IO.File]::WriteAllText("C:\Users\I572929\campaign-calendar-site\data-seqprospect.js", $js, [System.Text.Encoding]::UTF8)
Write-Host "Written: $($js.Length) bytes"

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
