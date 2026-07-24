# import-seqprospect.ps1
# Reads Prospects_Export.xlsx and builds data-seqprospect.js (window.SEQ_PROSPECT_DATA)
# Uses zip/XML approach — no Excel COM required

$xlsxPath = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Outreach\Prospects_Export.xlsx"
$outDir   = "C:\Users\I572929\campaign-calendar-site"
$zipPath  = "$outDir\__pros_temp.zip"
$extPath  = "$outDir\__pros_extracted"

Write-Host "Copying file..."
Copy-Item $xlsxPath $zipPath -Force

Write-Host "Extracting..."
if (Test-Path $extPath) { Remove-Item $extPath -Recurse -Force }
Expand-Archive -Path $zipPath -DestinationPath $extPath -Force

$ssXml = New-Object System.Xml.XmlDocument
$ssXml.Load("$extPath\xl\sharedStrings.xml")
$strings = $ssXml.sst.si | ForEach-Object {
    if ($null -ne $_.t) { $_.t }
    elseif ($_.r) { ($_.r | ForEach-Object { $_.t }) -join '' }
    else { '' }
}

$sheetXml = New-Object System.Xml.XmlDocument
$sheetXml.Load("$extPath\xl\worksheets\sheet1.xml")
$rows = $sheetXml.worksheet.sheetData.row

function Get-CellValue($cell) {
    $v = $cell.v
    if ($null -eq $v) { return '' }
    if ($cell.t -eq 's') { return $strings[[int]$v] }
    return $v
}

function EscapeJson($val) {
    if ($null -eq $val) { return '' }
    $s = [string]$val
    $s = $s.Replace('\','\\').Replace('"','\"').Replace("`r`n",' ').Replace("`n",' ').Replace("`r",' ').Replace("`t",' ')
    return $s
}

# Build header index
$headers = @{}
foreach ($cell in $rows[0].c) {
    $col = $cell.r -replace '\d+', ''
    $headers[$col] = Get-CellValue $cell
}
$colIdx = @{}
foreach ($k in $headers.Keys) { $colIdx[$headers[$k]] = $k }

Write-Host "Headers: $($colIdx.Keys.Count) columns found"

function GetByCol($rowVals, $hdrName) {
    $letter = $colIdx[$hdrName]
    if ($null -eq $letter) { return '' }
    $v = $rowVals[$letter]
    if ($null -eq $v) { return '' }
    return [string]$v
}

# Build name->ID map from Sequence Stats CSV
$seqStatsCsv = Get-ChildItem "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Outreach\" -Filter "Sequence_Stats*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $seqStatsCsv) { Write-Host "ERROR: No Sequence_Stats CSV found"; exit 1 }
Write-Host "Using Sequence Stats: $($seqStatsCsv.Name)"
$seqStats = Import-Csv $seqStatsCsv.FullName -Encoding UTF8
$nameToId = @{}
foreach ($row in $seqStats) {
    $sid = $row."Sequence ID".Trim()
    $sname = $row."Sequence Name".Trim()
    if ($sid -and $sname -and -not $nameToId.ContainsKey($sname)) { $nameToId[$sname] = $sid }
}
Write-Host "Sequence name->ID map: $($nameToId.Count) entries"

# Build tracked sequence IDs — include ALL sequences from Sequence Stats
$trackedSids = New-Object System.Collections.Generic.HashSet[string]
foreach ($row in $seqStats) {
    $sid = $row."Sequence ID".Trim()
    if ($sid) { [void]$trackedSids.Add($sid) }
}
# Also add from campaign calendar
$campJs = [System.IO.File]::ReadAllText("$outDir\data-camp.js")
[regex]::Matches($campJs, '"Sequence ID":"([^"]+)"') | ForEach-Object { [void]$trackedSids.Add($_.Groups[1].Value.Trim()) }
# Also from outreach.json
if (Test-Path "$outDir\outreach.json") {
    $outJson = Get-Content "$outDir\outreach.json" -Raw | ConvertFrom-Json
    $outJson | ForEach-Object { [void]$trackedSids.Add("$($_.'Sequence ID')".Trim()) }
}
Write-Host "Tracked sequences: $($trackedSids.Count)"

Write-Host "Processing $($rows.Count - 1) data rows..."
$entries  = [System.Collections.Generic.List[string]]::new()
$seenPid  = New-Object System.Collections.Generic.HashSet[string]
$totalRows = 0; $matchedRows = 0

for ($i = 1; $i -lt $rows.Count; $i++) {
    $row = $rows[$i]
    $rowVals = @{}
    foreach ($cell in $row.c) {
        $col = $cell.r -replace '\d+', ''
        $rowVals[$col] = Get-CellValue $cell
    }

    $prospId  = [string](GetByCol $rowVals 'ID')
    if (-not $prospId) { continue }
    $totalRows++

    $active   = [string](GetByCol $rowVals 'Active Sequences')
    $finished = [string](GetByCol $rowVals 'Finished Sequences')
    $company  = EscapeJson (GetByCol $rowVals 'Company')
    $touched  = EscapeJson (GetByCol $rowVals 'Touched At')
    $source   = EscapeJson (GetByCol $rowVals 'Source')
    $coType   = EscapeJson (GetByCol $rowVals 'Company Type')
    $cf51     = EscapeJson (GetByCol $rowVals 'Custom Field 51')
    $country  = EscapeJson (GetByCol $rowVals 'Country')
    $created  = EscapeJson (GetByCol $rowVals 'Created At')

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
                $entries.Add("{""sid"":""$sid"",""pid"":""$prospId"",""co"":""$company"",""touched"":""$touched"",""src"":""$source"",""ct"":""$coType"",""cf51"":""$cf51"",""country"":""$country"",""created"":""$created""}")
                $matchedRows++
            }
        }
    }

    if ($i % 10000 -eq 0) { Write-Host "Processed $i / $($rows.Count-1) rows, $matchedRows entries so far" }
}

Write-Host "Total rows: $totalRows | Matched entries: $matchedRows"
$js = "window.SEQ_PROSPECT_DATA=[" + ($entries -join ",") + "];"
[System.IO.File]::WriteAllText("$outDir\data-seqprospect.js", $js, [System.Text.Encoding]::UTF8)
Write-Host "Written data-seqprospect.js ($([Math]::Round($js.Length/1KB))KB)"

Remove-Item $zipPath -Force
Remove-Item $extPath -Recurse -Force
Write-Host "Done."

# Also refresh data-prospects.js (tags, touch dates etc)
Write-Host "`nRefreshing data-prospects.js..."
& "$outDir\import-prospects.ps1"
