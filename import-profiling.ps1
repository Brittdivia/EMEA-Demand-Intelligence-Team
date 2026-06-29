# import-profiling.ps1
# Reads Profiling request.xlsx and generates data-profiling-req.js
# Produces: window.CAMP_PROF_META = { [ID]: { id, title, tag, tagOut, wbs, ddm1, createdBy, created, acctRequested } }
# Also produces: window.PROFILING_ACCT_REQUESTED = sum of unique acctRequested

$xlsxPath = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Profiling request.xlsx"
$outDir   = "C:\Users\I572929\campaign-calendar-site"
$zipPath  = "$outDir\__prof_temp.zip"
$extPath  = "$outDir\__prof_extracted"

Write-Host "Copying file..."
Copy-Item $xlsxPath $zipPath -Force

Write-Host "Extracting..."
if (Test-Path $extPath) { Remove-Item $extPath -Recurse -Force }
Expand-Archive -Path $zipPath -DestinationPath $extPath -Force

$ssXml = New-Object System.Xml.XmlDocument
$ssXml.Load("$extPath\xl\sharedStrings.xml")
$strings = $ssXml.sst.si | ForEach-Object {
    if ($null -ne $_.t) { $_.t }
    elseif ($_.r)       { ($_.r | ForEach-Object { $_.t }) -join '' }
    else                { '' }
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

function ConvertTo-ExcelDate($serial) {
    if ([string]::IsNullOrWhiteSpace($serial)) { return '' }
    # Try numeric serial first
    try {
        $d = [double]$serial
        if ($d -gt 1 -and $d -lt 200000) {
            return ([DateTime]"1899-12-30").AddDays($d).ToString("yyyy-MM-dd")
        }
    } catch {}
    # Try parsing as date string (e.g. "6/16/2026 11:23 AM")
    try {
        $part = $serial.Trim().Split(' ')[0]
        $dt = [DateTime]::ParseExact($part, @('M/d/yyyy','MM/dd/yyyy','d/M/yyyy'), [System.Globalization.CultureInfo]::InvariantCulture, 'None')
        return $dt.ToString("yyyy-MM-dd")
    } catch {}
    try {
        $dt = [DateTime]::Parse($serial, [System.Globalization.CultureInfo]::InvariantCulture)
        return $dt.ToString("yyyy-MM-dd")
    } catch {}
    return $serial
}

# Build header index
$headers = @{}
foreach ($cell in $rows[0].c) {
    $col = $cell.r -replace '\d+', ''
    $headers[$col] = Get-CellValue $cell
}
$colIdx = @{}
foreach ($k in $headers.Keys) { $colIdx[$headers[$k]] = $k }

Write-Host "Processing $($rows.Count - 1) data rows..."

$entries  = [System.Collections.Generic.List[string]]::new()
$totalAcct = 0

for ($i = 1; $i -lt $rows.Count; $i++) {
    $row = $rows[$i]

    $rowVals = @{}
    foreach ($cell in $row.c) {
        $col = $cell.r -replace '\d+', ''
        $rowVals[$col] = Get-CellValue $cell
    }

    function GetByHeader($hdrName) {
        $letter = $colIdx[$hdrName]
        if ($null -eq $letter) { return '' }
        $v = $rowVals[$letter]
        if ($null -eq $v) { return '' }
        return $v
    }

    $id = GetByHeader 'ID'
    if ([string]::IsNullOrWhiteSpace($id)) { continue }

    $title       = GetByHeader 'Title'
    $tag         = GetByHeader 'Tag of Prospects'
    $tagOut      = GetByHeader 'Tag for Outreach'
    $wbs         = GetByHeader 'Campaign Code'
    $ddm1        = GetByHeader 'DDM1'
    $createdBy   = GetByHeader 'Created By'
    $created     = ConvertTo-ExcelDate (GetByHeader 'Created')
    $acctReqRaw  = GetByHeader 'Number of Accounts Requested'
    $acctReq     = 0
    if (-not [string]::IsNullOrWhiteSpace($acctReqRaw)) {
        try { $acctReq = [int][double]$acctReqRaw } catch {}
    }
    $totalAcct += $acctReq

    $entry = '"' + (EscapeJson $id) + '":{'
    $entry += '"id":"'          + (EscapeJson $id)        + '",'
    $entry += '"title":"'       + (EscapeJson $title)     + '",'
    $entry += '"tag":"'         + (EscapeJson $tag)       + '",'
    $entry += '"tagOut":"'      + (EscapeJson $tagOut)    + '",'
    $entry += '"wbs":"'         + (EscapeJson $wbs)       + '",'
    $entry += '"ddm1":"'        + (EscapeJson $ddm1)      + '",'
    $entry += '"createdBy":"'   + (EscapeJson $createdBy) + '",'
    $entry += '"created":"'     + (EscapeJson $created)   + '",'
    $entry += '"acctRequested":' + $acctReq
    $entry += '}'
    $entries.Add($entry)
}

Write-Host "Writing data-profiling-req.js ($($entries.Count) entries, total acct=$totalAcct)..."
$js  = 'window.CAMP_PROF_META={' + ($entries -join ',') + '};'
$js += "`nwindow.PROFILING_ACCT_REQUESTED=$totalAcct;"
[System.IO.File]::WriteAllText("$outDir\data-profiling-req.js", $js, [System.Text.Encoding]::UTF8)

# Cleanup
Remove-Item $zipPath -Force
Remove-Item $extPath -Recurse -Force

Write-Host "Done."
