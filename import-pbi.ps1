# import-pbi.ps1
# Reads Campaign Calendars and Profiling.xlsx and generates:
#   data-camp.js          -> window.CAMP_DATA
#   data-profiling-tags.js -> window.CAMP_PROF_TAGS
#   data-profiling-req.js  -> window.PROFILING_ACCT_REQUESTED

$xlsxPath = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Calendars\Campaign Calendars combined.xlsx"
$outDir   = "C:\Users\I572929\campaign-calendar-site"
$zipPath  = "$outDir\__pbi_import_temp.zip"
$extPath  = "$outDir\__pbi_import_extracted"

Write-Host "Copying file..."
Copy-Item $xlsxPath $zipPath -Force

Write-Host "Extracting..."
if (Test-Path $extPath) { Remove-Item $extPath -Recurse -Force }
Expand-Archive -Path $zipPath -DestinationPath $extPath -Force

# Load shared strings (explicit UTF-8 to avoid mojibake)
$ssXml = New-Object System.Xml.XmlDocument
$ssXml.Load("$extPath\xl\sharedStrings.xml")
$strings = $ssXml.sst.si | ForEach-Object {
    if ($null -ne $_.t) { $_.t }
    elseif ($_.r)       { ($_.r | ForEach-Object { $_.t }) -join '' }
    else                { '' }
}

# Load sheet
$sheetXml = New-Object System.Xml.XmlDocument
$sheetXml.Load("$extPath\xl\worksheets\sheet1.xml")
[xml]$sheet = $sheetXml
$rows = $sheet.worksheet.sheetData.row

# Helper: resolve cell value
function Get-CellValue($cell) {
    $v = $cell.v
    if ($null -eq $v) { return '' }
    if ($cell.t -eq 's') { return $strings[[int]$v] }
    # Date serial (Execution Start/End Date cols are date formatted)
    return $v
}

# Helper: convert Excel date serial to YYYY-MM-DD
function ConvertTo-ExcelDate($serial) {
    if ([string]::IsNullOrWhiteSpace($serial)) { return '' }
    try {
        $d = [double]$serial
        if ($d -lt 1 -or $d -gt 100000) { return $serial }
        # Excel epoch is 1900-01-00, with leap year bug, so subtract 2
        return ([DateTime]"1899-12-30").AddDays($d).ToString("yyyy-MM-dd")
    } catch { return $serial }
}

# Helper: escape value for JSON string
function EscapeJson($val) {
    if ($null -eq $val) { return '' }
    $s = [string]$val
    $s = $s.Replace('\', '\\').Replace('"', '\"').Replace("`r`n", ' ').Replace("`n", ' ').Replace("`r", ' ').Replace("`t", ' ')
    return $s
}

# Get headers
$headers = @{}
$row0 = $rows[0]
foreach ($cell in $row0.c) {
    $col = $cell.r -replace '\d+', ''   # e.g. "A", "B", "AA"
    $headers[$col] = Get-CellValue $cell
}

# Build column index: header name -> column letter
$colIdx = @{}
foreach ($k in $headers.Keys) { $colIdx[$headers[$k]] = $k }

# Columns needed for CAMP_DATA
$campCols = @{
    'Campaign Name'               = 'Campaign Name'
    'Campaign No'                 = 'Campaign No'
    'Campaign/WBS Code'           = 'Campaign/WBS Code'
    'Status'                      = 'Status'
    'Campaign Type'               = 'Campaign Type'
    'Campaign Priority'           = 'Campaign Priority'
    'Campaign Objective'          = 'Campaign Objective'
    'Campaign Origin'             = 'Campaign Origin'
    'Execution Start Date'        = 'Execution Start Date'
    'Execution End Date'          = 'Execution End Date'
    'Starting Quarter'            = 'Starting Quarter'
    'Region Name (level 2)'       = 'Region Name (level 2)'
    'Region Name (level 3)'       = 'Region Name (level 3)'
    'Demand Manager'              = 'Demand Manager'
    'Sales Bag'                   = 'Sales Bag'
    'Sub Sales Bag'               = 'Sub Sales Bag'
    'IAC'                         = 'IAC'
    'SoD'                         = 'SoD'
    'IB/NNN'                      = 'IB/NNN'
    'Industry (MC)'               = 'Industry (MC)'
    'Activity Sub-Type'           = 'Activity Sub-Type'
    'Target Pipeline - value kEUR'= 'Target Pipeline - value kEUR'
    'Total Account (Validated)'   = 'Total Account (Validated)'
    'DG PROGRAM'                  = 'DG Program'
    'Executor.title'              = 'Executor'
    'Sequence ID'                 = 'Sequence ID'
    'Digital Assets team support' = 'Digital Assets team support'
    'Existing Prospects by SDE'   = 'Existing Prospects by SDE'
    'Profiling CRM Outreach'      = 'Profiling CRM Outreach'
    'Profiling CRM Nuevo'         = 'Profiling CRM Nuevo'
    'Profiling NO CRM Accounts'   = 'Profiling NO CRM Accounts'
    'Outreach TAGS'               = 'Outreach TAGS'
    'Accounts enriched after complete Profiling' = 'Accounts enriched after complete Profiling'
    'Index'                       = 'Index'
}

$dateColsSrc = @('Execution Start Date', 'Execution End Date')

# Profiling request columns (for CAMP_PROF_TAGS and PROFILING_ACCT_REQUESTED)
# CAMP_PROF_TAGS = { WBSCode -> OutreachTag }
# PROFILING_ACCT_REQUESTED = sum of Number of Accounts Requested

Write-Host "Processing $($rows.Count - 1) data rows..."

$campRows = [System.Collections.Generic.List[string]]::new()
$profTags = [System.Collections.Generic.Dictionary[string,string]]::new()
$seenProfIds = [System.Collections.Generic.Dictionary[string,bool]]::new()
$totalAcctRequested = 0

for ($i = 1; $i -lt $rows.Count; $i++) {
    $row = $rows[$i]

    # Build a lookup of col letter -> value for this row
    $rowVals = @{}
    foreach ($cell in $row.c) {
        $col = $cell.r -replace '\d+', ''
        $rowVals[$col] = Get-CellValue $cell
    }

    # Helper to get value by header name
    function GetByHeader($hdrName) {
        $letter = $colIdx[$hdrName]
        if ($null -eq $letter) { return '' }
        $v = $rowVals[$letter]
        if ($null -eq $v) { return '' }
        return $v
    }

    # Skip blank rows (no Campaign Name)
    $campName = GetByHeader 'Campaign Name'
    if ([string]::IsNullOrWhiteSpace($campName)) { continue }

    # Build camp JSON object
    $fields = [System.Collections.Generic.List[string]]::new()
    foreach ($srcHdr in $campCols.Keys) {
        $destKey = $campCols[$srcHdr]
        $val = GetByHeader $srcHdr
        if ($dateColsSrc -contains $srcHdr) { $val = ConvertTo-ExcelDate $val }
        $fields.Add('"' + (EscapeJson $destKey) + '":"' + (EscapeJson $val) + '"')
    }
    $campRows.Add('{' + ($fields -join ',') + '}')

    # CAMP_PROF_TAGS: WBS -> Outreach Tag (use Outreach TAGS first, fall back to Profiling request.Tag for Outreach)
    $wbs = GetByHeader 'Campaign/WBS Code'
    $tag = GetByHeader 'Outreach TAGS'
    if ([string]::IsNullOrWhiteSpace($tag)) { $tag = GetByHeader 'Profiling request.Tag for Outreach' }
    if (-not [string]::IsNullOrWhiteSpace($wbs) -and -not [string]::IsNullOrWhiteSpace($tag)) {
        if (-not $profTags.ContainsKey($wbs)) {
            $profTags[$wbs] = $tag
        }
    }

    # PROFILING_ACCT_REQUESTED: deduplicated by Profiling request.ID
    $profReqId = GetByHeader 'Profiling request.ID'
    $acctReq   = GetByHeader 'Profiling request.Number of Accounts Requested'
    if (-not [string]::IsNullOrWhiteSpace($profReqId) -and -not [string]::IsNullOrWhiteSpace($acctReq)) {
        if (-not $seenProfIds.ContainsKey($profReqId)) {
            $seenProfIds[$profReqId] = $true
            try { $totalAcctRequested += [int][double]$acctReq } catch {}
        }
    }
}

Write-Host "Writing data-camp.js ($($campRows.Count) rows)..."
$campJs = 'window.CAMP_DATA=[' + ($campRows -join ',') + '];'
[System.IO.File]::WriteAllText("$outDir\data-camp.js", $campJs, [System.Text.Encoding]::UTF8)

Write-Host "Writing data-profiling-tags.js ($($profTags.Count) entries)..."
$tagPairs = $profTags.Keys | ForEach-Object { '"' + (EscapeJson $_) + '":"' + (EscapeJson $profTags[$_]) + '"' }
$tagsJs = 'window.CAMP_PROF_TAGS={' + ($tagPairs -join ',') + '};'
[System.IO.File]::WriteAllText("$outDir\data-profiling-tags.js", $tagsJs, [System.Text.Encoding]::UTF8)

Write-Host "Writing data-profiling-req.js (total=$totalAcctRequested)..."
$reqJs = "window.PROFILING_ACCT_REQUESTED=$totalAcctRequested;"
[System.IO.File]::WriteAllText("$outDir\data-profiling-req.js", $reqJs, [System.Text.Encoding]::UTF8)

# Cleanup
Remove-Item $zipPath -Force
Remove-Item $extPath -Recurse -Force

Write-Host "Done."
