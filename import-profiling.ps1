# import-profiling.ps1
# Reads Profiling request.csv and generates data-profiling-req.js
# Produces: window.CAMP_PROF_META = { [ID]: { id, title, tag, tagOut, wbs, ddm1, createdBy, created, acctRequested } }
# Also produces: window.PROFILING_ACCT_REQUESTED = sum of unique acctRequested

$csvPath = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Profiling request.csv"
$outDir  = "C:\Users\I572929\campaign-calendar-site"

function EscapeJson($val) {
    if ($null -eq $val) { return '' }
    $s = [string]$val
    $s = $s.Replace('\','\\').Replace('"','\"').Replace("`r`n",' ').Replace("`n",' ').Replace("`r",' ').Replace("`t",' ')
    return $s
}

function ConvertToDate($serial) {
    if ([string]::IsNullOrWhiteSpace($serial)) { return '' }
    $s = $serial.Trim()
    # Numeric Excel serial
    $d = 0.0
    if ([double]::TryParse($s, [ref]$d) -and $d -gt 1 -and $d -lt 200000) {
        return ([DateTime]"1899-12-30").AddDays($d).ToString("yyyy-MM-dd")
    }
    # Date string — extract date part before any space (e.g. "6/16/2026 11:07 AM" -> "6/16/2026")
    $part = $s.Split(' ')[0]
    $dt = [DateTime]::MinValue
    # M/d/yyyy (US format — confirmed format in this CSV)
    $fmts = @('M/d/yyyy','MM/dd/yyyy','d/M/yyyy','dd/MM/yyyy','yyyy-MM-dd')
    foreach ($fmt in $fmts) {
        if ([DateTime]::TryParseExact($part, $fmt, [System.Globalization.CultureInfo]::InvariantCulture, 'None', [ref]$dt)) {
            return $dt.ToString("yyyy-MM-dd")
        }
    }
    # Last resort
    if ([DateTime]::TryParse($s, [ref]$dt)) { return $dt.ToString("yyyy-MM-dd") }
    return ''
}

Write-Host "Reading CSV..."
$rows = Import-Csv -Path $csvPath -Encoding UTF8

Write-Host "Processing $($rows.Count) rows..."

$entries  = [System.Collections.Generic.List[string]]::new()
$totalAcct = 0

foreach ($row in $rows) {
    $id = $row.ID
    if ([string]::IsNullOrWhiteSpace($id)) { continue }

    $title      = $row.Title
    $tag        = $row.'Tag of Prospects'
    $tagOut     = $row.'Tag for Outreach'
    $wbs        = $row.'Campaign Code'
    $ddm1       = $row.DDM1
    $createdBy  = $row.'Created By'
    $created    = ConvertToDate $row.Created
    $acctReq    = 0
    $acctRaw    = $row.'Number of Accounts Requested'
    if (-not [string]::IsNullOrWhiteSpace($acctRaw)) {
        try { $acctReq = [int][double]$acctRaw } catch {}
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

Write-Host "Done."
