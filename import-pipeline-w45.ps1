# import-pipeline-w45.ps1
# Reads Week 45.csv and generates data-pipe-w45.js + data-pipe.js
# Handles column offset: some regions are missing the 'AI Flag' column,
# shifting all cols from index 3 onwards by -1. Offset is auto-detected per row
# by finding which column contains the Opportunity ID (006... pattern).

$csvPath = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Pipeline\Week 45.csv"
$outDir  = "C:\Users\I572929\campaign-calendar-site"

function EscapeJson($val) {
    if ($null -eq $val) { return '' }
    $s = [string]$val
    $s = $s.Replace('\','\\').Replace('"','\"').Replace("`r`n",' ').Replace("`n",' ').Replace("`r",' ').Replace("`t",' ')
    return $s
}

function ParseDate($d) {
    if ([string]::IsNullOrWhiteSpace($d)) { return '' }
    if ($d -match '^\d{1,2}/\d{1,2}/\d{4}$') {
        try { return [DateTime]::ParseExact($d, "dd/MM/yyyy", [System.Globalization.CultureInfo]::InvariantCulture).ToString("yyyy-MM-dd") } catch {}
        try { return [DateTime]::ParseExact($d, "d/M/yyyy",  [System.Globalization.CultureInfo]::InvariantCulture).ToString("yyyy-MM-dd") } catch {}
    }
    if ($d -match '^\d{4}-\d{2}-\d{2}$') { return $d }
    return $d
}

Write-Host "Reading CSV with TextFieldParser..."
Add-Type -AssemblyName Microsoft.VisualBasic
$parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($csvPath, [System.Text.Encoding]::UTF8)
$parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
$parser.SetDelimiters(',')
$parser.HasFieldsEnclosedInQuotes = $true

# Skip junk row 0
$parser.ReadLine() | Out-Null
# Read real headers from row 1
$hdrs = $parser.ReadFields()
for ($i = 0; $i -lt $hdrs.Count; $i++) { $hdrs[$i] = $hdrs[$i].Trim() }
if ([string]::IsNullOrWhiteSpace($hdrs[$hdrs.Count - 1])) { $hdrs[$hdrs.Count - 1] = 'Eur' }
Write-Host "Headers ($($hdrs.Count)): last=[$($hdrs[$hdrs.Count-1])]"

# Build header→index map
$hdrIdx = @{}
for ($i = 0; $i -lt $hdrs.Count; $i++) { $hdrIdx[$hdrs[$i]] = $i }
$oppHdrIdx = $hdrIdx['Opportunity ID']  # = 20

# GetV: get value from $vals by header name, applying column offset for shifted rows
# Cols 0-2 are never shifted; cols 3+ use (hdrIndex + offset)
function GetV($vals, $hdrName, $offset) {
    $bi = $hdrIdx[$hdrName]
    if ($null -eq $bi) { return '' }
    $idx = if ($bi -le 2) { $bi } else { $bi + $offset }
    if ($idx -ge 0 -and $idx -lt $vals.Count) { return ([string]$vals[$idx]).Trim() } else { return '' }
}

# Output columns (same schema as prior weeks)
$COLS = @(
    "Region Lvl 2","Region Lvl 3","DRM Category","Account Name","Opportunity ID",
    "Closing Qtr","IAC (Engagement Model)","Opp Campaign ID","Solution Area (L1)","Sub-Solution Area (L2)",
    "Create Quarter","IC_ACCOUNT_SRG_TEXT","MM Identifier","SDE Handover Date",
    "SDE Initial Engagement Name","SDE Primary Engagement Name","SDE Secondary Engagement Name",
    "SDE Handover Role","Opp Description","TCP Pipeline Source Desc","SAP Mastercode",
    "Create Date","SDE Engagement Type","Untouched Days","Channel Partner","RBC",
    "Eur","MM Pipe Source","Local Account Owner Name"
)

Write-Host "Processing rows..."
$seen     = [System.Collections.Generic.HashSet[string]]::new()
$dataRows = [System.Collections.Generic.List[object[]]]::new()
$totalRead = 0; $skippedNoOpp = 0; $offsetCounts = @{0=0;'-1'=0;'other'=0}

while (!$parser.EndOfData) {
    try { $vals = $parser.ReadFields() } catch { continue }
    if ($null -eq $vals) { continue }
    $totalRead++

    # No column shift in cleaned-up file — all rows have consistent layout
    $offset = 0

    # Read Opportunity ID using header-mapped position
    $oppId = if ($vals.Count -gt $oppHdrIdx) { $vals[$oppHdrIdx].Trim() } else { '' }
    if ([string]::IsNullOrWhiteSpace($oppId)) { $skippedNoOpp++; continue }

    if ($offset -eq 0) { $offsetCounts[0]++ }
    elseif ($offset -eq -1) { $offsetCounts['-1']++ }
    else { $offsetCounts['other']++ }

    $subSol = GetV $vals 'Sub-Solution Area (L2)' $offset
    $key = $oppId + '|' + $subSol
    if (!$seen.Add($key)) { continue }

    $gbId = GetV $vals 'GB Identifier' $offset
    $mmId = $gbId

    $eur = 0.0
    $eurRaw = GetV $vals 'Eur' $offset
    if ($eurRaw) { [double]::TryParse($eurRaw, [ref]$eur) | Out-Null }

    $untouched = $null
    $udRaw = GetV $vals 'Untouched Days' $offset
    if ($udRaw) {
        $ud = 0.0
        if ([double]::TryParse($udRaw, [ref]$ud)) { $untouched = [int]$ud }
    }

    $vals2 = [object[]]@(
        (EscapeJson (GetV $vals 'Region Lvl 2' $offset))
        (EscapeJson (GetV $vals 'Region Lvl 3' $offset))
        (EscapeJson (GetV $vals 'DRM Category' $offset))
        (EscapeJson (GetV $vals 'Account Name' $offset))
        (EscapeJson $oppId)
        (EscapeJson (GetV $vals 'Closing Qtr' $offset))
        (EscapeJson (GetV $vals 'IAC (Engagement Model)' $offset))
        (EscapeJson (GetV $vals 'Opp Campaign ID' $offset))
        (EscapeJson (GetV $vals 'Solution Area (L1)' $offset))
        (EscapeJson $subSol)
        (EscapeJson (GetV $vals 'Create Quarter' $offset))
        (EscapeJson (GetV $vals 'IC_ACCOUNT_SRG_TEXT' $offset))
        (EscapeJson $mmId)
        (EscapeJson (ParseDate (GetV $vals 'SDE Handover Date' $offset)))
        (EscapeJson (GetV $vals 'SDE Initial Engagement Name' $offset))
        (EscapeJson (GetV $vals 'SDE Primary Engagement Name' $offset))
        (EscapeJson (GetV $vals 'SDE Secondary Engagement Name' $offset))
        (EscapeJson (GetV $vals 'SDE Handover Role' $offset))
        (EscapeJson (GetV $vals 'Opp Description' $offset))
        (EscapeJson (GetV $vals 'TCP Pipeline Source Desc' $offset))
        (EscapeJson (GetV $vals 'SAP Mastercode' $offset))
        (EscapeJson (ParseDate (GetV $vals 'Create Date' $offset)))
        (EscapeJson (GetV $vals 'SDE Engagement Type' $offset))
        $untouched
        (EscapeJson (GetV $vals 'Channel Partner' $offset))
        (EscapeJson (GetV $vals 'RBC' $offset))
        $eur
        (EscapeJson (GetV $vals 'MM Pipe Source' $offset))
        (EscapeJson (GetV $vals 'Local Acct Owner' $offset))
    )
    $dataRows.Add($vals2)
}
$parser.Dispose()

Write-Host "Total rows read: $totalRead | skipped (no OppId): $skippedNoOpp"
Write-Host "Offset distribution: offset=0: $($offsetCounts[0]) | offset=-1: $($offsetCounts['-1']) | other: $($offsetCounts['other'])"
Write-Host "Building JSON ($($dataRows.Count) unique rows)..."
$hJson = ($COLS | ForEach-Object { '"' + $_ + '"' }) -join ','

$sb = [System.Text.StringBuilder]::new(200MB)
$first = $true
foreach ($v2 in $dataRows) {
    if (!$first) { $sb.Append(',') | Out-Null }
    $first = $false
    $sb.Append('[') | Out-Null
    for ($ci = 0; $ci -lt $COLS.Count; $ci++) {
        if ($ci -gt 0) { $sb.Append(',') | Out-Null }
        $col = $COLS[$ci]; $v = $v2[$ci]
        if ($col -eq 'Eur') { $sb.Append([string]$v) | Out-Null }
        elseif ($col -eq 'Untouched Days') { if ($null -eq $v) { $sb.Append('null') | Out-Null } else { $sb.Append([string]$v) | Out-Null } }
        else { $sb.Append('"').Append([string]$v).Append('"') | Out-Null }
    }
    $sb.Append(']') | Out-Null
}

$dJson = $sb.ToString()
$js45 = 'window.PIPE_DATA_W45={h:[' + $hJson + '],d:[' + $dJson + ']};'
[System.IO.File]::WriteAllText("$outDir\data-pipe-w45.js", $js45, [System.Text.Encoding]::UTF8)
Write-Host "Written: data-pipe-w45.js ($([Math]::Round($js45.Length/1MB,1)) MB)"

$jsBase = 'window.PIPE_DATA={h:[' + $hJson + '],d:[' + $dJson + ']};'
[System.IO.File]::WriteAllText("$outDir\data-pipe.js", $jsBase, [System.Text.Encoding]::UTF8)
Write-Host "Written: data-pipe.js"
Write-Host "Done."
