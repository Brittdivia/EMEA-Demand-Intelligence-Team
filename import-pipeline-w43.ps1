# import-pipeline-w43.ps1
# Reads Week 43.csv and generates data-pipe-w43.js + data-pipe.js

$csvPath = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Pipeline\Week 43.csv"
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

Write-Host "Reading CSV..."
$allRows = Import-Csv -Path $csvPath -Encoding UTF8
Write-Host "Total rows: $($allRows.Count)"

# Output columns (same schema as w42)
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

foreach ($row in $allRows) {
    $oppId = if ($row.PSObject.Properties['Opportunity ID']) { [string]$row.'Opportunity ID' } else { '' }
    $oppId = $oppId.Trim()
    if ([string]::IsNullOrWhiteSpace($oppId)) { continue }

    $subSol = if ($row.PSObject.Properties['Sub-Solution Area (L2)']) { [string]$row.'Sub-Solution Area (L2)' } else { '' }
    $subSol = $subSol.Trim()
    $key = $oppId + '|' + $subSol
    if (!$seen.Add($key)) { continue }

    $gbId = if ($row.PSObject.Properties['GB Identifier']) { [string]$row.'GB Identifier' } else { '' }
    $mmId = $gbId.Trim()

    $eur = 0.0
    $eurRaw = if ($row.PSObject.Properties['Eur']) { [string]$row.'Eur' } else { '' }
    if ($eurRaw) { [double]::TryParse($eurRaw.Trim(), [ref]$eur) | Out-Null }

    $untouched = $null
    $udRaw = if ($row.PSObject.Properties['Untouched Days']) { [string]$row.'Untouched Days' } else { '' }
    if ($udRaw.Trim()) {
        $ud = 0.0
        if ([double]::TryParse($udRaw.Trim(), [ref]$ud)) { $untouched = [int]$ud }
    }

    function G($r, $n) {
        if ($r.PSObject.Properties[$n]) { return ([string]$r.$n).Trim() } else { return '' }
    }

    $vals = [object[]]@(
        (EscapeJson (G $row 'Region Lvl 2'))
        (EscapeJson (G $row 'Region Lvl 3'))
        (EscapeJson (G $row 'DRM Category'))
        (EscapeJson (G $row 'Account Name'))
        (EscapeJson $oppId)
        (EscapeJson (G $row 'Closing Qtr'))
        (EscapeJson (G $row 'IAC (Engagement Model)'))
        (EscapeJson (G $row 'Opp Campaign ID'))
        (EscapeJson (G $row 'Solution Area (L1)'))
        (EscapeJson $subSol)
        (EscapeJson (G $row 'Create Quarter'))
        (EscapeJson (G $row 'IC_ACCOUNT_SRG_TEXT'))
        (EscapeJson $mmId)
        (EscapeJson (ParseDate (G $row 'SDE Handover Date')))
        (EscapeJson (G $row 'SDE Initial Engagement Name'))
        (EscapeJson (G $row 'SDE Primary Engagement Name'))
        (EscapeJson (G $row 'SDE Secondary Engagement Name'))
        (EscapeJson (G $row 'SDE Handover Role'))
        (EscapeJson (G $row 'Opp Description'))
        (EscapeJson (G $row 'TCP Pipeline Source Desc'))
        (EscapeJson (G $row 'SAP Mastercode'))
        (EscapeJson (ParseDate (G $row 'Create Date')))
        (EscapeJson (G $row 'SDE Engagement Type'))
        $untouched
        (EscapeJson (G $row 'Channel Partner'))
        (EscapeJson (G $row 'RBC'))
        $eur
        (EscapeJson (G $row 'MM Pipe Source'))
        (EscapeJson (G $row 'Local Acct Owner'))
    )
    $dataRows.Add($vals)
}

Write-Host "Building JSON ($($dataRows.Count) unique rows)..."
$hJson = ($COLS | ForEach-Object { '"' + $_ + '"' }) -join ','

$sb = [System.Text.StringBuilder]::new(200MB)
$first = $true
foreach ($vals in $dataRows) {
    if (!$first) { $sb.Append(',') | Out-Null }
    $first = $false
    $sb.Append('[') | Out-Null
    for ($ci = 0; $ci -lt $COLS.Count; $ci++) {
        if ($ci -gt 0) { $sb.Append(',') | Out-Null }
        $col = $COLS[$ci]
        $v   = $vals[$ci]
        if ($col -eq 'Eur') {
            $sb.Append([string]$v) | Out-Null
        } elseif ($col -eq 'Untouched Days') {
            if ($null -eq $v) { $sb.Append('null') | Out-Null } else { $sb.Append([string]$v) | Out-Null }
        } else {
            $sb.Append('"').Append([string]$v).Append('"') | Out-Null
        }
    }
    $sb.Append(']') | Out-Null
}

$dJson = $sb.ToString()

$js43 = 'window.PIPE_DATA_W43={h:[' + $hJson + '],d:[' + $dJson + ']};'
[System.IO.File]::WriteAllText("$outDir\data-pipe-w43.js", $js43, [System.Text.Encoding]::UTF8)
Write-Host "Written: data-pipe-w43.js ($([Math]::Round($js43.Length/1MB,1)) MB)"

$jsBase = 'window.PIPE_DATA={h:[' + $hJson + '],d:[' + $dJson + ']};'
[System.IO.File]::WriteAllText("$outDir\data-pipe.js", $jsBase, [System.Text.Encoding]::UTF8)
Write-Host "Written: data-pipe.js"

Write-Host "Done."
