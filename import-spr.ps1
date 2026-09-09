# import-spr.ps1
# Reads "SPR Add on.xlsx" and generates data-spr.js -> window.SPR_DATA

$xlsxPath = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Targets\SPR Add on.xlsx"
$outDir   = "C:\Users\I572929\campaign-calendar-site"
$zipPath  = "$outDir\__spr_temp.zip"
$extPath  = "$outDir\__spr_extracted"

Write-Host "Copying and extracting..."
Copy-Item $xlsxPath $zipPath -Force
if (Test-Path $extPath) { Remove-Item $extPath -Recurse -Force }
Expand-Archive -Path $zipPath -DestinationPath $extPath -Force

$ssXml = New-Object System.Xml.XmlDocument
$ssXml.Load("$extPath\xl\sharedStrings.xml")
$strings = $ssXml.sst.si | ForEach-Object { $_.InnerText }

function GCV($cell) {
    $v = $cell.v; if ($null -eq $v) { return '' }
    if ($cell.t -eq 's') { return $strings[[int]$v] }
    return $v
}
function EscJS($s) {
    return [string]$s -replace '\\','\\' -replace '"','\"' -replace "`r`n",' ' -replace "`n",' ' -replace "`r",' '
}

# Quarter flag conversion: Q126->2026-Q1, Q226->2026-Q2, Q326->2026-Q3, Q426->2026-Q4
function ConvQtr($q) {
    $q = [string]$q
    if ($q -match '^Q([1-4])26$') { return "2026-Q$($matches[1])" }
    return $q
}

# Normalise L2 region names to match existing budget region names
function NormMu($s) {
    switch ([string]$s.Trim().ToLower()) {
        "benelux" { return "Benelux" }
        "belux"   { return "Benelux" }  # fallback if L2 shows BeLux
        "france"  { return "France" }
        "italy"   { return "Italy" }
        "nordic"  { return "Nordic" }
        "southern europe" { return "Southern Europe" }
        "uki"     { return "UKI" }
        "mea north" { return "MEA North" }
        "mea south" { return "MEA South" }
        default   { return $s.Trim() }
    }
}

$sh = New-Object System.Xml.XmlDocument
$sh.Load("$extPath\xl\worksheets\sheet1.xml")
$rows = $sh.worksheet.sheetData.row
Write-Host "Total rows: $($rows.Count - 1)"

# Use fixed column letters (verified from header row):
# A=IBP Region, B=IBP Market Unit (L2), C=IBP Market Unit-1 (L3),
# I=Quarter Flag, K=Cloud N&U Quota, L=Team, M=IAC

function GetCell($row, $col) {
    foreach ($cell in $row.c) {
        $cl = $cell.r -replace '\d+', ''
        if ($cl -eq $col) { return GCV $cell }
    }
    return ''
}

# Handle rich-text inline string cells (Team column)
function GetCellRich($row, $col) {
    foreach ($cell in $row.c) {
        $cl = $cell.r -replace '\d+', ''
        if ($cl -eq $col) {
            # shared string
            if ($cell.t -eq 's') { return $strings[[int]$cell.v] }
            # inline string (is element)
            if ($cell.is) {
                $t = $cell.is.t
                if ($t -is [string]) { return $t }
                # rich text runs
                if ($cell.is.r) { return ($cell.is.r | ForEach-Object { $_.t }) -join '' }
            }
            if ($null -ne $cell.v) { return $cell.v }
            return ''
        }
    }
    return ''
}

$entries = [System.Collections.Generic.List[string]]::new()

for ($i = 1; $i -lt $rows.Count; $i++) {
    $row = $rows[$i]

    $mu   = NormMu (GetCell $row 'B')
    $mu1  = EscJS  (GetCell $row 'C')
    if ([string]::IsNullOrWhiteSpace($mu1)) { $mu1 = $mu }  # fall back to L2 when L3 is blank
    $team = EscJS  (GetCell $row 'L')
    $iac  = EscJS  (GetCell $row 'M')
    $qtr  = EscJS  (ConvQtr (GetCell $row 'I'))

    $quotaRaw = GetCell $row 'K'
    $quota = 0.0
    if (-not [string]::IsNullOrWhiteSpace($quotaRaw)) {
        [double]::TryParse($quotaRaw, [Globalization.NumberStyles]::Any,
            [Globalization.CultureInfo]::InvariantCulture, [ref]$quota) | Out-Null
    }

    if ([string]::IsNullOrWhiteSpace($mu)) { continue }

    $entries.Add("{`"mu`":`"$mu`",`"mu1`":`"$(EscJS $mu1)`",`"team`":`"$team`",`"iac`":`"$iac`",`"qtr`":`"$qtr`",`"quota`":$([Math]::Round($quota,2))}")
}

Write-Host "Writing $($entries.Count) rows..."
$js = "window.SPR_DATA=[" + ($entries -join ',') + "];"
[System.IO.File]::WriteAllText("$outDir\data-spr.js", $js, [System.Text.Encoding]::UTF8)

Remove-Item $zipPath -Force
Remove-Item $extPath -Recurse -Force

Write-Host "Done: $([Math]::Round((Get-Item "$outDir\data-spr.js").Length/1KB))KB"
