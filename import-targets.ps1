# import-targets.ps1
# Reads EMEA 2026 TCP Targets.xlsx Sheet1 (Region targets) and writes data-targets.js

$xlsxPath = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Targets\EMEA 2026 TCP Targets.xlsx"
$outDir   = "C:\Users\I572929\campaign-calendar-site"
$zipPath  = "$outDir\__tgt_temp.zip"
$extPath  = "$outDir\__tgt_extracted"

Write-Host "Copying and extracting..."
Copy-Item $xlsxPath $zipPath -Force
if (Test-Path $extPath) { Remove-Item $extPath -Recurse -Force }
Expand-Archive -Path $zipPath -DestinationPath $extPath -Force

$ssXml = New-Object System.Xml.XmlDocument
$ssXml.Load("$extPath\xl\sharedStrings.xml")
$strings = $ssXml.sst.si | ForEach-Object {
    if ($null -ne $_.t) { $_.t } elseif ($_.r) { ($_.r | ForEach-Object { $_.t }) -join '' } else { '' }
}

function Get-CellValue($cell) {
    $v = $cell.v; if ($null -eq $v) { return '' }
    if ($cell.t -eq 's') { return $strings[[int]$v] }
    return $v
}

# Sheet1 = Region summary
$sh1 = New-Object System.Xml.XmlDocument
$sh1.Load("$extPath\xl\worksheets\sheet1.xml")
$rows1 = $sh1.worksheet.sheetData.row

$entries = [System.Collections.Generic.List[string]]::new()
for ($i = 1; $i -lt $rows1.Count; $i++) {
    $vals = @{}
    foreach ($cell in $rows1[$i].c) {
        $col = [int]($cell.r -replace '[A-Z]','') ; $colL = $cell.r -replace '\d+',''
        $vals[$colL] = Get-CellValue $cell
    }
    $region = $vals['A']; $target = $vals['B']
    if (-not $region -or $region -eq 'Row Labels' -or $region -eq 'Grand Total') { continue }
    $eur = [double]0; [double]::TryParse($target, [ref]$eur) | Out-Null
    $region = $region.Replace('\','\\').Replace('"','\"')
    $entries.Add("{""region"":""$region"",""target"":$([Math]::Round($eur,2))}")
}

Write-Host "Sheet1 targets: $($entries.Count) regions"

# Sheet2 = Detailed breakdown (Region/MU/SA/SSA/Quarter/Source)
$sh2 = New-Object System.Xml.XmlDocument
$sh2.Load("$extPath\xl\worksheets\sheet2.xml")
$rows2 = $sh2.worksheet.sheetData.row

# Build header index from sheet2
$hdr2 = @{}
foreach ($cell in $rows2[0].c) {
    $col = $cell.r -replace '\d+',''
    $hdr2[$col] = Get-CellValue $cell
}
$colIdx2 = @{}
foreach ($k in $hdr2.Keys) { $colIdx2[$hdr2[$k]] = $k }

$entries2 = [System.Collections.Generic.List[string]]::new()
for ($i = 1; $i -lt $rows2.Count; $i++) {
    $vals = @{}
    foreach ($cell in $rows2[$i].c) {
        $col = $cell.r -replace '\d+',''
        $vals[$col] = Get-CellValue $cell
    }
    function GetVal($h) { $l=$colIdx2[$h]; if (-not $l) { return '' }; $v=$vals[$l]; if ($null -eq $v) { return '' }; return [string]$v }
    $regVal  = (GetVal 'Region').Replace('"','\"')
    $muVal   = (GetVal 'MU').Replace('"','\"')
    $saVal   = (GetVal 'SA').Replace('"','\"')
    $ssaVal  = (GetVal 'SSA').Replace('"','\"')
    $qtrVal  = GetVal 'Creation Quarter'
    $srcVal  = (GetVal 'TCP Source').Replace('"','\"')
    $tgtRaw  = GetVal '2026 TCP Target kEUR'
    $tgtVal  = [double]0; [double]::TryParse($tgtRaw, [ref]$tgtVal) | Out-Null
    if (-not $regVal) { continue }
    $entries2.Add("{""region"":""$regVal"",""mu"":""$muVal"",""sa"":""$saVal"",""ssa"":""$ssaVal"",""qtr"":""$qtrVal"",""src"":""$srcVal"",""target"":$([Math]::Round($tgtVal,2))}")
}
Write-Host "Sheet2 detail: $($entries2.Count) rows"

$js = "window.TCP_TARGETS_REGION=[" + ($entries -join ',') + "];`n"
$js += "window.TCP_TARGETS_DETAIL=[" + ($entries2 -join ',') + "];"
[System.IO.File]::WriteAllText("$outDir\data-targets.js", $js, [System.Text.Encoding]::UTF8)
Write-Host "Written data-targets.js ($([Math]::Round((Get-Item "$outDir\data-targets.js").Length/1KB))KB)"

Remove-Item $zipPath -Force
Remove-Item $extPath -Recurse -Force
Write-Host "Done."
