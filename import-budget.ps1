# import-budget.ps1
# Reads Budget 24 25 26.xlsx Sheet1 and writes data-budget.js

$xlsxPath = "C:\Users\I572929\OneDrive - SAP SE\2026\DLs for Update\Budget\Budget 24 25 26.xlsx"
$outDir   = "C:\Users\I572929\campaign-calendar-site"
$zipPath  = "$outDir\__bud_temp.zip"
$extPath  = "$outDir\__bud_extracted"

Write-Host "Copying and extracting..."
Copy-Item $xlsxPath $zipPath -Force
if (Test-Path $extPath) { Remove-Item $extPath -Recurse -Force }
Expand-Archive -Path $zipPath -DestinationPath $extPath -Force

$ssXml = New-Object System.Xml.XmlDocument
$ssXml.Load("$extPath\xl\sharedStrings.xml")
$strings = $ssXml.sst.si | ForEach-Object {
    if ($null -ne $_.t) { $_.t } elseif ($_.r) { ($_.r | ForEach-Object { $_.t }) -join '' } else { '' }
}

function GCV($cell) {
    $v = $cell.v; if ($null -eq $v) { return '' }
    if ($cell.t -eq 's') { return $strings[[int]$v] }
    return $v
}
function EscJS($s) {
    return [string]$s -replace '\\','\\' -replace '"','\"' -replace "`r`n",' ' -replace "`n",' ' -replace "`r",' '
}

$sh = New-Object System.Xml.XmlDocument
$sh.Load("$extPath\xl\worksheets\sheet1.xml")
$rows = $sh.worksheet.sheetData.row
Write-Host "Total rows: $($rows.Count - 1)"

$entries = [System.Collections.Generic.List[string]]::new()

for ($i = 1; $i -lt $rows.Count; $i++) {
    $vals = @{}
    foreach ($cell in $rows[$i].c) {
        $col = $cell.r -replace '\d+', ''
        $vals[$col] = GCV $cell
    }
    $reg  = EscJS $vals['A']
    $pc   = EscJS $vals['B']
    $qtr  = EscJS $vals['C']
    $iac  = EscJS $vals['D']
    $sol2 = EscJS $vals['E']
    $sol3 = EscJS $vals['F']
    $eurRaw = $vals['H']
    $eur  = [double]0; [double]::TryParse($eurRaw, [ref]$eur) | Out-Null

    if (-not $reg -or $qtr -eq 'Result') { continue }

    $entries.Add("{""reg"":""$reg"",""pc"":""$pc"",""qtr"":""$qtr"",""iac"":""$iac"",""sol2"":""$sol2"",""sol3"":""$sol3"",""eur"":$([Math]::Round($eur,3))}")

    if ($i % 5000 -eq 0) { Write-Host "  $i / $($rows.Count-1)" }
}

Write-Host "Writing $($entries.Count) rows..."
$js = "window.BUDGET_DATA=[" + ($entries -join ',') + "];"
[System.IO.File]::WriteAllText("$outDir\data-budget.js", $js, [System.Text.Encoding]::UTF8)
Write-Host "Done: $([Math]::Round((Get-Item "$outDir\data-budget.js").Length/1KB))KB"

Remove-Item $zipPath -Force
Remove-Item $extPath -Recurse -Force
