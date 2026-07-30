$xlsxPath = "C:\Users\I572929\OneDrive - SAP SE\2026\Security Lookup.xlsx"
$outDir   = "C:\Users\I572929\campaign-calendar-site"
$zipPath  = "$outDir\__sec_tmp.zip"
$extPath  = "$outDir\__sec_tmp_ext"

Copy-Item $xlsxPath $zipPath -Force
if (Test-Path $extPath) { Remove-Item $extPath -Recurse -Force }
Expand-Archive -Path $zipPath -DestinationPath $extPath -Force

$ssXml = New-Object System.Xml.XmlDocument
$ssXml.Load("$extPath\xl\sharedStrings.xml")
$strings = $ssXml.sst.si | ForEach-Object {
    if ($null -ne $_.t) { $_.t }
    elseif ($_.r) { ($_.r | ForEach-Object { $_.t }) -join '' }
    else { '' }
}
$shXml = New-Object System.Xml.XmlDocument
$shXml.Load("$extPath\xl\worksheets\sheet1.xml")
$rows = $shXml.worksheet.sheetData.row

function Get-Val($cell) {
    $v = $cell.v; if ($null -eq $v) { return '' }
    if ($cell.t -eq 's') { return $strings[[int]$v] }
    return $v
}

# Build header index
$headers = @{}
$col = 0
foreach ($cell in $rows[0].c) {
    $colLetter = $cell.r -replace '\d+', ''
    $headers[$colLetter] = Get-Val $cell
}
$colIdx = @{}
foreach ($k in $headers.Keys) { $colIdx[$headers[$k]] = $k }

$entries = [System.Collections.Generic.List[string]]::new()

for ($i = 1; $i -lt $rows.Count; $i++) {
    $row = $rows[$i]
    $rowVals = @{}
    foreach ($cell in $row.c) {
        $cl = $cell.r -replace '\d+', ''
        $rowVals[$cl] = Get-Val $cell
    }
    function GetByHdr($h) { $l = $colIdx[$h]; if ($null -eq $l) { return '' }; $v = $rowVals[$l]; if ($null -eq $v) { return '' }; return $v }

    $email = (GetByHdr 'Email').Trim().ToLower()
    if (-not $email) { continue }
    $rl2 = (GetByHdr 'Region Level 2').Trim()
    $rl3 = (GetByHdr 'Region Level 3').Trim()

    $entry = '{"email":"' + $email + '","rl2":"' + $rl2.Replace('"','\"') + '","rl3":"' + $rl3.Replace('"','\"') + '"}'
    $entries.Add($entry)
}

$js = "window.SECURITY_LOOKUP=[" + ($entries -join ',') + "];"
[System.IO.File]::WriteAllText("$outDir\data-security.js", $js, (New-Object System.Text.UTF8Encoding $false))
Write-Host "Written $($entries.Count) entries to data-security.js"

Remove-Item $zipPath -Force
Remove-Item $extPath -Recurse -Force
