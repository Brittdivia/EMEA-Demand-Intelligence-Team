# add-summary-sheet.ps1
# Adds a "Country & Source Summary" sheet to Prospects_Export.xlsx
# Uses zip/XML approach - no Excel COM required

$xlsxPath = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Outreach\Prospects_Export.xlsx"
$outDir   = "C:\Users\I572929\campaign-calendar-site"
$zipPath  = "$outDir\__pros_sum_temp.zip"
$extPath  = "$outDir\__pros_sum_extracted"

Write-Host "Copying and extracting..."
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

$sheetXml = New-Object System.Xml.XmlDocument
$sheetXml.Load("$extPath\xl\worksheets\sheet1.xml")
$rows = $sheetXml.worksheet.sheetData.row

function Get-CellValue($cell) {
    $v = $cell.v; if ($null -eq $v) { return '' }
    if ($cell.t -eq 's') { return $strings[[int]$v] }
    return $v
}

function EscapeXml($s) { return [System.Security.SecurityElement]::Escape([string]$s) }

# Build header index
$headers = @{}
foreach ($cell in $rows[0].c) {
    $col = $cell.r -replace '\d+', ''
    $headers[$col] = Get-CellValue $cell
}
$colIdx = @{}
foreach ($k in $headers.Keys) { $colIdx[$headers[$k]] = $k }

function GetByHeader($rowVals, $hdrName) {
    $letter = $colIdx[$hdrName]; if ($null -eq $letter) { return '' }
    $v = $rowVals[$letter]; if ($null -eq $v) { return '' }; return [string]$v
}

Write-Host "Processing $($rows.Count - 1) rows..."
$summary = @{}
$seenPid = [System.Collections.Generic.HashSet[string]]::new()
$allAccounts = [System.Collections.Generic.HashSet[string]]::new()

for ($i = 1; $i -lt $rows.Count; $i++) {
    $row = $rows[$i]
    $rowVals = @{}
    foreach ($cell in $row.c) { $col = $cell.r -replace '\d+', ''; $rowVals[$col] = Get-CellValue $cell }

    $prospId = (GetByHeader $rowVals 'ID').Trim()
    if (-not $prospId -or -not $seenPid.Add($prospId)) { continue }

    $co      = (GetByHeader $rowVals 'Company').Trim()
    $country = (GetByHeader $rowVals 'Country').Trim(); if (-not $country) { $country = "(No Country)" }
    $src     = (GetByHeader $rowVals 'Source').Trim();  if (-not $src)     { $src = "(No Source)" }
    $ct      = (GetByHeader $rowVals 'Company Type').Trim()

    if (-not $summary.ContainsKey($country)) {
        $summary[$country] = @{ prospects = [System.Collections.Generic.HashSet[string]]::new(); accounts = [System.Collections.Generic.HashSet[string]]::new(); sources = @{} }
    }
    [void]$summary[$country].prospects.Add($prospId)
    if ($co) { [void]$summary[$country].accounts.Add($co); [void]$allAccounts.Add($co) }
    if (-not $summary[$country].sources.ContainsKey($src)) { $summary[$country].sources[$src] = 0 }
    $summary[$country].sources[$src]++

    if ($i % 50000 -eq 0) { Write-Host "  Processed $i / $($rows.Count-1)" }
}

Write-Host "Found $($summary.Count) countries, $($seenPid.Count) unique prospects, $($allAccounts.Count) unique accounts"

# Build new worksheet XML
$sortedEntries = $summary.GetEnumerator() | Sort-Object { $_.Value.prospects.Count } -Descending

# Shared strings for the new sheet - we'll write inline strings (type="inlineStr")
$wsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n"
$wsXml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + "`n"
$wsXml += '<sheetData>' + "`n"

# Header row
$wsXml += '<row r="1"><c r="A1" t="inlineStr"><is><t>Country</t></is></c><c r="B1" t="inlineStr"><is><t>Unique Prospects</t></is></c><c r="C1" t="inlineStr"><is><t>Unique Accounts</t></is></c><c r="D1" t="inlineStr"><is><t>By Source (detail)</t></is></c></row>' + "`n"

$rowNum = 2
foreach ($entry in $sortedEntries) {
    $srcDetail = ($entry.Value.sources.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { "$($_.Key): $($_.Value)" }) -join "; "
    $countryEsc = EscapeXml $entry.Key
    $srcEsc = EscapeXml $srcDetail
    $wsXml += "<row r=`"$rowNum`"><c r=`"A$rowNum`" t=`"inlineStr`"><is><t>$countryEsc</t></is></c><c r=`"B$rowNum`"><v>$($entry.Value.prospects.Count)</v></c><c r=`"C$rowNum`"><v>$($entry.Value.accounts.Count)</v></c><c r=`"D$rowNum`" t=`"inlineStr`"><is><t>$srcEsc</t></is></c></row>`n"
    $rowNum++
}
# Totals row
$wsXml += "<row r=`"$rowNum`"><c r=`"A$rowNum`" t=`"inlineStr`"><is><t>TOTAL ($($summary.Count) countries)</t></is></c><c r=`"B$rowNum`"><v>$($seenPid.Count)</v></c><c r=`"C$rowNum`"><v>$($allAccounts.Count)</v></c><c r=`"D$rowNum`" t=`"inlineStr`"><is><t></t></is></c></row>`n"

$wsXml += '</sheetData></worksheet>'

# Find next sheet number
$existingSheets = Get-ChildItem "$extPath\xl\worksheets" -Filter "sheet*.xml"
$nextNum = ($existingSheets.Count + 1)
$newSheetFile = "$extPath\xl\worksheets\sheet$nextNum.xml"
[System.IO.File]::WriteAllText($newSheetFile, $wsXml, [System.Text.Encoding]::UTF8)

# Update workbook.xml to add the new sheet
$wbXml = [System.IO.File]::ReadAllText("$extPath\xl\workbook.xml", [System.Text.Encoding]::UTF8)
$sheetId = $nextNum + 10
$newSheetEntry = "<sheet name=`"Country Source Summary`" sheetId=`"$sheetId`" r:id=`"rId$nextNum`"/>"
$wbXml = $wbXml -replace '</sheets>', "$newSheetEntry</sheets>"
[System.IO.File]::WriteAllText("$extPath\xl\workbook.xml", $wbXml, [System.Text.Encoding]::UTF8)

# Update workbook.xml.rels
$relsPath = "$extPath\xl\_rels\workbook.xml.rels"
$relsXml = [System.IO.File]::ReadAllText($relsPath, [System.Text.Encoding]::UTF8)
$newRel = "<Relationship Id=`"rId$nextNum`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$nextNum.xml`"/>"
$relsXml = $relsXml -replace '</Relationships>', "$newRel</Relationships>"
[System.IO.File]::WriteAllText($relsPath, $relsXml, [System.Text.Encoding]::UTF8)

# Update [Content_Types].xml
$ctPath = "$extPath\[Content_Types].xml"
$ctXml = [System.IO.File]::ReadAllText($ctPath, [System.Text.Encoding]::UTF8)
$newCt = "<Override PartName=`"/xl/worksheets/sheet$nextNum.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
$ctXml = $ctXml -replace '</Types>', "$newCt</Types>"
[System.IO.File]::WriteAllText($ctPath, $ctXml, [System.Text.Encoding]::UTF8)

# Repack
Write-Host "Repacking..."
$outZip = $zipPath -replace '\.zip$', '_out.zip'
if (Test-Path $outZip) { Remove-Item $outZip -Force }
Compress-Archive -Path "$extPath\*" -DestinationPath $outZip -Force
Copy-Item $outZip $xlsxPath -Force

Remove-Item $zipPath -Force
Remove-Item $outZip -Force
Remove-Item $extPath -Recurse -Force
Write-Host "Done. Sheet 'Country Source Summary' added to Prospects_Export.xlsx"
