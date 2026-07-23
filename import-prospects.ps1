# import-prospects.ps1
# Rebuilds data-prospects.js from Prospects_Export.xlsx
# Strips unused fields: aid, added, owner, ti

$xlsxPath = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Outreach\Prospects_Export.xlsx"
$outDir   = "C:\Users\I572929\campaign-calendar-site"
$zipPath  = "$outDir\__prosp_temp.zip"
$extPath  = "$outDir\__prosp_extracted"

Write-Host "Copying file..."
Copy-Item $xlsxPath $zipPath -Force

Write-Host "Extracting..."
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
function EscJS($s) {
    if ($null -eq $s) { return '' }
    return [string]$s -replace '\\','\\' -replace '"','\"' -replace "`r`n",' ' -replace "`n",' ' -replace "`r",' '
}

# Build header index
$headers = @{}
foreach ($cell in $rows[0].c) {
    $col = $cell.r -replace '\d+', ''
    $headers[$col] = Get-CellValue $cell
}
$colIdx = @{}
foreach ($k in $headers.Keys) { $colIdx[$headers[$k]] = $k }
Write-Host "Headers found: $($colIdx.Count)"

# Only keep fields actually used in index.html
$fieldMap = @{
    'ID'                 = 'id'
    'External ID'        = 'eid'
    'Company'            = 'co'
    'Touched At'         = 'touched'
    'Stage Changed At'   = 'sc'
    'Created At'         = 'created'
    'Tags'               = 'tg'
    'Stage Name'         = 'stage'
    'Email'              = 'email'
    'Assigned Users'     = 'au'
    'Persona Name'       = 'pn'
    'Active Sequences'   = 'as'
    'Finished Sequences' = 'fs'
}

Write-Host "Processing $($rows.Count - 1) rows..."
$entries = [System.Collections.Generic.List[string]]::new()
$seenIds = [System.Collections.Generic.HashSet[string]]::new()

for ($i = 1; $i -lt $rows.Count; $i++) {
    $rowVals = @{}
    foreach ($cell in $rows[$i].c) {
        $col = $cell.r -replace '\d+', ''
        $rowVals[$col] = Get-CellValue $cell
    }

    $id = ''
    foreach ($h in $fieldMap.Keys) { if ($colIdx[$h]) { $v = $rowVals[$colIdx[$h]]; if ($h -eq 'ID' -and $v) { $id = $v } } }
    if (-not $id -or -not $seenIds.Add($id)) { continue }

    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($h in @('ID','External ID','Company','Touched At','Stage Changed At','Created At','Tags','Stage Name','Email','Assigned Users','Persona Name','Active Sequences','Finished Sequences')) {
        $f = $fieldMap[$h]
        $v = if ($colIdx[$h]) { EscJS $rowVals[$colIdx[$h]] } else { '' }
        $parts.Add("""$f"":""$v""")
    }
    $entries.Add('{' + ($parts -join ',') + '}')

    if ($i % 50000 -eq 0) { Write-Host "  Processed $i / $($rows.Count-1)" }
}

Write-Host "Writing $($entries.Count) prospects..."
$js = "window.PROSPECT_DATA=[" + ($entries -join ',') + "];"
[System.IO.File]::WriteAllText("$outDir\data-prospects.js", $js, [System.Text.Encoding]::UTF8)
Write-Host "Done: $([Math]::Round((Get-Item "$outDir\data-prospects.js").Length/1MB, 1))MB"

Remove-Item $zipPath -Force
Remove-Item $extPath -Recurse -Force
