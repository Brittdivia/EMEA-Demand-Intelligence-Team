# import-pipeline.ps1
# Reads Week 22 DL.xlsx and generates data-pipe.js -> window.PIPE_DATA

$xlsxPath = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Pipeline\Week 22 DL.xlsx"
$outDir   = "C:\Users\I572929\campaign-calendar-site"
$zipPath  = "$outDir\__pipe_temp.zip"
$extPath  = "$outDir\__pipe_extracted"

Write-Host "Copying file..."
Copy-Item $xlsxPath $zipPath -Force

Write-Host "Extracting..."
if (Test-Path $extPath) { Remove-Item $extPath -Recurse -Force }
Expand-Archive -Path $zipPath -DestinationPath $extPath -Force

$ssXml = New-Object System.Xml.XmlDocument
$ssXml.Load("$extPath\xl\sharedStrings.xml")
$strings = $ssXml.sst.si | ForEach-Object {
    if ($null -ne $_.t) { $_.t }
    elseif ($_.r)       { ($_.r | ForEach-Object { $_.t }) -join '' }
    else                { '' }
}

$sheetXml = New-Object System.Xml.XmlDocument
$sheetXml.Load("$extPath\xl\worksheets\sheet1.xml")
$rows = $sheetXml.worksheet.sheetData.row

function Get-CellValue($cell) {
    $v = $cell.v
    if ($null -eq $v) { return '' }
    if ($cell.t -eq 's') { return $strings[[int]$v] }
    return $v
}

function EscapeJson($val) {
    if ($null -eq $val) { return '' }
    $s = [string]$val
    $s = $s.Replace('\','\\').Replace('"','\"').Replace("`r`n",' ').Replace("`n",' ').Replace("`r",' ').Replace("`t",' ')
    return $s
}

function ConvertTo-DateStr($serial) {
    if ([string]::IsNullOrWhiteSpace($serial)) { return '' }
    try {
        $d = [double]$serial
        if ($d -gt 1 -and $d -lt 200000) { return ([DateTime]"1899-12-30").AddDays($d).ToString("yyyy-MM-dd") }
    } catch {}
    try { return ([DateTime]::Parse($serial, [System.Globalization.CultureInfo]::InvariantCulture)).ToString("yyyy-MM-dd") } catch {}
    return $serial
}

# Build header index
$headers = @{}
foreach ($cell in $rows[0].c) {
    $col = $cell.r -replace '\d+', ''
    $headers[$col] = Get-CellValue $cell
}
$colIdx = @{}
foreach ($k in $headers.Keys) { $colIdx[$headers[$k]] = $k }

Write-Host "Processing $($rows.Count - 1) data rows..."

$pipeRows = [System.Collections.Generic.List[string]]::new()

for ($i = 1; $i -lt $rows.Count; $i++) {
    $row = $rows[$i]

    $rowVals = @{}
    foreach ($cell in $row.c) {
        $col = $cell.r -replace '\d+', ''
        $rowVals[$col] = Get-CellValue $cell
    }

    function GetByHeader($hdrName) {
        $letter = $colIdx[$hdrName]
        if ($null -eq $letter) { return '' }
        $v = $rowVals[$letter]
        if ($null -eq $v) { return '' }
        return $v
    }

    $oppId = GetByHeader 'Opportunity ID'
    if ([string]::IsNullOrWhiteSpace($oppId)) { continue }

    $eur = 0.0
    $eurRaw = GetByHeader 'Eur'
    if (-not [string]::IsNullOrWhiteSpace($eurRaw)) {
        [double]::TryParse($eurRaw, [ref]$eur) | Out-Null
    }

    $fields  = '"Opp Campaign ID":"'         + (EscapeJson (GetByHeader 'Opp Campaign ID'))           + '"'
    $fields += ',"Opportunity ID":"'         + (EscapeJson $oppId)                                    + '"'
    $fields += ',"Eur":'                     + $eur
    $fields += ',"DRM Category":"'           + (EscapeJson (GetByHeader 'DRM Category'))              + '"'
    $fields += ',"SDE Handover Date":"'      + (EscapeJson (ConvertTo-DateStr (GetByHeader 'SDE Handover Date'))) + '"'
    $fields += ',"Create Date":"'            + (EscapeJson (ConvertTo-DateStr (GetByHeader 'Create Date')))       + '"'
    $fields += ',"Create Quarter":"'         + (EscapeJson (GetByHeader 'Create Quarter'))            + '"'
    $fields += ',"Closing Qtr":"'            + (EscapeJson (GetByHeader 'Closing Qtr'))               + '"'
    $fields += ',"TCP Pipeline Source Desc":"' + (EscapeJson (GetByHeader 'TCP Pipeline Source Desc')) + '"'
    $gbId = GetByHeader 'GB Identifier'
    if ([string]::IsNullOrWhiteSpace($gbId)) { $gbId = GetByHeader 'MM Identifier' }
    $fields += ',"MM Identifier":"'          + (EscapeJson $gbId)                                     + '"'
    $fields += ',"RBC":"'                    + (EscapeJson (GetByHeader 'RBC'))                       + '"'
    $fields += ',"Account Name":"'           + (EscapeJson (GetByHeader 'Account Name'))              + '"'
    $fields += ',"Region Lvl 2":"'           + (EscapeJson (GetByHeader 'Region Lvl 2'))              + '"'
    $fields += ',"Region Lvl 3":"'           + (EscapeJson (GetByHeader 'Region Lvl 3'))              + '"'
    $fields += ',"SAP Mastercode":"'         + (EscapeJson (GetByHeader 'SAP Mastercode'))            + '"'
    $fields += ',"Solution Area (L1)":"'     + (EscapeJson (GetByHeader 'Solution Area (L1)'))        + '"'
    $fields += ',"Sub-Solution Area (L2)":"' + (EscapeJson (GetByHeader 'Sub-Solution Area (L2)'))    + '"'
    $fields += ',"IAC (Engagement Model)":"' + (EscapeJson (GetByHeader 'IAC (Engagement Model)'))    + '"'
    $fields += ',"SDE Engagement Type":"'    + (EscapeJson (GetByHeader 'SDE Engagement Type'))       + '"'
    $fields += ',"SDE Initial Engagement Role":"' + (EscapeJson (GetByHeader 'SDE Initial Engagement Role')) + '"'
    $fields += ',"SDE Territory Owner":"'    + (EscapeJson (GetByHeader 'SDE Current Territory Owner Name')) + '"'
    $fields += ',"Opp Description":"'        + (EscapeJson (GetByHeader 'Opp Description'))           + '"'

    $pipeRows.Add('{' + $fields + '}')
}

Write-Host "Writing data-pipe.js ($($pipeRows.Count) rows)..."
$js = 'window.PIPE_DATA=[' + ($pipeRows -join ',') + '];'
[System.IO.File]::WriteAllText("$outDir\data-pipe.js", $js, [System.Text.Encoding]::UTF8)

Remove-Item $zipPath -Force
Remove-Item $extPath -Recurse -Force

Write-Host "Done."
