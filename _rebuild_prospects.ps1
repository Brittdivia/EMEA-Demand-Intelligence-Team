# Rebuild data-prospects.js from Prospects_Export.xlsx
# Run whenever a new Prospects_Export is downloaded from Outreach

$xlFile  = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Outreach\Prospects_Export.xlsx"
$outFile = "C:\Users\I572929\campaign-calendar-site\data-prospects.js"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($xlFile, 0, $true)
$ws = $wb.Sheets.Item(1)

# Map headers
$headers = @{}
for ($col = 1; $col -le 120; $col++) {
    $val = $ws.Cells.Item(1, $col).Value2
    if ($val) { $headers["$val"] = $col }
}
Write-Host "Columns found: $($headers.Count)"

$colId       = $headers["Id"]
$colCo       = $headers["Company"]
$colTouched  = $headers["Touched At"]
$colSc       = $headers["Stage Changed At"]
$colAdded    = $headers["Added At"]
$colCreated  = $headers["Created At"]
$colTg       = $headers["Tags"]
$colOwner    = $headers["Owner Name"]
$colStage    = $headers["Stage Name"]
$colEmail    = $headers["Email"]
$colAu       = $headers["Assigned Users"]
$colTitle    = $headers["Title"]
$colPersona  = if ($headers["Persona Name"]) { $headers["Persona Name"] } else { 0 }
$colActive   = $headers["Active Sequences"]
$colFinished = $headers["Finished Sequences"]

Write-Host "Key cols: Id=$colId Co=$colCo Stage=$colStage Persona=$colPersona Active=$colActive Finished=$colFinished"

function ToDate($val) {
    if (-not $val) { return "" }
    $s = "$val".Trim()
    if ($s -match '(\d{4}-\d{2}-\d{2})') { return $Matches[1] }
    try { return ([DateTime]::Parse($s)).ToString("yyyy-MM-dd") } catch { return "" }
}

function EscJS($s) {
    if ($null -eq $s) { return "" }
    $s = "$s".Trim()
    $s = $s -replace '\\', '\\'
    $s = $s -replace '"',  '\"'
    $s = $s -replace "`r", ''
    $s = $s -replace "`n", ' '
    $s = $s -replace "`t", ' '
    return $s
}

function GetVal($row, $col) {
    if ($col -le 0) { return "" }
    try { $v = $ws.Cells.Item($row, $col).Value2; if ($null -eq $v) { return "" }; return "$v" }
    catch { return "" }
}

$entries = [System.Collections.Generic.List[string]]::new()
$lastRow = $ws.UsedRange.Rows.Count
Write-Host "Total rows: $lastRow"

for ($r = 2; $r -le $lastRow; $r++) {
    $id = GetVal $r $colId; if (-not $id) { continue }
    $co      = EscJS (GetVal $r $colCo)
    $touched = ToDate (GetVal $r $colTouched)
    $sc      = ToDate (GetVal $r $colSc)
    $added   = ToDate (GetVal $r $colAdded)
    $created = ToDate (GetVal $r $colCreated)
    $tg      = EscJS (GetVal $r $colTg)
    $owner   = EscJS (GetVal $r $colOwner)
    $stage   = EscJS (GetVal $r $colStage)
    $email   = EscJS (GetVal $r $colEmail)
    $au      = EscJS (GetVal $r $colAu)
    $title   = EscJS (GetVal $r $colTitle)
    $persona = EscJS (GetVal $r $colPersona)
    $active   = EscJS (GetVal $r $colActive)
    $finished = EscJS (GetVal $r $colFinished)
    $entries.Add("{`"id`":`"$id`",`"co`":`"$co`",`"touched`":`"$touched`",`"sc`":`"$sc`",`"added`":`"$added`",`"created`":`"$created`",`"tg`":`"$tg`",`"owner`":`"$owner`",`"stage`":`"$stage`",`"email`":`"$email`",`"au`":`"$au`",`"ti`":`"$title`",`"pn`":`"$persona`",`"as`":`"$active`",`"fs`":`"$finished`"}")
    if ($r % 25000 -eq 0) {
        Write-Host "Processed $r / $lastRow ($($entries.Count) entries)"
        # Flush to disk incrementally to avoid memory issues
    }
}

$wb.Close($false); $excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$js = "window.PROSPECT_DATA=[" + ($entries -join ",") + "];"
[System.IO.File]::WriteAllText($outFile, $js, [System.Text.Encoding]::UTF8)

$kb = [Math]::Round((Get-Item $outFile).Length / 1KB)
Write-Host "Prospects written: $($entries.Count) - ${kb}KB"
