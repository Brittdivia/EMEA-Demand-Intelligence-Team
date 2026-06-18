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
Write-Host "Key cols - Id:$($headers['Id']) Touched At:$($headers['Touched At']) Stage Changed At:$($headers['Stage Changed At']) Tags:$($headers['Tags'])"

function ToDate($val) {
    if (-not $val) { return "" }
    $s = "$val".Trim()
    if ($s -match '(\d{4}-\d{2}-\d{2})') { return $Matches[1] }
    try { return ([DateTime]::Parse($s)).ToString("yyyy-MM-dd") } catch { return "" }
}

function EscJS($s) {
    $s = "$s".Trim()
    $s = $s -replace '\\', '\\'
    $s = $s -replace '"',  '\"'
    $s = $s -replace "`r", ''
    $s = $s -replace "`n", ' '
    $s = $s -replace "`t", ' '
    return $s
}

$entries = [System.Collections.Generic.List[string]]::new()
$lastRow = $ws.UsedRange.Rows.Count

for ($r = 2; $r -le $lastRow; $r++) {
    $id      = $ws.Cells.Item($r, $headers["Id"]).Value2;           if (-not $id) { continue }
    $co      = EscJS $ws.Cells.Item($r, $headers["Company"]).Value2
    $touched = ToDate $ws.Cells.Item($r, $headers["Touched At"]).Value2
    $sc      = ToDate $ws.Cells.Item($r, $headers["Stage Changed At"]).Value2
    $added   = ToDate $ws.Cells.Item($r, $headers["Added At"]).Value2
    $created = ToDate $ws.Cells.Item($r, $headers["Created At"]).Value2
    $tg      = EscJS $ws.Cells.Item($r, $headers["Tags"]).Value2
    $owner   = EscJS $ws.Cells.Item($r, $headers["owner"]).Value2
    $stage   = EscJS $ws.Cells.Item($r, $headers["Stage Name"]).Value2
    $email   = EscJS $ws.Cells.Item($r, $headers["Email"]).Value2
    $au      = EscJS $ws.Cells.Item($r, $headers["Assigned Users"]).Value2

    $entries.Add("{`"id`":`"$id`",`"co`":`"$co`",`"touched`":`"$touched`",`"sc`":`"$sc`",`"added`":`"$added`",`"created`":`"$created`",`"tg`":`"$tg`",`"owner`":`"$owner`",`"stage`":`"$stage`",`"email`":`"$email`",`"au`":`"$au`"}")
}

$wb.Close($false); $excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$js = "window.PROSPECT_DATA=[" + ($entries -join ",") + "];"
[System.IO.File]::WriteAllText($outFile, $js, [System.Text.Encoding]::UTF8)

$kb = [Math]::Round((Get-Item $outFile).Length / 1KB)
Write-Host "Prospects written: $($entries.Count) - ${kb}KB"
