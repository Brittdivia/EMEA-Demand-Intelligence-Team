# Converts SPR Add on.xlsx -> data-spr.js
# Source: "SPR Dashboard KPIs (16)" sheet
# Cols: 2=IBP Market Unit (mu), 3=IBP Market Unit-1 (mu1), 9=Quarter Flag (qtr),
#        11=Cloud N&U Quota (quota), 12=Team (team), 13=IAC (iac)

$xlsxPath = 'C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Targets\SPR Add on.xlsx'
$outFile  = 'C:\Users\I572929\campaign-calendar-site\data-spr.js'

$muMap = @{
    'BeNeLux'      = 'Benelux'
    'France'       = 'France'
    'Italy'        = 'Italy'
    'MEA North'    = 'MEA North'
    'MEA South'    = 'MEA South'
    'Nordic'       = 'Nordic'
    'Southern Europe' = 'Southern Europe'
    'UKI'          = 'UKI'
}

$qtrMap = @{
    'Q126' = '2026-Q1'
    'Q226' = '2026-Q2'
    'Q326' = '2026-Q3'
    'Q426' = '2026-Q4'
}

function EscJson($s) {
    $cr = [string][char]13; $lf = [string][char]10
    return $s.Replace('\','\\').Replace('"','\"').Replace($cr,'').Replace($lf,'')
}

Write-Host "Opening $xlsxPath ..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($xlsxPath, 0, $true)
$ws = $wb.Sheets.Item(1)
$lastRow = [Math]::Min($ws.Cells.Item(2,1).End(-4121).Row, 10000)
Write-Host "Rows: $($lastRow - 1)"

$range = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item($lastRow, 13))
$data  = $range.Value2
Write-Host 'Bulk read done.'

$sw = [System.IO.StreamWriter]::new($outFile, $false, [System.Text.Encoding]::UTF8)
$sw.Write('window.SPR_DATA=[')
$first = $true
$rowCount = 0
$invCulture = [System.Globalization.CultureInfo]::InvariantCulture

for ($r = 2; $r -le $lastRow; $r++) {
    $muRaw  = [string]$data.GetValue($r, 2)
    if ($muRaw -eq '') { continue }

    $mu1Raw = [string]$data.GetValue($r, 3)
    $qtrRaw = [string]$data.GetValue($r, 9)
    $teamRaw= [string]$data.GetValue($r, 12)
    $iacRaw = [string]$data.GetValue($r, 13)
    $qv     = $data.GetValue($r, 11)

    $mu  = if ($muMap.ContainsKey($muRaw))  { $muMap[$muRaw] }  else { $muRaw }
    $mu1 = $mu1Raw.Trim()
    $qtr = if ($qtrMap.ContainsKey($qtrRaw)) { $qtrMap[$qtrRaw] } else { $qtrRaw }

    $quota = 0.0
    if ($qv -ne $null) {
        [double]::TryParse([string]$qv, [System.Globalization.NumberStyles]::Any, $invCulture, [ref]$quota) | Out-Null
    }

    $quotaStr = $quota.ToString('G', $invCulture)

    $obj = '{"mu":"' + (EscJson $mu) + '","mu1":"' + (EscJson $mu1) + '","team":"' + (EscJson $teamRaw.Trim()) + '","iac":"' + (EscJson $iacRaw.Trim()) + '","qtr":"' + (EscJson $qtr) + '","quota":' + $quotaStr + '}'

    if (-not $first) { $sw.Write(',') }
    $sw.Write($obj)
    $first = $false
    $rowCount++
}

$sw.Write('];')
$sw.Close()

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws)    | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)    | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

$sz = (Get-Item $outFile).Length
Write-Host "Done: $rowCount records -> $outFile ($([int]($sz/1KB)) KB)"
