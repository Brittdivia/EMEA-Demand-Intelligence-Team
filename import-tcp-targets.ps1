$csvPath = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Targets\Britt TCP Targets.csv"
$outDir  = "C:\Users\I572929\campaign-calendar-site"

Write-Host "Reading CSV..."
$rows = Import-Csv $csvPath

Write-Host "Processing $($rows.Count) rows..."

function EscapeJson($val) {
    if ($null -eq $val) { return '' }
    $s = [string]$val
    $s = $s.Replace('\','\\').Replace('"','\"').Replace("`r`n",' ').Replace("`n",' ').Replace("`r",' ').Replace("`t",' ')
    return $s
}

$entries = [System.Collections.Generic.List[string]]::new()

foreach ($row in $rows) {
    $tcpReqd = 0.0
    [double]::TryParse($row.'TCP Req''d', [ref]$tcpReqd) | Out-Null

    $obj = '{"Region Lvl 2":"'       + (EscapeJson $row.'Region Lvl 2')            + '"' `
         + ',"Region Lvl 3":"'       + (EscapeJson $row.'Region Lvl 3')            + '"' `
         + ',"Segment":"'            + (EscapeJson $row.'Segment')                 + '"' `
         + ',"Solution Area (L1)":"' + (EscapeJson $row.'Solution Area (L1)')      + '"' `
         + ',"Sub-Solution Area (L2)":"' + (EscapeJson $row.'Sub-Solution Area (L2)') + '"' `
         + ',"TCP Req d":' + $tcpReqd `
         + '}'
    $entries.Add($obj)
}

Write-Host "Writing data-tcp-targets.js ($($entries.Count) rows)..."
$js = 'window.TCP_TARGETS_BRITT=[' + ($entries -join ',') + '];'
[System.IO.File]::WriteAllText("$outDir\data-tcp-targets.js", $js, [System.Text.Encoding]::UTF8)

Write-Host "Done."
