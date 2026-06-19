$site = "C:\Users\I572929\campaign-calendar-site"
$pipeFile = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Pipeline\Week 22 DL.xlsx"
$csvTemp = "$site\_pipe_temp.csv"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($pipeFile, 0, $true)
$wb.SaveAs($csvTemp, 6)
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "CSV saved: $([Math]::Round((Get-Item $csvTemp).Length/1MB,1)) MB"

function ToISO($d) {
    if (-not $d -or $d.Trim() -eq "") { return "" }
    try { return [DateTime]::Parse($d.Trim()).ToString("yyyy-MM-dd") } catch { return "" }
}

$keep = @("Opp Campaign ID","Opportunity ID","Eur","DRM Category","SDE Handover Date","Create Date","Create Quarter","Closing Qtr","TCP Pipeline Source Desc","MM Identifier","RBC","SAP Mastercode","IAC (Engagement Model)","Solution Area (L1)","Sub-Solution Area (L2)","Account Name","Opp Description","Region Lvl 2")
$csv = Import-Csv $csvTemp
$records = $csv | ForEach-Object {
    $obj = [ordered]@{}
    foreach ($h in $keep) {
        $val = if ($_.PSObject.Properties[$h]) { $_.$h } else { "" }
        if ($h -match "Date") { $val = ToISO $val }
        elseif ($h -eq "Eur") { try { $val = [double]($val -replace '[^0-9.\-]','') } catch { $val = 0 } }
        $obj[$h] = $val
    }
    [PSCustomObject]$obj
}
Remove-Item $csvTemp -Force
Write-Host "Records: $($records.Count)"
$json = $records | ConvertTo-Json -Compress
"window.PIPE_DATA=$json;" | Set-Content "$site\data-pipe.js" -Encoding UTF8
Write-Host "Done - $([Math]::Round((Get-Item "$site\data-pipe.js").Length/1MB,1)) MB"
