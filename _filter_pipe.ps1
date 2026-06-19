$campJs = Get-Content "C:\Users\I572929\campaign-calendar-site\data-camp.js" -Raw
$wbsCodes = @{}
[regex]::Matches($campJs, '"Campaign/WBS Code":"([^"]+)"') | ForEach-Object { 
    $raw = $_.Groups[1].Value.Trim()
    # Remove any \r\n or whitespace from the code
    $clean = $raw -replace '[\r\n\s]',''
    if ($clean) { $wbsCodes[$clean] = 1 }
}
Write-Host "WBS codes: $($wbsCodes.Count)"
# Show samples
$wbsCodes.Keys | Select-Object -First 5 | ForEach-Object { Write-Host "  '$_'" }

$csvTemp = "C:\Users\I572929\campaign-calendar-site\_pipe_temp.csv"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open("C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Pipeline\Week 22 DL.xlsx", 0, $true)
$wb.SaveAs($csvTemp, 6)
$wb.Close($false); $excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "CSV saved"

function ToISO($d) {
    if (-not $d -or $d.Trim() -eq "") { return "" }
    try { return [DateTime]::Parse($d.Trim()).ToString("yyyy-MM-dd") } catch { return "" }
}

$keep = @("Opp Campaign ID","Opportunity ID","Eur","DRM Category","SDE Handover Date","Create Date","Create Quarter","Closing Qtr","TCP Pipeline Source Desc","MM Identifier","RBC","SAP Mastercode","IAC (Engagement Model)","Solution Area (L1)","Sub-Solution Area (L2)","Region Lvl 2")
$csv = Import-Csv $csvTemp
$records = $csv | Where-Object { 
    $wbs = $_."Opp Campaign ID".Trim() -replace '[\r\n\s]',''
    $wbsCodes.ContainsKey($wbs) 
} | ForEach-Object {
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
Write-Host "Filtered records: $($records.Count)"
$json = $records | ConvertTo-Json -Compress
"window.PIPE_DATA=$json;" | Set-Content "C:\Users\I572929\campaign-calendar-site\data-pipe.js" -Encoding UTF8
Write-Host "Done - $([Math]::Round((Get-Item 'C:\Users\I572929\campaign-calendar-site\data-pipe.js').Length/1MB,1)) MB"
