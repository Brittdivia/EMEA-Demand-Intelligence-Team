$site = "C:\Users\I572929\campaign-calendar-site"
$xlFile = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Calendars\Campaign Calendars combined.xlsx"
$csvTemp = "$site\_camp_temp.csv"
$tmpOut = "$site\data-camp.tmp"
$outFile = "$site\data-camp.js"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($xlFile, 0, $true)
$wb.SaveAs($csvTemp, 6)
$wb.Close($false); $excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "CSV saved: $([Math]::Round((Get-Item $csvTemp).Length/1MB,1)) MB"

$csv = Import-Csv $csvTemp
# Detect column names dynamically (handles trailing spaces, quote variants)
$includeCol = $csv[0].PSObject.Properties.Name | Where-Object { $_ -match 'Include' }
$tpColName  = $csv[0].PSObject.Properties.Name | Where-Object { $_ -match 'Target Pipeline - value' }
Write-Host "Include col: [$includeCol]  TP col: [$tpColName]"

$keep = @("Campaign Name","Campaign No","Campaign Origin","Campaign Priority","Campaign Type","Campaign/WBS Code","Demand Manager","Execution End Date","Execution Start Date","Executor","IAC","IB/NNN","Industry (MC)","Number of Accounts","Region Name (level 2)","Region Name (level 3)","Sales Bag","Sequence ID","SoD","Solution Area L1","Solution Area L2","Starting Quarter","Status","Sub Sales Bag","Activity Sub-Type","Campaign Objective","DG PROGRAM","ID")

$records = $csv | Where-Object { $_.$includeCol -eq 'TRUE' } | ForEach-Object {
    $obj = [ordered]@{}
    foreach ($h in $keep) {
        $key = if ($h -eq "Number of Accounts") { "# Accounts" } elseif ($h -eq "DG PROGRAM") { "DG Program" } else { $h }
        $obj[$key] = if ($_.PSObject.Properties[$h]) { $_.$h } else { "" }
    }
    # Parse TP value cleanly (col may have trailing space in name)
    $tpRaw = if ($_.PSObject.Properties[$tpColName]) { $_.$tpColName } else { "" }
    $tpClean = $tpRaw -replace '[^0-9\.]',''
    $obj["Target Pipeline - value kEUR"] = if ($tpClean) { try { [double]$tpClean } catch { 0 } } else { 0 }
    [PSCustomObject]$obj
}
Remove-Item $csvTemp -Force
Write-Host "Records: $($records.Count)"
$tpSum = ($records | Measure-Object -Property "Target Pipeline - value kEUR" -Sum).Sum
Write-Host "TP sum: $([Math]::Round($tpSum/1000,1))M"
$json = $records | ConvertTo-Json -Compress -Depth 2
"window.CAMP_DATA=$json;" | Set-Content $tmpOut -Encoding UTF8
Move-Item $tmpOut $outFile -Force
Write-Host "Done: $([Math]::Round((Get-Item $outFile).Length/1MB,1)) MB"
