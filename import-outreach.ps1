# import-outreach.ps1
# Rebuilds data-out.js from Sequence Stats + campaign calendar sequences
# Uses campaign calendar as source of truth for tracked sequences

$seqStatsPath = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Outreach\Sequence_Stats_2026-06-02.csv"
$outDir = "C:\Users\I572929\campaign-calendar-site"

# Load Sequence Stats
$seqStats = Import-Csv $seqStatsPath
$seqStatsMap = @{}
$seqStats | ForEach-Object {
    $sid = $_."Sequence ID".Trim()
    if ($sid) { $seqStatsMap[$sid] = $_ }
}
Write-Host "Sequence Stats entries: $($seqStatsMap.Count)"

# Build sequence->campaign info map from data-camp.js
$campJs = [System.IO.File]::ReadAllText("$outDir\data-camp.js")
$campSids = [regex]::Matches($campJs, '"Sequence ID":"([^"]+)"') | ForEach-Object { $_.Groups[1].Value.Trim() } | Where-Object { $_ } | Sort-Object -Unique
Write-Host "Sequences in campaign calendar: $($campSids.Count)"

# Build sid->campaign info from data-camp.js
$sidToCamp = @{}
$campEntries = [regex]::Matches($campJs, '\{[^}]+\}')
foreach ($entry in $campEntries) {
    $sid = [regex]::Match($entry.Value, '"Sequence ID":"([^"]+)"').Groups[1].Value.Trim()
    if (-not $sid) { continue }
    if ($sidToCamp.ContainsKey($sid)) { continue }
    $sidToCamp[$sid] = @{
        wbs    = [regex]::Match($entry.Value, '"Campaign/WBS Code":"([^"]+)"').Groups[1].Value
        name   = [regex]::Match($entry.Value, '"Campaign Name":"([^"]+)"').Groups[1].Value
        dm     = [regex]::Match($entry.Value, '"Demand Manager":"([^"]+)"').Groups[1].Value
        reg2   = [regex]::Match($entry.Value, '"Region Name \(level 2\)":"([^"]+)"').Groups[1].Value
        reg3   = [regex]::Match($entry.Value, '"Region Name \(level 3\)":"([^"]+)"').Groups[1].Value
        sb     = [regex]::Match($entry.Value, '"Sales Bag":"([^"]+)"').Groups[1].Value
        iac    = [regex]::Match($entry.Value, '"IAC":"([^"]+)"').Groups[1].Value
        sod    = [regex]::Match($entry.Value, '"SoD":"([^"]+)"').Groups[1].Value
    }
}

# Also include sequences from outreach.json for backward compatibility
$outBase = Get-Content "$outDir\outreach.json" -Raw | ConvertFrom-Json
$outBaseMap = @{}
$outBase | ForEach-Object { $sid = "$($_.'Sequence ID')".Trim(); if ($sid) { $outBaseMap[$sid] = $_ } }
Write-Host "Sequences in outreach.json: $($outBaseMap.Count)"

# Build records for all tracked sequences
$outRecords = [System.Collections.Generic.List[object]]::new()

foreach ($sid in $campSids) {
    $stats = $seqStatsMap[$sid]
    if (-not $stats) { continue }  # Skip if not in Sequence Stats

    # Get base info from outreach.json if available, else build from stats
    $base = $outBaseMap[$sid]

    $obj = [ordered]@{}
    $ci = $sidToCamp[$sid]
    if ($base) {
        $base.PSObject.Properties | ForEach-Object { $obj[$_.Name] = $_.Value }
        # Override with campaign calendar info if available
        if ($ci) {
            if ($ci.wbs)  { $obj["Campaign/WBS Code"]        = $ci.wbs  }
            if ($ci.name) { $obj["Campaign Name"]            = $ci.name }
            if ($ci.dm)   { $obj["Demand Manager"]           = $ci.dm   }
            if ($ci.reg2) { $obj["Region Name (level 2)"]    = $ci.reg2 }
            if ($ci.reg3) { $obj["Region Name (level 3)"]    = $ci.reg3 }
        }
    } else {
        $obj["Sequence ID"]              = $sid
        $obj["Sequence Name"]            = $stats."Sequence Name"
        $obj["Owner"]                    = $stats."Owner"
        $obj["Region Name (level 2)"]    = if ($ci) { $ci.reg2 } else { "" }
        $obj["Region Name (level 3)"]    = if ($ci) { $ci.reg3 } else { "" }
        $obj["Demand Manager"]           = if ($ci) { $ci.dm }   else { "" }
        $obj["Campaign Name"]            = if ($ci) { $ci.name } else { "" }
        $obj["Campaign/WBS Code"]        = if ($ci) { $ci.wbs }  else { "" }
        $obj["Sales Bag"]                = if ($ci) { $ci.sb }   else { "" }
        $obj["IAC"]                      = if ($ci) { $ci.iac }  else { "" }
        $obj["SoD"]                      = if ($ci) { $ci.sod }  else { "" }
        $obj["Open Rate"]                = ""
        $obj["Reply Rate"]               = ""
        $obj["Prospects Invited"]        = 0
        $obj["Engagement Score"]         = ""
        $obj["Tags"]                     = $stats."Tags"
    }

    $obj["Prospects - Total"]     = $stats."Prospects - Total"
    $obj["Prospects - Active"]    = $stats."Prospects - Active"
    $obj["Prospects - Finished"]  = $stats."Prospects - Finished"
    $obj["Prospects - Delivered"] = $stats."Prospects - Delivered"
    $obj["Prospects - Opened"]    = $stats."Prospects - Opened"
    $obj["Prospects - Replied"]   = $stats."Prospects - Replied"
    $obj["Emails - Deliveries"]   = $stats."Emails - Deliveries"
    $obj["Emails - Opens"]        = $stats."Emails - Opens"
    $obj["Emails - Replies"]      = $stats."Emails - Replies"
    $obj["Emails - Bounces"]      = $stats."Emails - Bounces"
    $obj["Meetings Booked"]       = 0
    $obj["Accounts Added"]        = 0
    $obj["Last used at"]          = ""

    $outRecords.Add([PSCustomObject]$obj)
}

Write-Host "Total records: $($outRecords.Count)"
$json = $outRecords | ConvertTo-Json -Compress -Depth 3
"window.OUT_DATA=$json;" | Set-Content "$outDir\data-out.js" -Encoding UTF8
Write-Host "Done - $([Math]::Round((Get-Item "$outDir\data-out.js").Length/1KB))KB"
