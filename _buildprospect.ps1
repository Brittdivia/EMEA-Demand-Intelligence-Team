$csv = Import-Csv "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Outreach\Sequence_States_2026-06-06.csv"
$outJson = Get-Content "C:\Users\I572929\campaign-calendar-site\outreach.json" -Raw | ConvertFrom-Json
$trackedSeqIDs = @{}
$outJson | ForEach-Object { $trackedSeqIDs["$($_.'Sequence ID')"] = 1 }

function Get-StageKey($row) {
    $stage = $row."Stage"; $state = $row."Sequence State"
    if ($stage -eq "Bad Email" -or $state -eq "Bounced" -or $state -eq "Pending") { return "Not Delivered" }
    if ($stage -eq "Replied") { return "Replied" }
    if ($stage -eq "Key Executive Contact") { return "Key Exec Contact" }
    if ($stage -eq "Unresponsive") { return "Unresponsive" }
    if ($stage -eq "Sequence Started") { return "Seq. Started" }
    if ($stage -eq "Requested Not to be Contacted" -or $state -eq "Opted Out") { return "Opt Out/RNTBC" }
    return "Other"
}

function Escape-JS($s) {
    $s = $s -replace '\\', '\\'
    $s = $s -replace '"', '\"'
    $s = $s -replace "`r", ''
    $s = $s -replace "`n", ' '
    $s = $s -replace "`t", ' '
    return $s
}

$entries = [System.Collections.Generic.List[string]]::new()
$count = 0
foreach ($row in $csv) {
    $seqId = "$($row.'Sequence ID')"
    if (-not $trackedSeqIDs.ContainsKey($seqId)) { continue }
    $prospId = "$($row.'Prospect ID')"
    $emailed = "$($row.'Emailed?')"
    $opened  = "$($row.'Opened?')"
    $replied = "$($row.'Replied?')"
    $co = Escape-JS $row."Company"
    $nm = Escape-JS $row."Name"
    $ti = Escape-JS $row."Title"
    $st = Get-StageKey $row
    if (-not $co) { continue }
    $entries.Add("{`"sid`":`"$seqId`",`"pid`":`"$prospId`",`"em`":`"$emailed`",`"op`":`"$opened`",`"re`":`"$replied`",`"co`":`"$co`",`"nm`":`"$nm`",`"ti`":`"$ti`",`"st`":`"$st`"}")
    $count++
}

Write-Host "Prospects: $count"

$js = "window.SEQ_PROSPECT_DATA=[" + ($entries -join ",") + "];"
[System.IO.File]::WriteAllText("C:\Users\I572929\campaign-calendar-site\data-seqprospect.js", $js, [System.Text.Encoding]::UTF8)
Write-Host "Written: $($js.Length) bytes"
Write-Host "Starts: $($js.Substring(0,50))"
Write-Host "Ends with ];" $js.TrimEnd().EndsWith('];')
