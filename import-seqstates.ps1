# import-seqstates.ps1
# Reads Sequence_States CSV and builds:
#   1. window.SEQ_STATES_TAG_MAP  — tag -> [sequenceIds]
#   2. window.SEQ_PROSPECT_FLAGS  — [{sid,pid,em,op,cl,re}] per prospect+sequence
# Run manually when you get a new Sequence States export - update the path below

$csvPath = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Outreach\Sequence_States_2026-06-06.csv"
$outDir  = "C:\Users\I572929\campaign-calendar-site"

if (-not (Test-Path $csvPath)) {
    Write-Host "ERROR: File not found: $csvPath"
    Write-Host "Please update the path in import-seqstates.ps1 to point to the latest Sequence States export."
    exit 1
}

Write-Host "Reading $csvPath..."
$rows = Import-Csv -Path $csvPath -Encoding UTF8
Write-Host "Rows: $($rows.Count)"

# 1. Build tag -> Set of sequence IDs
$tagToSids = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.HashSet[string]]]::new()

foreach ($row in $rows) {
    $sid = $row."Sequence ID".Trim()
    if (-not $sid) { continue }
    foreach ($t in $row."Tags".Split(',')) {
        $t = $t.Trim()
        if (-not $t) { continue }
        if (-not $tagToSids.ContainsKey($t)) { $tagToSids[$t] = [System.Collections.Generic.HashSet[string]]::new() }
        [void]$tagToSids[$t].Add($sid)
    }
}

Write-Host "Unique tags: $($tagToSids.Count)"

$pairs = [System.Collections.Generic.List[string]]::new()
foreach ($kvp in $tagToSids.GetEnumerator()) {
    $tag = $kvp.Key.Replace('\','\\').Replace('"','\"')
    $pairs.Add('"' + $tag + '":["' + ($kvp.Value -join '","') + '"]')
}

$js1 = "window.SEQ_STATES_TAG_MAP={" + ($pairs -join ',') + "};"
[System.IO.File]::WriteAllText("$outDir\data-seqstates-tags.js", $js1, [System.Text.Encoding]::UTF8)
Write-Host "SEQ_STATES_TAG_MAP done - $($tagToSids.Count) tags, $([Math]::Round((Get-Item "$outDir\data-seqstates-tags.js").Length/1KB))KB"

# 2. Build per-prospect flags dataset
function EscJ($v) { return $v.Replace('\','\\').Replace('"','\"') }

$flags = [System.Collections.Generic.List[string]]::new()
$seen  = [System.Collections.Generic.HashSet[string]]::new()

foreach ($row in $rows) {
    $sid = $row."Sequence ID".Trim()
    $prospId = $row."Prospect ID".Trim()
    if (-not $sid -or -not $prospId) { continue }
    $key = "$sid|$prospId"
    if (-not $seen.Add($key)) { continue }
    $em = if ($row."Emailed?"  -eq "Yes") { "1" } else { "0" }
    $op = if ($row."Opened?"   -eq "Yes") { "1" } else { "0" }
    $cl = if ($row."Clicked?"  -eq "Yes") { "1" } else { "0" }
    $re = if ($row."Replied?"  -eq "Yes") { "1" } else { "0" }
    $flags.Add("{""sid"":""$(EscJ $sid)"",""pid"":""$(EscJ $prospId)"",""em"":$em,""op"":$op,""cl"":$cl,""re"":$re}")
}

Write-Host "Prospect flags: $($flags.Count) entries"
$js2 = "window.SEQ_PROSPECT_FLAGS=[" + ($flags -join ',') + "];"
[System.IO.File]::WriteAllText("$outDir\data-seqstates-flags.js", $js2, [System.Text.Encoding]::UTF8)
Write-Host "SEQ_PROSPECT_FLAGS done - $([Math]::Round((Get-Item "$outDir\data-seqstates-flags.js").Length/1KB))KB"
