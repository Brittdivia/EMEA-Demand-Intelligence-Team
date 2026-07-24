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
# Read all lines and parse manually to handle duplicate column headers
$allLines = [System.IO.File]::ReadAllLines($csvPath, [System.Text.Encoding]::UTF8)
Write-Host "Total lines: $($allLines.Count)"

# Parse header — make unique
$rawHeaders = $allLines[0].Split(',') | ForEach-Object { $_.Trim('"').Trim() }
$seenH = @{}; $uniqueHeaders = @()
foreach ($h in $rawHeaders) {
    if ($seenH.ContainsKey($h)) { $seenH[$h]++; $uniqueHeaders += "${h}_$($seenH[$h])" }
    else { $seenH[$h] = 0; $uniqueHeaders += $h }
}

# File has no header row — use fixed column positions based on Outreach Sequence States export format
$colSid  = 0   # Sequence ID
$colPid  = 2   # Prospect ID
$colTags = 17  # Tags (quoted, comma-separated)
$colEm   = 20  # Emailed?
$colOp   = 21  # Opened?
$colCl   = 22  # Clicked?
$colRe   = 23  # Replied?
Write-Host "Using fixed columns: SID=$colSid PID=$colPid Tags=$colTags Em=$colEm Op=$colOp Cl=$colCl Re=$colRe"

function ParseCsvLine($line) {
    $result = @()
    $inQuote = $false; $current = ""
    foreach ($ch in $line.ToCharArray()) {
        if ($ch -eq '"') { $inQuote = !$inQuote }
        elseif ($ch -eq ',' -and -not $inQuote) { $result += $current; $current = "" }
        else { $current += $ch }
    }
    $result += $current
    return $result
}

$rows = [System.Collections.Generic.List[string[]]]::new()
for ($i = 0; $i -lt $allLines.Count; $i++) {
    if ([string]::IsNullOrWhiteSpace($allLines[$i])) { continue }
    $rows.Add((ParseCsvLine $allLines[$i]))
}
Write-Host "Rows: $($rows.Count)"

# 1. Build tag -> Set of sequence IDs
$tagToSids = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.HashSet[string]]]::new()

foreach ($row in $rows) {
    $sid = if ($colSid -ge 0 -and $colSid -lt $row.Count) { $row[$colSid].Trim() } else { "" }
    if (-not $sid) { continue }
    $tags = if ($colTags -ge 0 -and $colTags -lt $row.Count) { $row[$colTags] } else { "" }
    foreach ($t in $tags.Split(',')) {
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
    $sid    = if ($colSid -ge 0 -and $colSid -lt $row.Count) { $row[$colSid].Trim() } else { "" }
    $prospId = if ($colPid -ge 0 -and $colPid -lt $row.Count) { $row[$colPid].Trim() } else { "" }
    if (-not $sid -or -not $prospId) { continue }
    $key = "$sid|$prospId"
    if (-not $seen.Add($key)) { continue }
    $em = if ($colEm -ge 0 -and $colEm -lt $row.Count -and $row[$colEm] -eq "Yes") { "1" } else { "0" }
    $op = if ($colOp -ge 0 -and $colOp -lt $row.Count -and $row[$colOp] -eq "Yes") { "1" } else { "0" }
    $cl = if ($colCl -ge 0 -and $colCl -lt $row.Count -and $row[$colCl] -eq "Yes") { "1" } else { "0" }
    $re = if ($colRe -ge 0 -and $colRe -lt $row.Count -and $row[$colRe] -eq "Yes") { "1" } else { "0" }
    $flags.Add("{""sid"":""$(EscJ $sid)"",""pid"":""$(EscJ $prospId)"",""em"":$em,""op"":$op,""cl"":$cl,""re"":$re}")
}

Write-Host "Prospect flags: $($flags.Count) entries"
$js2 = "window.SEQ_PROSPECT_FLAGS=[" + ($flags -join ',') + "];"
[System.IO.File]::WriteAllText("$outDir\data-seqstates-flags.js", $js2, [System.Text.Encoding]::UTF8)
Write-Host "SEQ_PROSPECT_FLAGS done - $([Math]::Round((Get-Item "$outDir\data-seqstates-flags.js").Length/1KB))KB"
