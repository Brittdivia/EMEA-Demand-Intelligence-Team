# import-prospects-csv.ps1
# Rebuilds data-prospects.js from combined Prospects CSV
# Strips unused fields, deduplicates by ID

$csvPath = "C:\Users\I572929\Downloads\Prospects_combined.csv"
$outDir  = "C:\Users\I572929\campaign-calendar-site"

Write-Host "Reading CSV headers..."
$reader = [System.IO.StreamReader]::new($csvPath, [System.Text.Encoding]::UTF8)
$headerLine = $reader.ReadLine()
# Remove BOM if present
$headerLine = $headerLine.TrimStart([char]0xFEFF)
$headers = $headerLine.Split(',')

# Find column indices for fields we need
function FindCol($name) {
    for($i=0; $i -lt $headers.Count; $i++) {
        if($headers[$i].Trim('"') -eq $name) { return $i }
    }
    return -1
}

$colID      = FindCol 'ID'
$colCo      = FindCol 'Company'
$colTouched = FindCol 'Touched At'
$colSc      = FindCol 'Stage Changed At'
$colCreated = FindCol 'Created At'
$colTg      = FindCol 'Tags'
$colStage   = FindCol 'Stage Name'
$colEmail   = FindCol 'Email'
$colAu      = FindCol 'Assigned Users'
$colPn      = FindCol 'Persona Name'
$colAs      = FindCol 'Active Sequences'
$colFs      = FindCol 'Finished Sequences'

Write-Host "Key columns: ID=$colID Co=$colCo Touched=$colTouched Stage=$colStage"

function EscJS($s) {
    if ($null -eq $s) { return '' }
    return [string]$s -replace '\\','\\' -replace '"','\"' -replace "`r`n",' ' -replace "`n",' ' -replace "`r",' '
}

function GetField($fields, $idx) {
    if($idx -lt 0 -or $idx -ge $fields.Count) { return '' }
    $v = $fields[$idx].Trim()
    if($v.StartsWith('"') -and $v.EndsWith('"')) { $v = $v.Substring(1, $v.Length-2) }
    return $v
}

$entries = [System.Collections.Generic.List[string]]::new()
$seenIds = [System.Collections.Generic.HashSet[string]]::new()
$lineNum = 0

while(-not $reader.EndOfStream) {
    $line = $reader.ReadLine()
    $lineNum++
    if($lineNum % 100000 -eq 0) { Write-Host "  Processed $lineNum lines..." }

    # Simple CSV split (handles quoted fields with commas)
    $fields = [System.Collections.Generic.List[string]]::new()
    $inQuote = $false
    $current = [System.Text.StringBuilder]::new()
    foreach($ch in $line.ToCharArray()) {
        if($ch -eq '"') { $inQuote = -not $inQuote }
        elseif($ch -eq ',' -and -not $inQuote) { $fields.Add($current.ToString()); $current.Clear() | Out-Null }
        else { $current.Append($ch) | Out-Null }
    }
    $fields.Add($current.ToString())

    $id = GetField $fields $colID
    if(-not $id -or -not $seenIds.Add($id)) { continue }

    $parts = @(
        """id"":""$(EscJS $id)""",
        """co"":""$(EscJS (GetField $fields $colCo))""",
        """touched"":""$(EscJS (GetField $fields $colTouched))""",
        """sc"":""$(EscJS (GetField $fields $colSc))""",
        """created"":""$(EscJS (GetField $fields $colCreated))""",
        """tg"":""$(EscJS (GetField $fields $colTg))""",
        """stage"":""$(EscJS (GetField $fields $colStage))""",
        """email"":""$(EscJS (GetField $fields $colEmail))""",
        """au"":""$(EscJS (GetField $fields $colAu))""",
        """pn"":""$(EscJS (GetField $fields $colPn))""",
        """as"":""$(EscJS (GetField $fields $colAs))""",
        """fs"":""$(EscJS (GetField $fields $colFs))"""
    )
    $entries.Add('{' + ($parts -join ',') + '}')
}
$reader.Close()

Write-Host "Writing $($entries.Count) prospects..."
$js = "window.PROSPECT_DATA=[" + ($entries -join ',') + "];"
[System.IO.File]::WriteAllText("$outDir\data-prospects.js", $js, [System.Text.Encoding]::UTF8)
Write-Host "Done: $([Math]::Round((Get-Item "$outDir\data-prospects.js").Length/1MB, 1))MB"
