# Converts Week 42 CSV -> data-pipe-w42.js
# Week 42 CSV has TWO header rows: row 0 is garbage (,,,...,Measures,ACV kEUR), row 1 is real headers
# Uses TextFieldParser for correct quoted-CSV handling

$csvPath = 'C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Pipeline\week 42.csv'
$outFile = 'C:\Users\I572929\campaign-calendar-site\data-pipe-w42.js'

Add-Type -AssemblyName Microsoft.VisualBasic

$fieldMap = [ordered]@{
    'Region Lvl 2'                    = 'Region Lvl 2'
    'Region Lvl 3'                    = 'Region Lvl 3'
    'DRM Category'                    = 'DRM Category'
    'Account Name'                    = 'Account Name'
    'Opportunity ID'                  = 'Opportunity ID'
    'Closing Qtr'                     = 'Closing Qtr'
    'IAC (Engagement Model)'          = 'IAC (Engagement Model)'
    'Opp Campaign ID'                 = 'Opp Campaign ID'
    'Solution Area (L1)'              = 'Solution Area (L1)'
    'Sub-Solution Area (L2)'          = 'Sub-Solution Area (L2)'
    'Create Quarter'                  = 'Create Quarter'
    'IC_ACCOUNT_SRG_TEXT'             = 'IC_ACCOUNT_SRG_TEXT'
    'GB Identifier'                   = 'MM Identifier'
    'SDE Handover Date'               = 'SDE Handover Date'
    'SDE Initial Engagement Name'     = 'SDE Initial Engagement Name'
    'SDE Primary Engagement Name'     = 'SDE Primary Engagement Name'
    'SDE Secondary Engagement Name'   = 'SDE Secondary Engagement Name'
    'SDE Handover Role'               = 'SDE Handover Role'
    'Opp Description'                 = 'Opp Description'
    'TCP Pipeline Source Desc'        = 'TCP Pipeline Source Desc'
    'SAP Mastercode'                  = 'SAP Mastercode'
    'Create Date'                     = 'Create Date'
    'SDE Engagement Type'             = 'SDE Engagement Type'
    'Untouched Days'                  = 'Untouched Days'
    'Channel Partner'                 = 'Channel Partner'
    'RBC'                             = 'RBC'
    'ACV kEUR'                        = 'Eur'
    'MM Pipe Source'                  = 'MM Pipe Source'
    'Local Acct Owner'                = 'Local Account Owner Name'
}
$numericFields = @('Eur', 'Untouched Days')
$invCulture = [System.Globalization.CultureInfo]::InvariantCulture
$numStyles  = [System.Globalization.NumberStyles]::Any

function EscJson($s) {
    $cr = [string][char]13; $lf = [string][char]10
    return $s.Replace('\','\\').Replace('"','\"').Replace($cr,'').Replace($lf,'')
}

Write-Host "=== Week 42 CSV -> $outFile ==="

$parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($csvPath, [System.Text.Encoding]::UTF8)
$parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
$parser.SetDelimiters(',')
$parser.HasFieldsEnclosedInQuotes = $true

# Row 0: garbage (,,,,...,Measures,ACV kEUR) — skip
$garbageRow = $parser.ReadFields()

# Row 1: real column headers
$headerRow = $parser.ReadFields()
$colCount = $headerRow.Count

# Last column is ACV kEUR (empty header in row 1, "ACV kEUR" in row 0)
if ($headerRow[$colCount - 1] -eq '') { $headerRow[$colCount - 1] = 'ACV kEUR' }

Write-Host "  Columns: $colCount"

# Build column index map
$colIdxMap = @{}
foreach ($csvName in $fieldMap.Keys) {
    for ($i = 0; $i -lt $colCount; $i++) {
        if ($headerRow[$i].Trim() -eq $csvName) { $colIdxMap[$csvName] = $i; break }
    }
}

# Report any missing fields
foreach ($csvName in $fieldMap.Keys) {
    if ($null -eq $colIdxMap[$csvName]) {
        Write-Host "  WARNING: column '$csvName' not found in headers"
    }
}
Write-Host '  Column map built.'

# Build ordered output field names from fieldMap
$outFields = @($fieldMap.Values)

$sw = [System.IO.StreamWriter]::new($outFile, $false, [System.Text.Encoding]::UTF8)
# Compact columnar format: {h:[headers], d:[[row],[row],...]}
$headerJson = ($outFields | ForEach-Object { '"' + $_ + '"' }) -join ','
$sw.Write('window.PIPE_DATA_W42={h:[' + $headerJson + '],d:[')
$firstRecord = $true
$rowCount = 0

while (-not $parser.EndOfData) {
    $fields = $parser.ReadFields()
    $sb = New-Object System.Text.StringBuilder
    $sb.Append('[') | Out-Null
    $firstField = $true
    foreach ($csvName in $fieldMap.Keys) {
        $outName = $fieldMap[$csvName]
        $idx = $colIdxMap[$csvName]
        if ($null -eq $idx) { continue }
        $raw = if ($idx -lt $fields.Count) { $fields[$idx].Trim() } else { '' }
        if (-not $firstField) { $sb.Append(',') | Out-Null }
        $firstField = $false
        if ($outName -in $numericFields) {
            $num = 0.0
            if ([double]::TryParse($raw, $numStyles, $invCulture, [ref]$num)) {
                $sb.Append($num.ToString('G', $invCulture)) | Out-Null
            } else { $sb.Append('0') | Out-Null }
        } else {
            $sb.Append('"' + (EscJson $raw) + '"') | Out-Null
        }
    }
    $sb.Append(']') | Out-Null
    if (-not $firstRecord) { $sw.Write(',') }
    $sw.Write($sb.ToString())
    $firstRecord = $false
    $rowCount++
    if ($rowCount % 20000 -eq 0) { Write-Host "  Rows: $rowCount" }
}

$parser.Close()
$sw.Write(']};')
$sw.Close()
$sz = (Get-Item $outFile).Length
Write-Host "  Done: $rowCount rows -> $outFile ($([int]($sz/1MB)) MB)"
