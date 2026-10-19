# Converts Campaign Calendars combined.xlsx -> data-camp.js
# Uses bulk Range.Value2 read (one COM call) for speed
$xlsxPath = 'C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Calendars\Campaign Calendars combined 2.xlsx'
$outFile  = 'C:\Users\I572929\campaign-calendar-site\data-camp.js'

# Column map: [xlsxColIndex (1-based), outputFieldName, isDate]
$colMap = @(
    @(1,  'Campaign Name',                               $false),
    @(2,  'Campaign Origin',                             $false),
    @(3,  'Index',                                       $false),
    @(5,  'Status',                                      $false),
    @(6,  'SoD',                                         $false),
    @(7,  'Campaign Type',                               $false),
    @(8,  'IAC',                                         $false),
    @(9,  'IB/NNN',                                      $false),
    @(10, 'Starting Quarter',                            $false),
    @(11, 'Execution Start Date',                        $true),
    @(12, 'Execution End Date',                          $true),
    @(13, 'Campaign/WBS Code',                           $false),
    @(15, 'Digital Assets team support',                 $false),
    @(18, 'Sequence ID',                                 $false),
    @(19, 'Industry (MC)',                               $false),
    @(22, 'Sales Bag',                                   $false),
    @(23, 'Sub Sales Bag',                               $false),
    @(25, '# Accounts',                                  $false),
    @(26, 'Total Account (Validated)',                   $false),
    @(27, 'Existing Prospects by SDE',                   $false),
    @(29, 'Target Pipeline - value kEUR',                $false),
    @(31, 'Region Name (level 2)',                       $false),
    @(32, 'Region Name (level 3)',                       $false),
    @(35, 'Campaign Priority',                           $false),
    @(36, 'Demand Manager',                              $false),
    @(38, 'DG Program',                                  $false),
    @(42, 'Campaign Objective',                          $false),
    @(47, 'Activity Sub-Type',                           $false),
    @(53, 'Campaign No',                                 $false),
    @(76, 'Executor',                                    $false),
    @(79, 'Executor1',                                   $false),
    @(81, 'Profiling to Outreach Tag',                   $false),
    @(83, 'Tag per SDE',                                 $false),
    @(84, 'Tag 2 per SDE',                               $false),
    @(89, 'Outreach TAGS',                               $false),
    @(101,'Accounts enriched after complete Profiling',  $false),
    @(105,'Profiling CRM Outreach',                      $false),
    @(106,'Profiling CRM Nuevo',                         $false),
    @(107,'Profiling NO CRM Accounts',                   $false)
)

function JsonStr($s) {
    $s = $s.Trim()
    $s = $s.Replace('\', '\\')
    $s = $s.Replace('"', '\"')
    $cr = [string][char]13
    $lf = [string][char]10
    $s = $s.Replace($cr, '').Replace($lf, '')
    return '"' + $s + '"'
}

Write-Host 'Opening XLSX read-only...'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($xlsxPath, 0, $true)  # UpdateLinks=0, ReadOnly=$true
$ws = $wb.Sheets.Item(1)
Write-Host 'Opened.'

# Find last row in col 1 using Excel End(xlDown) — single COM call
$xlDown = -4121
$lastDataRow = [Math]::Min($ws.Cells.Item(2, 1).End($xlDown).Row, 10000)
$totalCols = 114
Write-Host ("Data rows: " + ($lastDataRow - 1) + ", cols: $totalCols")

# Bulk read entire data range in one COM call
Write-Host 'Bulk reading...'
$range = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($lastDataRow, $totalCols))
$allData = $range.Value2
Write-Host 'Bulk read done. Building JSON...'

$sw = [System.IO.StreamWriter]::new($outFile, $false, [System.Text.Encoding]::UTF8)
$sw.Write('window.CAMP_DATA=[')
$firstRecord = $true
$rowCount = 0

for ($row = 2; $row -le $lastDataRow; $row++) {
    $campName = $allData.GetValue($row, 1)
    if ($null -eq $campName -or [string]$campName -eq '') { continue }

    $sb = New-Object System.Text.StringBuilder
    $sb.Append('{') | Out-Null
    $firstField = $true

    foreach ($entry in $colMap) {
        $colIdx  = $entry[0]
        $outName = $entry[1]
        $isDate  = $entry[2]

        $cellVal = $allData.GetValue($row, $colIdx)
        $strVal  = ''
        if ($null -ne $cellVal) {
            if ($isDate -and $cellVal -is [double]) {
                try { $strVal = [DateTime]::FromOADate([double]$cellVal).ToString('yyyy-MM-dd') }
                catch { $strVal = [string]$cellVal }
            } else {
                $strVal = [string]$cellVal
            }
        }

        if (-not $firstField) { $sb.Append(',') | Out-Null }
        $firstField = $false
        $sb.Append('"' + $outName + '":' + (JsonStr $strVal)) | Out-Null
    }
    $sb.Append('}') | Out-Null

    if (-not $firstRecord) { $sw.Write(',') }
    $sw.Write($sb.ToString())
    $firstRecord = $false
    $rowCount++
}

$sw.Write('];')
$sw.Close()

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

$size = (Get-Item $outFile).Length
Write-Host ("Done: $rowCount records -> $outFile ($([int]($size/1KB)) KB)")
