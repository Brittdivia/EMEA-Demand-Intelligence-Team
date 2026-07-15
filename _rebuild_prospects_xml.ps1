# Rebuild data-prospects.js from Prospects_Export.xlsx using XML (no Excel COM)
Add-Type -AssemblyName System.IO.Compression.FileSystem

$xlsxPath = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Outreach\Prospects_Export.xlsx"
$zipPath = $xlsxPath + ".zip"
Copy-Item $xlsxPath $zipPath -Force

$zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)

# Read shared strings
$ssEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/sharedStrings.xml" }
$ssXml = [System.Xml.XmlDocument]::new()
$ssXml.Load($ssEntry.Open())
$ssNs = [System.Xml.XmlNamespaceManager]::new($ssXml.NameTable)
$ssNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
$strings = $ssXml.SelectNodes("//x:si", $ssNs) | ForEach-Object {
    $t = $_.SelectNodes(".//x:t", $ssNs) | ForEach-Object { $_.InnerText }
    $t -join ""
}
Write-Host "Shared strings: $($strings.Count)"

# Read sheet1
$shEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/worksheets/sheet1.xml" }
$shXml = [System.Xml.XmlDocument]::new()
$shXml.Load($shEntry.Open())
$shNs = [System.Xml.XmlNamespaceManager]::new($shXml.NameTable)
$shNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

$rows = $shXml.SelectNodes("//x:row", $shNs)
Write-Host "Rows: $($rows.Count)"

# Get headers
$colIdx = @{}
foreach ($cell in $rows[0].SelectNodes("x:c", $shNs)) {
    $ref = $cell.GetAttribute("r") -replace '[0-9]',''
    $t = $cell.GetAttribute("t")
    $v = $cell.SelectSingleNode("x:v", $shNs)
    if ($v -and $t -eq "s") { $colIdx[$strings[[int]$v.InnerText]] = $ref }
}
Write-Host "Headers: $($colIdx.Keys -join ' | ')"

# Map column letters to field names — only fields actually used in index.html
$needed = @{
    "Id"="id"; "Company"="co"; "Touched At"="touched"; "Stage Changed At"="sc";
    "Created At"="created"; "Tags"="tg"; "Stage Name"="stage"; "Email"="email";
    "Assigned Users"="au"; "Persona Name"="pn"; "Active Sequences"="as"; "Finished Sequences"="fs"
}
$colToField = @{}
foreach ($h in $needed.Keys) { if ($colIdx[$h]) { $colToField[$colIdx[$h]] = $needed[$h] } }
Write-Host "Mapped fields: $($colToField.Count)"

function EscJS($s) { if (!$s) { return "" } return $s.Replace('\','\\').Replace('"','\"').Replace("`r","").Replace("`n"," ") }

$entries = [System.Collections.Generic.List[string]]::new()
for ($i = 1; $i -lt $rows.Count; $i++) {
    $vals = @{}
    foreach ($cell in $rows[$i].SelectNodes("x:c", $shNs)) {
        $ref = $cell.GetAttribute("r") -replace '[0-9]',''
        if (-not $colToField[$ref]) { continue }
        $t = $cell.GetAttribute("t")
        $v = $cell.SelectSingleNode("x:v", $shNs)
        if ($v) { $vals[$colToField[$ref]] = if ($t -eq "s") { $strings[[int]$v.InnerText] } else { $v.InnerText } }
    }
    if (-not $vals["id"]) { continue }
    $parts = @()
    foreach ($f in @("id","co","touched","sc","created","tg","stage","email","au","pn","as","fs")) {
        $v = EscJS ($vals[$f])
        $parts += """$f"":""$v"""
    }
    $entries.Add("{" + ($parts -join ",") + "}")
    if ($i % 10000 -eq 0) { Write-Host "Processed $i / $($rows.Count - 1)" }
}

$zip.Dispose()
Remove-Item $zipPath -Force

$js = "window.PROSPECT_DATA=[" + ($entries -join ",") + "];"
[System.IO.File]::WriteAllText("C:\Users\I572929\campaign-calendar-site\data-prospects.js", $js, [System.Text.Encoding]::UTF8)
Write-Host "Done: $($entries.Count) prospects, $([Math]::Round($js.Length/1KB))KB"
