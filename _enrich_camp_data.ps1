# Enrich data-camp.js: take the larger inline d-camp as base,
# then patch in DG Program and Digital Assets team support from the fresh Excel build
$site = "C:\Users\I572929\campaign-calendar-site"

# 1. Extract inline d-camp from index.html
Write-Host "Reading inline d-camp from index.html..."
$html = [System.IO.File]::ReadAllText("$site\index.html", [System.Text.Encoding]::UTF8)
$startTag = '<script type="application/json" id="d-camp">'
$endTag = '</script>'
$startIdx = $html.IndexOf($startTag) + $startTag.Length
$endIdx = $html.IndexOf($endTag, $startIdx)
$inlineJson = $html.Substring($startIdx, $endIdx - $startIdx)

# 2. Read current data-camp.js (fresh Excel rebuild with DG Program)
Write-Host "Reading data-camp.js (Excel rebuild)..."
$freshJs = [System.IO.File]::ReadAllText("$site\data-camp.js", [System.Text.Encoding]::UTF8)
$freshJson = $freshJs.Substring("window.CAMP_DATA=".Length).TrimEnd(';')

# 3. Parse both - build WBS->DG Program and WBS->DAT lookup from fresh
Write-Host "Building enrichment lookup..."
$freshLines = $freshJson.TrimStart('[').TrimEnd(']') -split '},\{'
$dgMap = @{}
$datMap = @{}
$exeMap = @{}
foreach ($line in $freshLines) {
    $line = $line.TrimStart('{').TrimEnd('}')
    $wbsMatch = [regex]::Match($line, '"Campaign/WBS Code"\s*:\s*"([^"]*)"')
    $dgMatch  = [regex]::Match($line, '"DG Program"\s*:\s*"([^"]*)"')
    $datMatch = [regex]::Match($line, '"Digital Assets team support"\s*:\s*"([^"]*)"')
    $exeMatch = [regex]::Match($line, '"Executor"\s*:\s*"([^"]*)"')
    if ($wbsMatch.Success) {
        $wbs = $wbsMatch.Groups[1].Value
        if ($dgMatch.Success -and $dgMatch.Groups[1].Value)  { $dgMap[$wbs]  = $dgMatch.Groups[1].Value }
        if ($datMatch.Success -and $datMatch.Groups[1].Value) { $datMap[$wbs] = $datMatch.Groups[1].Value }
        if ($exeMatch.Success -and $exeMatch.Groups[1].Value) { $exeMap[$wbs] = $exeMatch.Groups[1].Value }
    }
}
Write-Host "DG mappings: $($dgMap.Count)  DAT mappings: $($datMap.Count)  Executor mappings: $($exeMap.Count)"

# 4. Inject into inline JSON using regex replacement per record
Write-Host "Patching inline d-camp records..."
$patched = [regex]::Replace($inlineJson, '\{[^{}]+\}', {
    param($m)
    $rec = $m.Value
    $wbsM = [regex]::Match($rec, '"Campaign/WBS Code"\s*:\s*"([^"]*)"')
    if (-not $wbsM.Success) { return $rec }
    $wbs = $wbsM.Groups[1].Value
    # Remove existing DG Program / DAT / Executor fields if present
    $rec = [regex]::Replace($rec, ',?"DG Program"\s*:\s*(?:"[^"]*"|null)', '')
    $rec = [regex]::Replace($rec, ',?"Digital Assets team support"\s*:\s*(?:"[^"]*"|null)', '')
    $rec = [regex]::Replace($rec, ',?"Executor"\s*:\s*(?:"[^"]*"|null)', '')
    # Inject before closing brace
    $dg  = if ($dgMap[$wbs])   { $dgMap[$wbs].Replace('\','\\').Replace('"','\"')   } else { "" }
    $dat = if ($datMap[$wbs])  { $datMap[$wbs].Replace('\','\\').Replace('"','\"')  } else { "" }
    $exe = if ($exeMap[$wbs])  { $exeMap[$wbs].Replace('\','\\').Replace('"','\"')  } else { "" }
    $rec.TrimEnd('}') + ',"DG Program":"' + $dg + '","Digital Assets team support":"' + $dat + '","Executor":"' + $exe + '"}'
})

$js = "window.CAMP_DATA=$patched;"
[System.IO.File]::WriteAllText("$site\data-camp.js", $js, [System.Text.Encoding]::UTF8)
$kb = [Math]::Round((Get-Item "$site\data-camp.js").Length / 1KB)
Write-Host "data-camp.js updated - ${kb}KB"

# Spot check
$dgCount = ([regex]::Matches($patched, '"DG Program":"[^"]+')).Count
$datCount = ([regex]::Matches($patched, '"Digital Assets team support":"[^"]+')).Count
Write-Host "Records with DG Program value: $dgCount"
Write-Host "Records with DAT value: $datCount"
