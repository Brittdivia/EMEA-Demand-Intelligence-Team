# Fast version - reads entire range at once instead of cell by cell
$seqStats = Import-Csv "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Outreach\Sequence_Stats_2026-06-02.csv"
$nameToId = @{}
foreach ($row in $seqStats) {
    $sid = $row."Sequence ID".Trim()
    $sname = $row."Sequence Name".Trim()
    if ($sid -and $sname -and -not $nameToId.ContainsKey($sname)) { $nameToId[$sname] = $sid }
}
Write-Host "Sequence name->ID: $($nameToId.Count)"

$outJson = Get-Content "C:\Users\I572929\campaign-calendar-site\outreach.json" -Raw | ConvertFrom-Json
$trackedSids = New-Object System.Collections.Generic.HashSet[string]
$outJson | ForEach-Object { [void]$trackedSids.Add("$($_.'Sequence ID')".Trim()) }
Write-Host "Tracked sequences: $($trackedSids.Count)"

$csvPath = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\Profiling request.csv"
Write-Host "Reading CSV..."
$csvData = Import-Csv -Path $csvPath -Encoding UTF8
$lastRow = $csvData.Count
Write-Host "Rows: $lastRow"

$tagToProf = @{}
$campCodeToProf = @{}
$allCreatedBy = New-Object System.Collections.Generic.HashSet[string]
$allProfMeta = @{}

foreach ($row in $csvData) {
    $profId    = $row.ID.Trim()
    if (-not $profId) { continue }
    $title     = $row.Title.Trim().Replace('\','\\').Replace('"','\"')
    $ddm1      = $row.DDM1.Trim().Replace('"','\"')
    $reqType   = $row.'Request Type'.Trim().Replace('"','\"')
    $status    = $row.Status.Trim().Replace('"','\"')
    $createdBy = $row.'Created By'.Trim().Replace('"','\"')
    $campCode  = $row.'Campaign Code'.Trim().Replace('"','\"')
    $tagPros   = $row.'Tag of Prospects'.Trim().Trim('"')
    $tagOut    = $row.'Tag for Outreach'.Trim().Trim('"')
    $tagEnr    = $row.'Tag of Enriched Accounts'.Trim().Trim('"')
    $tagWave2  = $row.'Tag of Prospects (Wave2)'.Trim().Trim('"')

    if ($createdBy) { [void]$allCreatedBy.Add($createdBy) }

    $allProfMeta[$profId] = @{id=$profId;title=$title;ddm1=$ddm1;type=$reqType;status=$status;tag=$tagPros;tagOut=$tagOut;tagEnr=$tagEnr;tagWave2=$tagWave2;wbs="";createdBy=$createdBy;campaignCode=$campCode;created=""}

    if ($campCode) {
        $campCode -split '[,;\n]' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | ForEach-Object {
            if (-not $campCodeToProf[$_]) { $campCodeToProf[$_] = @{id=$profId;title=$title;ddm1=$ddm1;createdBy=$createdBy} }
        }
    }

    $allTags = @()
    if ($tagPros)  { $allTags += $tagPros  -split '[,;\n]' | ForEach-Object { $_.Trim().Trim('"') } | Where-Object { $_ } }
    if ($tagOut)   { $allTags += $tagOut   -split '[,;\n]' | ForEach-Object { $_.Trim().Trim('"') } | Where-Object { $_ } }
    if ($tagEnr)   { $allTags += $tagEnr   -split '[,;\n]' | ForEach-Object { $_.Trim().Trim('"') } | Where-Object { $_ } }
    if ($tagWave2) { $allTags += $tagWave2 -split '[,;\n]' | ForEach-Object { $_.Trim().Trim('"') } | Where-Object { $_ } }

    foreach ($tag in $allTags) {
        if (-not $tagToProf[$tag]) { $tagToProf[$tag] = [System.Collections.Generic.List[object]]::new() }
        $tagToProf[$tag].Add(@{id=$profId;title=$title;ddm1=$ddm1;type=$reqType;status=$status;tag=$tagPros;tagOut=$tagOut;createdBy=$createdBy;campaignCode=$campCode})
    }
}
Write-Host "Profiling tags indexed: $($tagToProf.Count)"

# Match sequences to profiling via WBS — update wbs on allProfMeta entries
$campProfTags = @{}
$matchCount = 0
foreach ($sid in $trackedSids) {
    $wbs = ($outJson | Where-Object { "$($_.'Sequence ID')".Trim() -eq $sid } | Select-Object -First 1)."Campaign/WBS Code"
    if (-not $wbs) { continue }
    $statsRow = $seqStats | Where-Object { $_."Sequence ID".Trim() -eq $sid } | Select-Object -First 1
    if (-not $statsRow) { continue }
    $tags = $statsRow.Tags.Trim().Trim('"') -split '[,\n]' | ForEach-Object { $_.Trim().Trim('"') } | Where-Object { $_ }
    foreach ($tag in $tags) {
        if ($tagToProf[$tag]) {
            if (-not $campProfTags[$wbs]) { $campProfTags[$wbs] = $tag }
            foreach ($prof in $tagToProf[$tag]) {
                if ($allProfMeta[$prof.id] -and -not $allProfMeta[$prof.id]["wbs"]) {
                    $allProfMeta[$prof.id]["wbs"] = $wbs; $matchCount++
                }
            }
        }
    }
}
Write-Host "Total matches: $matchCount"

$tagsJson = ($campProfTags.GetEnumerator() | ForEach-Object { """$($_.Key -replace '"','\"')"":""$($_.Value -replace '"','\"')""" }) -join ","
$metaEntries = $allProfMeta.GetEnumerator() | ForEach-Object {
    $k = $_.Key -replace '"','\"'; $v = $_.Value
    """$k"":{""id"":""$($v.id)"",""title"":""$($v.title)"",""ddm1"":""$($v.ddm1)"",""type"":""$($v.type)"",""status"":""$($v.status)"",""tag"":""$($v.tag -replace '"','\"')"",""tagOut"":""$($v.tagOut -replace '"','\"')"",""tagEnr"":""$($v.tagEnr -replace '"','\"')"",""tagWave2"":""$($v.tagWave2 -replace '"','\"')"",""wbs"":""$($v.wbs -replace '"','\"')"",""createdBy"":""$($v.createdBy -replace '"','\"')"",""campaignCode"":""$($v.campaignCode -replace '"','\"')""}"
}
$metaJson = $metaEntries -join ","
$codesEntries = $campCodeToProf.GetEnumerator() | ForEach-Object { """$($_.Key -replace '"','\"')"": {""id"":""$($_.Value.id)"",""title"":""$($_.Value.title)"",""ddm1"":""$($_.Value.ddm1 -replace '"','\"')"",""createdBy"":""$($_.Value.createdBy -replace '"','\"')""}" }
$codesJson = $codesEntries -join ","
$creatorsJson = ($allCreatedBy | ForEach-Object { """$($_ -replace '"','\"')""" }) -join ","

$js = "window.CAMP_PROF_TAGS={$tagsJson};`nwindow.CAMP_PROF_META={$metaJson};`nwindow.CAMP_PROF_CODES={$codesJson};`nwindow.CAMP_PROF_CREATORS=[$creatorsJson];"
[System.IO.File]::WriteAllText("C:\Users\I572929\campaign-calendar-site\data-profiling-tags.js", $js, [System.Text.Encoding]::UTF8)
Write-Host "Done: $([Math]::Round($js.Length/1KB))KB"
