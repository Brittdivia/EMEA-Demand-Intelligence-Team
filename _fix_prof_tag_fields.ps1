$htmlPath = "C:\Users\I572929\campaign-calendar-site\index.html"
$html = [System.IO.File]::ReadAllText($htmlPath, [System.Text.Encoding]::UTF8)
$dq = [char]34

$errors = 0
function DoSbReplace($label, [ref]$h, $old, $new) {
    if (-not $h.Value.Contains($old)) { Write-Host "FAIL: $label"; $script:errors++; return }
    $sb = [System.Text.StringBuilder]::new($h.Value)
    $sb.Replace($old, $new) | Out-Null
    $h.Value = $sb.ToString()
    Write-Host ("OK:   $label")
}

# 1. _profTagsByNormDcm: [m.tag,m.tagOut] -> [m.tag,m.tagOut,m.tagEnr,m.tagWave2]
DoSbReplace "_profTagsByNormDcm +tagEnr+tagWave2" ([ref]$html) `
    '[m.tag,m.tagOut].forEach(function(f){_splitDtrT(String(f||"")).forEach(function(t){_profTagsByNormDcm[key].add(t);});})' `
    '[m.tag,m.tagOut,m.tagEnr,m.tagWave2].forEach(function(f){_splitDtrT(String(f||"")).forEach(function(t){_profTagsByNormDcm[key].add(t);});})'

# 2. _profTagSetDcm: extend both _splitTags calls to include tagEnr and tagWave2
DoSbReplace "_profTagSetDcm +tagEnr+tagWave2" ([ref]$html) `
    ('profMetaVals.forEach(function(m){_splitTags(String(m.tagOut||' + $dq + $dq + ')).forEach(function(t){_profTagSetDcm.add(t);});_splitTags(String(m.tag||' + $dq + $dq + ')).forEach(function(t){_profTagSetDcm.add(t);});});') `
    ('profMetaVals.forEach(function(m){[m.tagOut,m.tag,m.tagEnr,m.tagWave2].forEach(function(f){_splitTags(String(f||' + $dq + $dq + ')).forEach(function(t){_profTagSetDcm.add(t);});});});')

# 3. _profTagsOut concat: add tagEnr and tagWave2 to the string concat
DoSbReplace "_profTagsOut +tagEnr+tagWave2" ([ref]$html) `
    ('var _profTagsOut=(String(m.tag||' + $dq + $dq + ')+' + $dq + '\n' + $dq + '+String(m.tagOut||' + $dq + $dq + ')).split(') `
    ('var _profTagsOut=(String(m.tag||' + $dq + $dq + ')+' + $dq + '\n' + $dq + '+String(m.tagOut||' + $dq + $dq + ')+' + $dq + '\n' + $dq + '+String(m.tagEnr||' + $dq + $dq + ')+' + $dq + '\n' + $dq + '+String(m.tagWave2||' + $dq + $dq + ')).split(')

if ($errors -gt 0) { Write-Host "ABORTED: $errors errors"; exit 1 }
[System.IO.File]::WriteAllText($htmlPath, $html, [System.Text.Encoding]::UTF8)
Write-Host ("Done. Size: " + (Get-Item $htmlPath).Length)

$h2 = [System.IO.File]::ReadAllText($htmlPath, [System.Text.Encoding]::UTF8)
Write-Host ("_profTagsByNormDcm tagEnr: " + $h2.Contains("m.tag,m.tagOut,m.tagEnr,m.tagWave2"))
Write-Host ("_profTagSetDcm tagEnr:     " + $h2.Contains("[m.tagOut,m.tag,m.tagEnr,m.tagWave2]"))
Write-Host ("_profTagsOut tagEnr:       " + $h2.Contains("String(m.tagEnr||"))
