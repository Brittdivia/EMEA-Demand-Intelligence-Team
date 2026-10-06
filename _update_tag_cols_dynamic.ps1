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

# Change 1+2: rawTags built from hardcoded 3-column string concat (appears twice — replaced both)
$oldRaw = ('var rawTags=(String(r[' + $dq + 'Profiling to Outreach Tag' + $dq + ']||' + $dq + $dq + ').trim()+' + $dq + '\n' + $dq + '+String(r[' + $dq + 'Tag per SDE' + $dq + ']||' + $dq + $dq + ').trim()+' + $dq + '\n' + $dq + '+String(r[' + $dq + 'Tag 2 per SDE' + $dq + ']||' + $dq + $dq + ').trim()).trim();if(!rawTags)return;')
$newRaw = ('var _tCols=Object.keys(r).filter(function(c){return c.toLowerCase().indexOf(' + $dq + 'tag' + $dq + ')>=0;});var rawTags=_tCols.map(function(c){return String(r[c]||' + $dq + $dq + ').trim();}).filter(Boolean).join(' + $dq + '\n' + $dq + ').trim();if(!rawTags)return;')
DoSbReplace "rawTags dynamic (x2)" ([ref]$html) $oldRaw $newRaw

# Change 3: _calTagsByNormDcm — hardcoded 4-column array
$oldDcm = ('[' + $dq + 'Profiling to Outreach Tag' + $dq + ',' + $dq + 'Tag per SDE' + $dq + ',' + $dq + 'Tag 2 per SDE' + $dq + ',' + $dq + 'Outreach TAGS' + $dq + '].forEach(function(f){_splitDtrT(String(r[f]||' + $dq + $dq + ')).forEach(function(t){_calTagsByNormDcm[key].add(t);});});')
$newDcm = ('Object.keys(r).filter(function(f){return f.toLowerCase().indexOf(' + $dq + 'tag' + $dq + ')>=0;}).forEach(function(f){_splitDtrT(String(r[f]||' + $dq + $dq + ')).forEach(function(t){_calTagsByNormDcm[key].add(t);});});')
DoSbReplace "_calTagsByNormDcm dynamic" ([ref]$html) $oldDcm $newDcm

# Change 4: dcmTagMap — hardcoded 3-column array
$oldMap = ('[String(r[' + $dq + 'Profiling to Outreach Tag' + $dq + ']||' + $dq + $dq + '),String(r[' + $dq + 'Tag per SDE' + $dq + ']||' + $dq + $dq + '),String(r[' + $dq + 'Tag 2 per SDE' + $dq + ']||' + $dq + $dq + ')].forEach(function(raw){')
$newMap = ('Object.keys(r).filter(function(c){return c.toLowerCase().indexOf(' + $dq + 'tag' + $dq + ')>=0;}).map(function(c){return String(r[c]||' + $dq + $dq + ');}).forEach(function(raw){')
DoSbReplace "dcmTagMap dynamic" ([ref]$html) $oldMap $newMap

if ($errors -gt 0) { Write-Host "ABORTED: $errors errors"; exit 1 }

[System.IO.File]::WriteAllText($htmlPath, $html, [System.Text.Encoding]::UTF8)
Write-Host ("Done. Size: " + (Get-Item $htmlPath).Length)

# Verify
$h2 = [System.IO.File]::ReadAllText($htmlPath, [System.Text.Encoding]::UTF8)
$remaining = ($h2 | Select-String -Pattern 'Profiling to Outreach Tag' -AllMatches).Matches.Count
Write-Host ("Remaining hardcoded 'Profiling to Outreach Tag' refs: $remaining")
$dynCount = ($h2 | Select-String -Pattern 'indexOf\("tag"\)>=0' -AllMatches).Matches.Count
Write-Host ("Dynamic 'indexOf(tag)' occurrences: $dynCount")
