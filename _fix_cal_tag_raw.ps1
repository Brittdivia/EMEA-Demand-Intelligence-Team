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

# Change 1: _calTagsByNormDcm — replace _splitDtrT (strips prefixes) with comma-split + lowercase only
DoSbReplace "_calTagsByNormDcm raw" ([ref]$html) `
    ('Object.keys(r).filter(function(f){return f.toLowerCase().indexOf(' + $dq + 'tag' + $dq + ')>=0;}).forEach(function(f){_splitDtrT(String(r[f]||' + $dq + $dq + ')).forEach(function(t){_calTagsByNormDcm[key].add(t);});});') `
    ('Object.keys(r).filter(function(f){return f.toLowerCase().indexOf(' + $dq + 'tag' + $dq + ')>=0;}).forEach(function(f){String(r[f]||' + $dq + $dq + ').split(' + $dq + ',' + $dq + ').forEach(function(t){t=t.trim();if(t)_calTagsByNormDcm[key].add(t.toLowerCase());});});')

# Change 2: dcmTagMap — _splitTags on calendar raw → comma-split + lowercase only
DoSbReplace "dcmTagMap raw" ([ref]$html) `
    ('Object.keys(r).filter(function(c){return c.toLowerCase().indexOf(' + $dq + 'tag' + $dq + ')>=0;}).map(function(c){return String(r[c]||' + $dq + $dq + ');}).forEach(function(raw){' + [char]13 + [char]10 + '        _splitTags(raw).forEach(function(t){') `
    ('Object.keys(r).filter(function(c){return c.toLowerCase().indexOf(' + $dq + 'tag' + $dq + ')>=0;}).forEach(function(c){String(r[c]||' + $dq + $dq + ').split(' + $dq + ',' + $dq + ').forEach(function(t){t=t.trim().toLowerCase();if(!t)return;')

# Change 3: _calTagToCampCache tagList — remove prefix stripping from inline .replace()
DoSbReplace "_calTagToCampCache raw" ([ref]$html) `
    ("tagList=rawTags.split('\\r\\n').join('\n').split('\\n').join('\n').split(/[\n,]+/).map(function(t){return t.trim().replace(/^(?:New Prospects\s+Tag|Existing\s+Tag|New\s+Tag|Tag|Existing|New):\s*/i," + $dq + $dq + ").trim();}).filter(function(t){return t&&t.length>1;});") `
    ("tagList=rawTags.split(/[\r\n,]+/).map(function(t){return t.trim();}).filter(function(t){return t&&t.length>1;});")

# Change 4: _wbsWithCalTag tagList — remove prefix/suffix stripping
DoSbReplace "_wbsWithCalTag raw" ([ref]$html) `
    ('rawTags.split(/[\r\n,;|]+|\\r\\n|\\n/).map(function(t){return t.trim().replace(/^(?:New Prospects\s+Tag|Existing\s+Tag|New\s+Tag|Tag|Existing|New)(?::\s*|\s+)/i,' + $dq + $dq + ').replace(/^(?:EX|NP)-/i,' + $dq + $dq + ').replace(/[\s-]+(?:existing|new|ex|np)$/i,' + $dq + $dq + ').trim();}).filter(function(t){return t&&t.length>1;})') `
    ('rawTags.split(/[\r\n,]+/).map(function(t){return t.trim();}).filter(function(t){return t&&t.length>1;})')

if ($errors -gt 0) { Write-Host "ABORTED: $errors errors"; exit 1 }

[System.IO.File]::WriteAllText($htmlPath, $html, [System.Text.Encoding]::UTF8)
Write-Host ("Done. Size: " + (Get-Item $htmlPath).Length)

# Verify no prefix-stripping regexes remain on calendar paths
$h2 = [System.IO.File]::ReadAllText($htmlPath, [System.Text.Encoding]::UTF8)
$remaining = ($h2 | Select-String -Pattern 'New Prospects.*Tag' -AllMatches).Matches.Count
Write-Host ("Remaining 'New Prospects Tag' cleaning refs: $remaining")
