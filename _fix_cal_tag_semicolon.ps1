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

# 1+2: rawTags tagList split — add semicolon (appears twice)
DoSbReplace "rawTags split +semicolon (x2)" ([ref]$html) `
    'rawTags.split(/[\r\n,]+/)' `
    'rawTags.split(/[\r\n,;]+/)'

# 3: _calTagsByNormDcm
DoSbReplace "_calTagsByNormDcm +semicolon" ([ref]$html) `
    ('String(r[f]||' + $dq + $dq + ').split(' + $dq + ',' + $dq + ').forEach(function(t){t=t.trim();if(t)_calTagsByNormDcm[key].add(t.toLowerCase());})') `
    ('String(r[f]||' + $dq + $dq + ').split(/[,;]/).forEach(function(t){t=t.trim();if(t)_calTagsByNormDcm[key].add(t.toLowerCase());})')

# 4: dcmTagMap
DoSbReplace "dcmTagMap +semicolon" ([ref]$html) `
    ('String(r[c]||' + $dq + $dq + ').split(' + $dq + ',' + $dq + ').forEach(function(t){t=t.trim().toLowerCase();if(!t)return;') `
    ('String(r[c]||' + $dq + $dq + ').split(/[,;]/).forEach(function(t){t=t.trim().toLowerCase();if(!t)return;')

if ($errors -gt 0) { Write-Host "ABORTED: $errors errors"; exit 1 }
[System.IO.File]::WriteAllText($htmlPath, $html, [System.Text.Encoding]::UTF8)
Write-Host ("Done. Size: " + (Get-Item $htmlPath).Length)

# Verify
$h2 = [System.IO.File]::ReadAllText($htmlPath, [System.Text.Encoding]::UTF8)
Write-Host ("rawTags semicolon: " + $h2.Contains('rawTags.split(/[\r\n,;]+/)'))
Write-Host ("_calTagsByNormDcm semicolon: " + $h2.Contains('split(/[,;]/).forEach(function(t){t=t.trim();if(t)_calTagsByNormDcm'))
Write-Host ("dcmTagMap semicolon: " + $h2.Contains('split(/[,;]/).forEach(function(t){t=t.trim().toLowerCase()'))
