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

# _calTagsByNormDcm: split(/[,;]/) -> split(/[,;\r\n]+/)
DoSbReplace "_calTagsByNormDcm +newline" ([ref]$html) `
    ('String(r[f]||' + $dq + $dq + ').split(/[,;]/).forEach(function(t){t=t.trim();if(t)_calTagsByNormDcm[key].add(t.toLowerCase());})') `
    ('String(r[f]||' + $dq + $dq + ').split(/[,;\r\n]+/).forEach(function(t){t=t.trim();if(t)_calTagsByNormDcm[key].add(t.toLowerCase());})')

# dcmTagMap: split(/[,;]/) -> split(/[,;\r\n]+/)
DoSbReplace "dcmTagMap +newline" ([ref]$html) `
    ('String(r[c]||' + $dq + $dq + ').split(/[,;]/).forEach(function(t){t=t.trim().toLowerCase();if(!t)return;') `
    ('String(r[c]||' + $dq + $dq + ').split(/[,;\r\n]+/).forEach(function(t){t=t.trim().toLowerCase();if(!t)return;')

if ($errors -gt 0) { Write-Host "ABORTED: $errors errors"; exit 1 }
[System.IO.File]::WriteAllText($htmlPath, $html, [System.Text.Encoding]::UTF8)
Write-Host ("Done. Size: " + (Get-Item $htmlPath).Length)
$h2 = [System.IO.File]::ReadAllText($htmlPath, [System.Text.Encoding]::UTF8)
Write-Host ("_calTagsByNormDcm newline: " + $h2.Contains('split(/[,;\r\n]+/).forEach(function(t){t=t.trim();if(t)_calTagsByNormDcm'))
Write-Host ("dcmTagMap newline:         " + $h2.Contains('split(/[,;\r\n]+/).forEach(function(t){t=t.trim().toLowerCase()'))
