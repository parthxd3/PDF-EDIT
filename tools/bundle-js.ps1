# Joins the editor's own scripts (in the order index.html loads them) into js/bundle.js and points index.html at it.
# Used by "Build Hostinger ZIP.bat" on the staged copy, so the hosted page needs one script request instead of twelve.
param([Parameter(Mandatory = $true)][string]$Stage)
$ErrorActionPreference = 'Stop'
$utf8 = New-Object Text.UTF8Encoding $false
$indexPath = Join-Path $Stage 'index.html'
$html = [IO.File]::ReadAllText($indexPath)
$tag = '<script src="js/([^"]+)"></script>\r?\n?'
$files = @([regex]::Matches($html, $tag) | ForEach-Object { $_.Groups[1].Value })
if ($files.Count -lt 2) { Write-Host '  (scripts already bundled)'; exit 0 }
$parts = $files | ForEach-Object { [IO.File]::ReadAllText((Join-Path $Stage ('js/' + $_))) }
[IO.File]::WriteAllText((Join-Path $Stage 'js/bundle.js'), ($parts -join "`n"), $utf8)
$first = $true
$html = [regex]::Replace($html, $tag, { param($m) if ($script:first) { $script:first = $false; '<script src="js/bundle.js"></script>' + "`n" } else { '' } })
[IO.File]::WriteAllText($indexPath, $html, $utf8)
$files | ForEach-Object { Remove-Item (Join-Path $Stage ('js/' + $_)) }
Write-Host ("  Joined {0} scripts into js/bundle.js" -f $files.Count)
