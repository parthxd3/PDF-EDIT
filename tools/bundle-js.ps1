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
# the service worker's CORE list (what it keeps for offline use) must name the files this build really has: js/bundle.js instead of the joined scripts
$swPath = Join-Path $Stage 'sw.js'
if (Test-Path $swPath) {
  $sw = [IO.File]::ReadAllText($swPath)
  $core = [regex]::Match($sw, 'const CORE = \[([^\]]*)\];')
  if (-not $core.Success) { throw 'sw.js: the CORE list was not found' }
  $joined = @{}; $files | ForEach-Object { $joined['js/' + $_] = $true }
  $list = New-Object Collections.Generic.List[string]; $dropped = @()
  foreach ($m in [regex]::Matches($core.Groups[1].Value, "'([^']*)'")) {
    $u = $m.Groups[1].Value
    if ($joined.ContainsKey($u)) { if (-not $list.Contains('js/bundle.js')) { $list.Add('js/bundle.js') }; continue }
    if ($u -ne './' -and -not (Test-Path -LiteralPath (Join-Path $Stage $u))) { $dropped += $u; continue }
    if (-not $list.Contains($u)) { $list.Add($u) }
  }
  if (-not $list.Contains('js/bundle.js')) { $list.Insert([Math]::Min(5, $list.Count), 'js/bundle.js') }
  $text = 'const CORE = [' + (($list | ForEach-Object { "'" + $_ + "'" }) -join ', ') + '];'
  $sw = $sw.Substring(0, $core.Index) + $text + $sw.Substring($core.Index + $core.Length)
  [IO.File]::WriteAllText($swPath, $sw, $utf8)
  Write-Host ("  sw.js keeps {0} core files offline{1}" -f $list.Count, $(if ($dropped.Count) { ' (left out, not in this build: ' + ($dropped -join ', ') + ')' } else { '' }))
}
