# Turns the packaged app folders (dist\XD3 PDF Editor-win32-<arch>) into files people can download:
#   XD3-PDF-Editor-Setup-<version>-<arch>.exe   installer (Start menu + desktop shortcuts, "Open with" for PDFs, uninstaller)
#   XD3-PDF-Editor-Portable-<version>-<arch>.zip  no installation: unzip and run "XD3 PDF Editor.exe"
#   manifest.json                                 what the website's "Download for Windows" dialog lists
# Used by "Build EXE.bat". Needs Node.js (electron-builder is fetched with npx the first time).
param([Parameter(Mandatory = $true)][string]$Out, [string]$Arches = 'x64 ia32')
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$dl = Join-Path $Out 'downloads'
New-Item -ItemType Directory -Force $dl | Out-Null
$version = (Get-Content -Raw (Join-Path $root 'desktop\package.json') | ConvertFrom-Json).version
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$files = @()
foreach ($arch in $Arches.Split(' ', [StringSplitOptions]::RemoveEmptyEntries)) {
  $dir = Join-Path $Out "XD3 PDF Editor-win32-$arch"
  if (-not (Test-Path (Join-Path $dir 'XD3 PDF Editor.exe'))) { Write-Host "  (no $arch build found - skipped)"; continue }

  Write-Host "  - $arch installer"
  Get-ChildItem $dl -Filter "XD3-PDF-Editor-Setup-*-$arch.exe*" -ErrorAction SilentlyContinue | Remove-Item -Force
  & npx.cmd --yes electron-builder@24.13.3 --win nsis "--$arch" --prepackaged "$dir" --projectDir (Join-Path $root 'desktop') --config builder.json --publish never
  if ($LASTEXITCODE -ne 0) { throw "electron-builder failed for $arch" }
  $setup = Get-ChildItem (Join-Path $root 'dist\downloads') -Filter "XD3-PDF-Editor-Setup-$version-$arch.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $setup) { $setup = Get-ChildItem (Join-Path $root 'desktop\dist') -Filter "XD3-PDF-Editor-Setup-$version-$arch.exe" -ErrorAction SilentlyContinue | Select-Object -First 1 }
  if ($setup -and $setup.DirectoryName -ne (Resolve-Path $dl).Path) { Move-Item -Force $setup.FullName $dl; $setup = Get-Item (Join-Path $dl $setup.Name) }
  if ($setup) { $files += [ordered]@{ file = $setup.Name; kind = 'installer'; arch = $arch; size = $setup.Length } }

  Write-Host "  - $arch portable ZIP"
  $zip = Join-Path $dl "XD3-PDF-Editor-Portable-$version-$arch.zip"
  Get-ChildItem $dl -Filter "XD3-PDF-Editor-Portable-*-$arch.zip" -ErrorAction SilentlyContinue | Remove-Item -Force # (older versions too)
  $z = [IO.Compression.ZipFile]::Open($zip, 'Create')
  try {
    $top = "XD3 PDF Editor"
    Get-ChildItem -LiteralPath $dir -Recurse -File | ForEach-Object { [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($z, $_.FullName, "$top/" + $_.FullName.Substring($dir.Length + 1).Replace('\', '/'), 'Optimal') }
  } finally { $z.Dispose() }
  $files += [ordered]@{ file = (Split-Path $zip -Leaf); kind = 'portable'; arch = $arch; size = (Get-Item $zip).Length }
}
# (electron-builder leaves helper files next to the installer)
Get-ChildItem $dl -Exclude *.exe, *.zip, manifest.json | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
if (Test-Path (Join-Path $root 'desktop\dist')) { Remove-Item -Recurse -Force (Join-Path $root 'desktop\dist') }
$manifest = [ordered]@{ name = 'XD3 PDF Editor'; version = $version; built = (Get-Date -Format 'yyyy-MM-dd'); files = $files }
[IO.File]::WriteAllText((Join-Path $dl 'manifest.json'), ($manifest | ConvertTo-Json -Depth 4), (New-Object Text.UTF8Encoding $false))
Write-Host ("  Downloads ready in {0} ({1} file(s))" -f $dl, $files.Count)
