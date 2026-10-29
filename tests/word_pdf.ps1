# Opens .docx files in Microsoft Word (hidden, read-only) and exports each one as a PDF next to it (name.docx → name.word.pdf).
# Used to check that "PDF → Word" output looks like the original PDF when Word lays it out.
# Usage: powershell -File tests\word_pdf.ps1 file1.docx [file2.docx …]
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Files)
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false; $word.DisplayAlerts = 0
try {
  foreach ($f in $Files) {
    $p = (Resolve-Path $f).Path; $out = [IO.Path]::ChangeExtension($p, '.word.pdf')
    try {
      $d = $word.Documents.Open($p, $false, $true, $false, '', '', $false, '', '', 0, [Type]::Missing, $false, $false, 0, $true)
      $pages = $d.ComputeStatistics(2)
      $d.ExportAsFixedFormat($out, 17) # wdExportFormatPDF
      $d.Close(0)
      [ordered]@{ file = (Split-Path $p -Leaf); ok = $true; pages = $pages; pdf = $out } | ConvertTo-Json -Compress
    } catch { [ordered]@{ file = (Split-Path $p -Leaf); ok = $false; error = $_.Exception.Message } | ConvertTo-Json -Compress }
  }
} finally { $word.Quit() }
