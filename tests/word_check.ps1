# Opens .docx files in Microsoft Word (hidden, read-only) and reports whether Word reads them cleanly — used to check
# files saved by the Word editor. Word is closed again afterwards; a file Word would need to "repair" is reported as such.
# Usage: powershell -File tests\word_check.ps1 file1.docx [file2.docx …]
param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Files)
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false; $word.DisplayAlerts = 0
try {
  foreach ($f in $Files) {
    $p = (Resolve-Path $f).Path
    try {
      # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog)
      $d = $word.Documents.Open($p, $false, $true, $false, '', '', $false, '', '', 0, [Type]::Missing, $false, $false, 0, $true)
      $o = [ordered]@{ file = (Split-Path $p -Leaf); ok = $true; paragraphs = $d.Paragraphs.Count; tables = $d.Tables.Count; pictures = $d.InlineShapes.Count; shapes = $d.Shapes.Count
        footnotes = $d.Footnotes.Count; comments = $d.Comments.Count; fields = $d.Fields.Count; sections = $d.Sections.Count; lists = $d.Lists.Count; pages = $d.ComputeStatistics(2)
        header = $d.Sections.Item(1).Headers.Item(1).Range.Text.Trim(); text = $d.Content.Text.Substring(0, [Math]::Min(160, $d.Content.Text.Length)) }
      $d.Close(0)
    } catch { $o = [ordered]@{ file = (Split-Path $p -Leaf); ok = $false; error = $_.Exception.Message } }
    $o | ConvertTo-Json -Compress
  }
} finally { $word.Quit() }
