# Builds one realistic Word document through COM, exports it to PDF and saves the .docx next to it.
# Usage: powershell -File word.ps1 -Name resume -Out <dir> -Pics <dir> -Tmp <dir>
param([string]$Name, [string]$Out, [string]$Pics, [string]$Tmp)
$ErrorActionPreference = 'Stop'
$Out = [IO.Path]::GetFullPath($Out); $Pics = [IO.Path]::GetFullPath($Pics); $Tmp = [IO.Path]::GetFullPath($Tmp)
$M = [System.Reflection.Missing]::Value
$em = [string][char]0x2014; $nd = [string][char]0x2013; $bul = [string][char]0x2022; $pound = [string][char]0x00A3
$rsq = [string][char]0x2019; $ldq = [string][char]0x201C; $rdq = [string][char]0x201D; $eur = [string][char]0x20AC

function C($r, $g, $b) { [int]($r + 256 * $g + 65536 * $b) }

$word = New-Object -ComObject Word.Application
$word.Visible = $false; $word.DisplayAlerts = 0
$word.Options.AutoFormatAsYouTypeApplyNumberedLists = $false
$word.Options.AutoFormatAsYouTypeApplyBulletedLists = $false
$word.Options.AutoFormatAsYouTypeReplaceQuotes = $false
try {
  $doc = $word.Documents.Add()
  $sel = $word.Selection

  function GoEnd { $e = $doc.Content.End - 1; $doc.Range($e, $e).Select() }
  # Start a paragraph in a style, with optional direct formatting. Leaves the cursor inside it.
  function Para([int]$style = -1, [hashtable]$f = @{}) {
    $sel.Style = $style
    $sel.ParagraphFormat.Reset(); $sel.Font.Reset()
    $pf = $sel.ParagraphFormat
    if ($f.ContainsKey('align')) { $pf.Alignment = $f.align }
    if ($f.ContainsKey('before')) { $pf.SpaceBefore = $f.before }
    if ($f.ContainsKey('after')) { $pf.SpaceAfter = $f.after }
    if ($f.ContainsKey('line')) { $pf.LineSpacingRule = 5; $pf.LineSpacing = 12 * $f.line } # wdLineSpaceMultiple
    if ($f.ContainsKey('left')) { $pf.LeftIndent = $f.left }
    if ($f.ContainsKey('right')) { $pf.RightIndent = $f.right }
    if ($f.ContainsKey('first')) { $pf.FirstLineIndent = $f.first }
    if ($f.ContainsKey('shade')) { $pf.Shading.BackgroundPatternColor = $f.shade }
    if ($f.ContainsKey('keep')) { $pf.KeepWithNext = -1 }
  }
  # Type a run with its own character formatting, then go back to the style's formatting.
  function Tx([string]$t, [hashtable]$f = @{}) {
    $ft = $sel.Font
    if ($f.ContainsKey('font')) { $ft.Name = $f.font }
    if ($f.ContainsKey('size')) { $ft.Size = [single]$f.size }
    if ($f.ContainsKey('bold')) { $ft.Bold = [int]$f.bold }
    if ($f.ContainsKey('italic')) { $ft.Italic = [int]$f.italic }
    if ($f.ContainsKey('color')) { $ft.Color = $f.color }
    if ($f.ContainsKey('caps')) { $ft.SmallCaps = -1 }
    if ($f.ContainsKey('allcaps')) { $ft.AllCaps = -1 }
    if ($f.ContainsKey('spacing')) { $ft.Spacing = [single]$f.spacing }
    if ($f.ContainsKey('under')) { $ft.Underline = 1 }
    if ($f.ContainsKey('sup')) { $ft.Superscript = -1 }
    $sel.TypeText($t)
    $sel.Font.Reset()
  }
  function EndP { $sel.TypeParagraph() }
  # One-run paragraph: paragraph props and run props share one table (keys don't overlap).
  function P([string]$t, [int]$style = -1, [hashtable]$f = @{}) { Para $style $f; Tx $t $f; EndP }
  function Bullets([string[]]$items, [hashtable]$f = @{}) { foreach ($i in $items) { P $i -49 $f } }
  function Numbers([string[]]$items, [hashtable]$f = @{}) {
    $first = $true
    foreach ($i in $items) {
      Para -50 $f
      if ($first) { $lf = $sel.Range.ListFormat; if ($lf.ListTemplate) { $lf.ApplyListTemplate($lf.ListTemplate, $false) }; $first = $false }
      Tx $i $f; EndP
    }
  }
  function Margins($top, $bottom, $left, $right, [switch]$A4) {
    $ps = $doc.PageSetup
    if ($A4) { $ps.PageWidth = 595.3; $ps.PageHeight = 841.9 } else { $ps.PageWidth = 612; $ps.PageHeight = 792 }
    $ps.TopMargin = $top; $ps.BottomMargin = $bottom; $ps.LeftMargin = $left; $ps.RightMargin = $right
  }
  function SetNormal($font, $size, $after = 6, $line = 1.08) {
    $n = $doc.Styles.Item(-1); $n.Font.Name = $font; $n.Font.Size = $size
    $n.ParagraphFormat.SpaceAfter = $after; $n.ParagraphFormat.LineSpacingRule = 5; $n.ParagraphFormat.LineSpacing = 12 * $line
  }
  function StyleFont([int]$id, $font, $size, $color, $bold = 1, $before = $null, $after = $null) {
    $s = $doc.Styles.Item($id); $s.Font.Name = $font; $s.Font.Size = $size; $s.Font.Color = $color; $s.Font.Bold = $bold
    if ($before -ne $null) { $s.ParagraphFormat.SpaceBefore = $before }
    if ($after -ne $null) { $s.ParagraphFormat.SpaceAfter = $after }
  }
  # Put a table at the cursor (cursor must sit in an empty Normal paragraph).
  function Table($rows, $cols, [object[]]$widths, [bool]$borders = $true) {
    Para -1
    $t = $doc.Tables.Add($sel.Range, $rows, $cols)
    $t.Borders.Enable = [int]$borders
    if ($widths) { for ($i = 0; $i -lt $widths.Count; $i++) { $t.Columns.Item($i + 1).Width = $widths[$i] } }
    $t
  }
  function Cell($t, $r, $c, [string]$text, [hashtable]$f = @{}) {
    $rg = $t.Cell($r, $c).Range; $rg.Text = $text
    if ($f.ContainsKey('bold')) { $rg.Font.Bold = [int]$f.bold }
    if ($f.ContainsKey('italic')) { $rg.Font.Italic = [int]$f.italic }
    if ($f.ContainsKey('size')) { $rg.Font.Size = [single]$f.size }
    if ($f.ContainsKey('color')) { $rg.Font.Color = $f.color }
    if ($f.ContainsKey('font')) { $rg.Font.Name = $f.font }
    if ($f.ContainsKey('align')) { $rg.ParagraphFormat.Alignment = $f.align }
    if ($f.ContainsKey('shade')) { $t.Cell($r, $c).Shading.BackgroundPatternColor = $f.shade }
    $rg.ParagraphFormat.SpaceAfter = $(if ($f.ContainsKey('after')) { $f.after } else { 0 })
  }
  function Row($t, $r, [string[]]$vals, [hashtable]$f = @{}, [int[]]$right = @()) {
    for ($c = 1; $c -le $vals.Count; $c++) {
      $g = $f.Clone(); if ($right -contains $c) { $g.align = 2 }
      Cell $t $r $c $vals[$c - 1] $g
    }
  }
  # Move the cursor into a cell to type rich content there.
  function InCell($t, $r, $c) { $t.Cell($r, $c).Range.Select(); $sel.Collapse(1) }
  function Picture($path, $w, $h) {
    $s = $sel.InlineShapes.AddPicture($path, $false, $true)
    $s.LockAspectRatio = 0; $s.Width = $w; $s.Height = $h; $s
  }
  function Border($side, $color, $width = 6, $space = 1) { # side: -1 top, -3 bottom
    $b = $sel.ParagraphFormat.Borders.Item($side); $b.LineStyle = 1; $b.LineWidth = $width; $b.Color = $color
    $sel.ParagraphFormat.Borders.DistanceFromBottom = $space
  }
  # Footer "text | Page X of Y" (or header text) on section 1.
  function FooterPages([string]$text, $align = 1, $size = 9, $color = 0, [switch]$OfTotal) {
    $ft = $doc.Sections.Item(1).Footers.Item(1).Range
    $ft.Text = $text; $ft.Font.Size = $size; $ft.Font.Color = $color; $ft.ParagraphFormat.Alignment = $align
    $r = $doc.Sections.Item(1).Footers.Item(1).Range; $r.MoveEnd(1, -1) | Out-Null; $r.Collapse(0)
    $doc.Fields.Add($r, 33) | Out-Null
    if ($OfTotal) {
      $r = $doc.Sections.Item(1).Footers.Item(1).Range; $r.MoveEnd(1, -1) | Out-Null; $r.Collapse(0); $r.InsertAfter(' of ')
      $r = $doc.Sections.Item(1).Footers.Item(1).Range; $r.MoveEnd(1, -1) | Out-Null; $r.Collapse(0)
      $doc.Fields.Add($r, 26) | Out-Null
    }
  }
  function Footnote([string]$text) { $doc.Footnotes.Add($sel.Range, $M, $text) | Out-Null; GoEnd }

  try {
  switch ($Name) {
  # ------------------------------------------------------------------ resume
  'resume' {
    Margins 40 40 46 46
    SetNormal 'Calibri' 10.5 3 1.0
    $navy = C 23 54 93; $teal = C 31 111 139; $side = C 234 240 246
    P 'Daniel R. Okafor' -1 @{ size = 26; bold = 1; color = $navy; after = 0; font = 'Calibri Light' }
    P 'Senior Data Analyst' -1 @{ size = 13; color = $teal; after = 4 }
    Para -1 @{ after = 10 }; Border -3 $teal 12 4
    Tx "Leeds, United Kingdom  |  +44 7700 900 418  |  daniel.okafor@mailbox.example  |  linkedin.com/in/danielokafor"  @{ size = 9.5; color = (C 80 80 80) }
    EndP
    $sel.ParagraphFormat.Borders.Enable = 0
    $t = Table 1 2 @(168, 352) $false
    $t.Cell(1, 1).Shading.BackgroundPatternColor = $side
    $t.Cell(1, 1).TopPadding = 6; $t.Cell(1, 1).LeftPadding = 8; $t.Cell(1, 1).RightPadding = 8
    $t.Cell(1, 2).LeftPadding = 14
    InCell $t 1 1
    $h = @{ size = 11; bold = 1; color = $navy; allcaps = 1; spacing = 0.5; before = 6; after = 3 }
    P 'Skills' -1 $h
    foreach ($s in 'SQL (PostgreSQL, BigQuery)', 'Python: pandas, scikit-learn', 'Power BI and Tableau', 'dbt and Airflow pipelines', 'A/B test design', 'Forecasting and regression', 'Stakeholder workshops') { P "$bul  $s" -1 @{ size = 9.5; after = 2 } }
    P 'Education' -1 $h
    P 'MSc Business Analytics' -1 @{ bold = 1; size = 9.5; after = 0 }
    P 'University of Manchester, 2016' -1 @{ size = 9.5; italic = 1; after = 5 }
    P 'BSc Economics (First Class)' -1 @{ bold = 1; size = 9.5; after = 0 }
    P 'University of Leeds, 2015' -1 @{ size = 9.5; italic = 1; after = 5 }
    P 'Certifications' -1 $h
    P 'Google Cloud Professional Data Engineer (2024)' -1 @{ size = 9.5; after = 2 }
    P 'Microsoft PL-300 Power BI Analyst (2022)' -1 @{ size = 9.5; after = 2 }
    P 'Languages' -1 $h
    P 'English (native), French (professional), Yoruba (conversational)' -1 @{ size = 9.5; after = 0 }
    InCell $t 1 2
    $h2 = @{ size = 12; bold = 1; color = $navy; allcaps = 1; spacing = 0.5; before = 4; after = 4 }
    P 'Profile' -1 $h2
    P "Data analyst with eight years${rsq} experience turning operational and financial data into decisions. I build reliable reporting pipelines, design experiments, and explain results to people who do not live in spreadsheets. Most recently led the analytics rebuild that cut month-end reporting from nine days to three." -1 @{ align = 3; after = 6 }
    P 'Experience' -1 $h2
    Para -1 @{ after = 0 }; Tx 'Senior Data Analyst' @{ bold = 1; size = 11 }; Tx "  $em  Harrow & Vale Logistics, Leeds" @{ size = 11 }; EndP
    P "March 2021 $nd present" -1 @{ italic = 1; color = (C 110 110 110); size = 9.5; after = 3 }
    Bullets @('Rebuilt the finance reporting stack in dbt and BigQuery; month-end close reporting now takes three days instead of nine.', 'Designed a delivery-slot pricing experiment across 42 depots that raised off-peak bookings by 17% with no loss of revenue.', 'Built the route-cost dashboard used weekly by 60 regional managers; adoption rose from 12 to 54 active users in six months.', 'Mentor two junior analysts and run the fortnightly analytics clinic for non-technical teams.') @{ size = 10; after = 2 }
    Para -1 @{ after = 0; before = 6 }; Tx 'Data Analyst' @{ bold = 1; size = 11 }; Tx "  $em  Pennine Building Society, Halifax" @{ size = 11 }; EndP
    P "September 2016 $nd February 2021" -1 @{ italic = 1; color = (C 110 110 110); size = 9.5; after = 3 }
    Bullets @('Produced the monthly arrears and mortgage pipeline packs for the executive committee.', 'Automated 30 Excel reports with Python and SQL Server Reporting Services, saving roughly 25 hours a month.', 'Developed a churn model for savings customers (AUC 0.81) that guided the 2019 retention campaign.') @{ size = 10; after = 2 }
    Para -1 @{ after = 0; before = 6 }; Tx 'Graduate Analyst (placement)' @{ bold = 1; size = 11 }; Tx "  $em  Leeds City Council" @{ size = 11 }; EndP
    P "June 2015 $nd August 2016" -1 @{ italic = 1; color = (C 110 110 110); size = 9.5; after = 3 }
    Bullets @('Analysed footfall and parking data to support the city-centre transport consultation.', 'Cleaned and published eleven open datasets on the council data portal.') @{ size = 10; after = 2 }
    P 'Selected projects' -1 $h2
    Para -1 @{ after = 3 }; Tx 'Depot demand forecast. ' @{ bold = 1 }; Tx 'Weekly parcel-volume forecast for 42 depots using gradient boosting; mean absolute percentage error of 6.8%, now used for staff rostering.'; EndP
    Para -1 @{ after = 0 }; Tx 'Open reporting standards. ' @{ bold = 1 }; Tx 'Wrote the internal guide to chart and table design adopted by the finance and operations teams.'; EndP
    GoEnd
  }
  # ------------------------------------------------------------------ report
  'report' {
    Margins 72 72 72 72 -A4
    SetNormal 'Calibri' 11 8 1.08
    $blue = C 31 78 121
    StyleFont -2 'Calibri Light' 16 $blue 0 18 6
    StyleFont -3 'Calibri Light' 13 (C 46 116 181) 0 12 4
    $doc.Sections.Item(1).Headers.Item(1).Range.Text = "Northbridge Analytics $nd Workflow Automation Pilot: Final Report"
    $hr = $doc.Sections.Item(1).Headers.Item(1).Range; $hr.Font.Size = 9; $hr.Font.Color = (C 110 110 110); $hr.ParagraphFormat.Alignment = 2
    FooterPages 'Confidential  |  Page ' 1 9 (C 110 110 110) -OfTotal
    P 'Workflow Automation Pilot' -63
    P 'Final report to the Operations Steering Group' -75
    Para -1 @{ after = 14 }; Tx 'Prepared by: ' @{ bold = 1 }; Tx 'Priya Natarajan and Tom Whitfield     '; Tx 'Date: ' @{ bold = 1 }; Tx '3 November 2026     '; Tx 'Version: ' @{ bold = 1 }; Tx '1.2'; EndP
    P '1. Executive summary' -2
    P 'Between January 2025 and June 2026 the Operations and Finance teams piloted a set of automated workflows for invoice matching, purchase-order approval and weekly capacity reporting. The pilot covered 214 staff across four sites and replaced eleven manual processes that had previously relied on shared spreadsheets and email.' -1 @{ align = 3 }
    Para -1 @{ align = 3 }; Tx 'The results were stronger than the business case assumed. Average weekly hours saved per team rose steadily through the pilot and reached 13.9 hours in Operations by the final quarter. Invoice exceptions fell by 38%, and the median time to approve a purchase order dropped from 4.2 working days to 1.1. The total cost of the pilot was ' ; Tx "${pound}186,400"; Tx ', against a projected annual saving of '; Tx "${pound}412,000" @{ bold = 1 }; Tx '.'; Footnote 'Savings are calculated at the blended staff cost of 31.20 GBP per hour used in the 2025 business case, excluding on-costs.'; EndP
    P 'We recommend that the Steering Group:' -1 @{ after = 4 }
    Numbers @('approve the roll-out of the three highest-value workflows to all eleven sites from January 2027;', 'fund one permanent automation analyst post within the Finance Systems team;', 'adopt the governance model described in Section 4, including quarterly benefit reviews.')
    P '2. Background and scope' -2
    P 'The 2024 operating review found that staff in Operations and Finance spent an estimated 22% of their time re-keying data between systems. Most of this effort sat in a small number of high-volume, rules-based tasks. The pilot was designed to test whether low-code automation could remove that work safely, without a major systems change.' -1 @{ align = 3 }
    P '2.1 Processes in scope' -3
    Bullets @('Three-way matching of supplier invoices against purchase orders and goods receipts.', 'Routing and approval of purchase orders under 25,000 GBP.', 'Weekly site capacity and overtime reporting.', 'Month-end accrual reminders and evidence collection.')
    P '2.2 Out of scope' -3
    P 'Payroll, customer billing and any process involving personal data beyond staff names and work email addresses were excluded. These areas will be assessed separately once the data protection impact assessment has been completed.' -1 @{ align = 3 }
    P '3. Results' -2
    P 'Table 1 summarises the key measures for each workflow. Baseline figures come from the six months before go-live; pilot figures are averages over the final two quarters.' -1 @{ align = 3 }
    P 'Table 1. Key measures before and during the pilot' -35 @{ keep = 1 }
    $t = Table 6 4 @(170, 90, 90, 101)
    $t.Rows.Item(1).Shading.BackgroundPatternColor = $blue; $t.Rows.Item(1).HeadingFormat = -1
    $t.Borders.InsideColor = (C 191 191 191); $t.Borders.OutsideColor = (C 191 191 191)
    Row $t 1 @('Workflow', 'Baseline', 'Pilot', 'Change') @{ bold = 1; color = (C 255 255 255) } @(2, 3, 4)
    Row $t 2 @('Invoice matching (exceptions per month)', '1,240', '769', "${nd}38.0%") @{} @(2, 3, 4)
    Row $t 3 @('PO approval time (working days)', '4.2', '1.1', "${nd}73.8%") @{} @(2, 3, 4)
    Row $t 4 @('Capacity report preparation (hours/week)', '9.5', '1.5', "${nd}84.2%") @{} @(2, 3, 4)
    Row $t 5 @('Accrual evidence received on time', '61%', '93%', '+32 pts') @{} @(2, 3, 4)
    Row $t 6 @('Staff satisfaction with tools (1-5)', '2.7', '4.1', '+1.4') @{} @(2, 3, 4)
    for ($r = 2; $r -le 6; $r += 2) { $t.Rows.Item($r).Shading.BackgroundPatternColor = (C 242 246 250) }
    $t.Rows.Alignment = 1
    GoEnd
    P '' -1 @{ after = 0 }
    P 'The savings grew over time rather than arriving at go-live. Figure 1 shows the average weekly hours saved per team in each quarter. The early quarters were limited by exception handling: staff had to check most automated matches by hand until the rules were tuned in the third quarter.' -1 @{ align = 3 }
    Para -1 @{ align = 1; after = 2; keep = 1 }; Picture "$Pics\chart.png" 430 229 | Out-Null; EndP
    P 'Figure 1. Average weekly hours saved per team, January 2025 to June 2026' -35 @{ align = 1 }
    P '3.1 What worked well' -3
    Bullets @('Starting with invoice matching, where rules were already documented and volumes were high.', 'Weekly show-and-tell sessions, which surfaced problems early and built trust in the automated results.', 'Keeping a human approval step for every payment above 5,000 GBP.')
    P '3.2 What was harder than expected' -3
    Bullets @('Supplier master data was inconsistent; around 9% of suppliers had duplicate records that blocked matching.', 'Two sites used local spreadsheets that were not visible to the central team until the pilot was under way.', 'Licence costs rose by 14% during the pilot after a vendor price change.')
    P '4. Recommendations and next steps' -2
    P 'We propose a phased roll-out over twelve months, starting with the workflows that showed the largest and most reliable benefits. Each phase will begin only after the previous phase has met its acceptance criteria.' -1 @{ align = 3 }
    Numbers @('Phase 1 (January to March 2027): invoice matching and PO approval at all sites.', 'Phase 2 (April to June 2027): capacity reporting and accrual reminders.', 'Phase 3 (July to December 2027): review of further candidates, including supplier onboarding.')
    P 'Governance will sit with the Finance Systems Board, which will review benefits, incidents and change requests each quarter. A short benefits report will be presented to this Steering Group after each phase.' -1 @{ align = 3 }
    P '5. Risks' -2
    P 'The main risks are poor supplier data, dependence on a single vendor and the loss of process knowledge as manual tasks disappear. Each risk has an owner and a mitigation plan in the project risk register; none is currently rated high.' -1 @{ align = 3 }
    GoEnd
  }
  # ------------------------------------------------------------------ letter
  'letter' {
    Margins 50 60 72 72 -A4
    SetNormal 'Georgia' 10.5 8 1.15
    $green = C 24 94 82
    $t = Table 1 2 @(80, 371) $false
    $r = $t.Cell(1, 1).Range; $r.Collapse(1)
    $pic = $doc.InlineShapes.AddPicture("$Pics\logo.png", $false, $true, $r); $pic.LockAspectRatio = 0; $pic.Width = 62; $pic.Height = 62
    InCell $t 1 2
    P 'Harrow & Vale Logistics Ltd' -1 @{ size = 18; bold = 1; color = $green; align = 2; after = 2 }
    P 'Unit 4, Aire Valley Business Park, Leeds LS9 0PT' -1 @{ size = 9; color = (C 90 90 90); align = 2; after = 0 }
    P 'Tel 0113 496 0721  |  accounts@harrowvale.example  |  www.harrowvale.example' -1 @{ size = 9; color = (C 90 90 90); align = 2; after = 0 }
    $t.Rows.Item(1).Cells.VerticalAlignment = 1
    $t.Borders.Item(-3).LineStyle = 1; $t.Borders.Item(-3).LineWidth = 12; $t.Borders.Item(-3).Color = $green
    GoEnd
    $ft = $doc.Sections.Item(1).Footers.Item(1).Range
    $ft.Text = 'Harrow & Vale Logistics Ltd is registered in England and Wales, company no. 08812345. Registered office as above. VAT reg. GB 284 5521 09.'
    $ft.Font.Size = 8; $ft.Font.Color = (C 120 120 120); $ft.ParagraphFormat.Alignment = 1
    P '' -1 @{ after = 12 }
    P '11 November 2026' -1 @{ after = 14 }
    foreach ($l in 'Ms Rachel Denholm', 'Head of Procurement', 'Calder Retail Group plc', '18 Wellington Street', 'Bradford', 'BD1 2SR') { P $l -1 @{ after = 0 } }
    P '' -1 @{ after = 0 }
    P 'Your ref: CRG/PR/2291      Our ref: HV/RD/1106' -1 @{ size = 9.5; color = (C 90 90 90); after = 12 }
    P 'Dear Ms Denholm,' -1 @{ after = 10 }
    P 'Renewal of the regional distribution agreement' -1 @{ bold = 1; after = 10 }
    P "Thank you for your letter of 28 October and for the time you and your team gave us at the review meeting in Bradford. I am pleased to confirm that we would like to renew the regional distribution agreement for a further three years from 1 February 2027, on the terms set out below." -1 @{ align = 3 }
    P "Over the past contract year we delivered 96.4% of your consignments within the agreed window, against a target of 95%, and we resolved all but two damage claims within ten working days. We recognise that the Christmas peak remains the most difficult period, and we have already reserved additional trailer capacity at our Wakefield cross-dock for December." -1 @{ align = 3 }
    P 'For the renewed agreement we propose:' -1 @{ after = 4 }
    Bullets @("a rate increase of 3.2% from February 2027, below the 4.1% rise in our fuel and labour costs;", 'a new next-day service for online orders placed before 4 pm, at the standard regional rate;', 'quarterly service reviews, alternating between your offices and our Leeds depot.') @{ font = 'Georgia'; size = 10.5 }
    P "I enclose the draft agreement and the revised rate card. If it would help, my colleague James Ashworth would be happy to take your team through the figures before your board meets on 2 December. We value our partnership with Calder Retail and look forward to continuing it." -1 @{ align = 3; before = 4 }
    P 'Yours sincerely,' -1 @{ after = 2; before = 6 }
    Para -1 @{ after = 0 }; Picture "$Pics\signature.png" 140 37 | Out-Null; EndP
    P 'Helen Marsh' -1 @{ bold = 1; after = 0 }
    P 'Commercial Director' -1 @{ after = 0 }
    P 'Harrow & Vale Logistics Ltd' -1 @{ after = 12 }
    P 'Enc. Draft agreement (v3), Rate card 2027' -1 @{ size = 9; italic = 1; color = (C 90 90 90) }
    GoEnd
  }
  # ------------------------------------------------------------------ newsletter
  'newsletter' {
    Margins 36 36 40 40 -A4
    SetNormal 'Segoe UI' 9.5 6 1.08
    $plum = C 91 44 111; $gold = C 196 140 30
    P 'THE RIVERSIDE BULLETIN' -1 @{ size = 30; bold = 1; color = (C 255 255 255); shade = $plum; align = 1; after = 0; before = 0; font = 'Georgia'; spacing = 2 }
    P "Kirkstall Riverside Residents${rsq} Association  |  Issue 47  |  Autumn 2026" -1 @{ size = 10; color = (C 255 255 255); shade = $plum; align = 1; after = 12 }
    Para -1 @{ after = 0 }
    $sel.InsertBreak(3)
    $doc.Sections.Item(2).PageSetup.TextColumns.SetCount(2)
    $doc.Sections.Item(2).PageSetup.TextColumns.Spacing = 22
    $h = @{ size = 14; bold = 1; color = $plum; font = 'Georgia'; after = 4; before = 2; keep = 1 }
    P 'Community centre opens its doors' -1 $h
    P 'After three years of fundraising and eighteen months of building work, the new Riverside Community Centre on Abbey Road opened on Saturday 4 October. More than 600 people came through the doors on the first day, and the cafe sold out of scones by eleven o clock.' -1 @{ align = 3 }
    P 'The centre has a main hall for up to 150 people, two meeting rooms, a kitchen run by volunteers, and a small garden designed by pupils from Abbey Grange Primary. Rooms can be booked from 12 GBP an hour, with a half-price rate for local groups and charities.' -1 @{ align = 3 }
    Para -1 @{ align = 1; after = 2; keep = 1 }; Picture "$Pics\photo.png" 240 149 | Out-Null; EndP
    P 'The new centre seen from the river path. Photo: Maureen Hall' -1 @{ italic = 1; size = 8; color = (C 100 100 100); align = 1; after = 10 }
    P 'Flood defence work starts in January' -1 $h
    P 'The Environment Agency has confirmed that the second phase of the Leeds flood alleviation scheme will reach our stretch of the river in January. Work will include raised embankments between the weir and the footbridge and a new pumping station behind the allotments.' -1 @{ align = 3 }
    P 'The towpath will close in sections for up to six weeks at a time. Diversion maps will be posted at both ends of the path and on our website. If you have questions, the project team holds a drop-in session at the centre on the first Tuesday of every month.' -1 @{ align = 3 }
    Para -1 @{ align = 1; before = 8; after = 10; left = 8; right = 8 }
    Border -1 $gold 18 4; Border -3 $gold 18 4
    Tx "${ldq}For the first time in years we have somewhere to meet that belongs to all of us.${rdq}" @{ size = 14; italic = 1; color = $plum; font = 'Georgia' }
    EndP
    $sel.ParagraphFormat.Borders.Enable = 0
    Para -1 @{ align = 1; after = 10 }; Tx "$em Jean Okoro, chair of the association" @{ size = 8.5; color = (C 100 100 100) }; EndP
    P 'Volunteers wanted for the winter' -1 $h
    P 'Our warm-space sessions return on Monday and Thursday afternoons from November to March. We need volunteers to make tea, chat with guests and help with lifts. No experience is needed and we provide a short induction.' -1 @{ align = 3 }
    Bullets @('Monday 1 pm to 4 pm: hall and cafe', 'Thursday 1 pm to 4 pm: hall, cafe and IT help desk', 'Driver rota: flexible, mileage paid') @{ size = 9.5 }
    P 'Diary dates' -1 $h
    foreach ($d in @(@('22 Nov', 'Litter pick, meet at the footbridge, 10 am'), @('6 Dec', 'Winter fair and lantern walk, 3 pm'), @('14 Jan', 'Annual general meeting, 7 pm'), @('8 Feb', 'Repair cafe: bikes, lamps and clothes'))) {
      Para -1 @{ after = 3 }; Tx ($d[0] + '   ') @{ bold = 1; color = $gold }; Tx $d[1]; EndP
    }
    P 'Contact the association at hello@riverside.example or leave a note in the box at the centre reception.' -1 @{ italic = 1; size = 8.5; color = (C 100 100 100); before = 6 }
    GoEnd
  }
  # ------------------------------------------------------------------ paper
  'paper' {
    Margins 72 72 72 72
    SetNormal 'Times New Roman' 11 6 1.0
    StyleFont -2 'Times New Roman' 12 0 1 12 6
    StyleFont -3 'Times New Roman' 11 0 1 8 4
    $doc.Styles.Item(-3).Font.Italic = 1
    FooterPages '' 1 10 0
    P 'Estimating Commuter Mode Shift from Bus Priority Corridors: Evidence from Three Mid-Sized UK Cities' -1 @{ size = 16; bold = 1; align = 1; after = 12; line = 1.0 }
    Para -1 @{ align = 1; after = 2 }; Tx 'Sarah J. Whitcombe'; Tx '1' @{ sup = 1 }; Tx ', Rahul Menon'; Tx '2' @{ sup = 1 }; Tx ' and Elena Kowalski'; Tx '1' @{ sup = 1 }; EndP
    P '1 Institute for Transport Studies, University of Leeds, UK     2 School of Geography, University of Nottingham, UK' -1 @{ size = 9; italic = 1; align = 1; after = 2 }
    P 'Corresponding author: s.whitcombe@leeds.example' -1 @{ size = 9; align = 1; after = 14 }
    P 'Abstract' -1 @{ bold = 1; align = 1; after = 4 }
    P 'Bus priority corridors are widely promoted as a low-cost way to reduce car commuting, but evidence on their effect on travel mode choice remains mixed. We combine automatic passenger counts, ticketing records and a panel survey of 2,184 commuters to estimate mode shift on six corridors opened between 2019 and 2023 in Leicester, Nottingham and Bristol. Using a difference-in-differences design with matched control corridors, we find that corridor opening increased bus mode share among commuters living within 800 m of the route by 4.6 percentage points (95% CI 2.9 to 6.3) after two years. Roughly 60% of new bus trips came from former car drivers. Effects were larger where journey-time savings exceeded five minutes and where parking at the destination was charged. The results suggest that corridor design and complementary parking policy matter more than corridor length.' -1 @{ align = 3; left = 36; right = 36; size = 10; after = 6 }
    Para -1 @{ left = 36; right = 36; after = 14 }; Tx 'Keywords: ' @{ bold = 1; size = 10 }; Tx 'bus priority; mode choice; difference-in-differences; urban transport policy' @{ size = 10; italic = 1 }; EndP
    P '1  Introduction' -2
    Para -1 @{ align = 3; first = 18 }; Tx 'Road space reallocation to buses has become a central element of local transport plans in England since the National Bus Strategy of 2021. Corridors that combine bus lanes, signal priority and improved stops promise faster and more reliable journeys at a fraction of the cost of light rail. Yet the empirical literature reports effects on mode choice that range from negligible to substantial [1, 3]. Much of this variation reflects differences in design, context and the counterfactual used.'; Footnote 'We use corridor to mean a continuous route of at least 3 km with bus lanes or bus gates covering more than half of its length.'; EndP
    P 'This paper contributes in three ways. First, it uses a panel survey that follows the same commuters before and after corridor opening, which avoids the composition bias of repeated cross-sections. Second, it compares treated corridors with matched controls in the same cities. Third, it examines how effects vary with journey-time savings and destination parking charges, two levers that local authorities control directly.' -1 @{ align = 3; first = 18 }
    P '2  Data and methods' -2
    P '2.1  Study corridors' -3
    P 'We study six corridors that opened between 2019 and 2023 (Table 1). Each was matched to a control corridor in the same city with similar baseline bus frequency, population density and distance to the city centre. Control corridors received no significant bus priority measures during the study period.' -1 @{ align = 3; first = 18 }
    P 'Table 1. Treated corridors and their main characteristics' -1 @{ size = 10; align = 1; before = 6; after = 4; keep = 1 }
    $t = Table 7 5 @(110, 70, 70, 85, 85)
    $t.Borders.Enable = 0
    foreach ($side in -1, -3) { $t.Rows.Item(1).Borders.Item($side).LineStyle = 1; $t.Rows.Item(1).Borders.Item($side).LineWidth = 8 }
    $t.Rows.Item(7).Borders.Item(-3).LineStyle = 1; $t.Rows.Item(7).Borders.Item(-3).LineWidth = 8
    Row $t 1 @('Corridor', 'City', 'Length (km)', 'Opened', 'Time saving (min)') @{ bold = 1; size = 10 } @(3, 5)
    Row $t 2 @('A6 London Road', 'Leicester', '5.8', 'Mar 2019', '6.2') @{ size = 10 } @(3, 5)
    Row $t 3 @('Hinckley Road', 'Leicester', '4.1', 'Sep 2021', '3.4') @{ size = 10 } @(3, 5)
    Row $t 4 @('Derby Road', 'Nottingham', '6.5', 'Jan 2020', '5.7') @{ size = 10 } @(3, 5)
    Row $t 5 @('Carlton Road', 'Nottingham', '3.6', 'Jun 2022', '2.9') @{ size = 10 } @(3, 5)
    Row $t 6 @('A38 Bedminster', 'Bristol', '4.9', 'Apr 2021', '7.1') @{ size = 10 } @(3, 5)
    Row $t 7 @('Fishponds Road', 'Bristol', '5.2', 'Nov 2023', '4.4') @{ size = 10 } @(3, 5)
    $t.Rows.Alignment = 1
    GoEnd
    P '' -1 @{ after = 0 }
    P '2.2  Estimation' -3
    P 'Let y(it) denote whether commuter i travels to work mainly by bus in wave t. We estimate a linear probability model with individual and wave fixed effects, in which the coefficient on the interaction between living near a treated corridor and the post-opening period identifies the average effect of corridor opening. Standard errors are clustered at the corridor level, and we report wild bootstrap confidence intervals because the number of clusters is small [2].' -1 @{ align = 3; first = 18 }
    P '3  Results' -2
    P 'Corridor opening increased the probability of commuting by bus by 4.6 percentage points two years after opening. The effect appears within six months and grows over the following eighteen months, consistent with gradual habit change. Among new bus commuters, 61% had previously driven alone, 22% had been car passengers and 17% had walked or cycled.' -1 @{ align = 3; first = 18 }
    P 'Effects were concentrated on corridors with journey-time savings above five minutes, where the estimated shift was 7.3 points, compared with 1.8 points elsewhere. Where the destination area had charged parking, the effect was roughly twice as large as where parking was free. We find no evidence of anticipation effects in the waves before opening.' -1 @{ align = 3; first = 18 }
    P '4  Conclusions' -2
    P 'Bus priority corridors can shift a meaningful share of commuters out of cars, but only when they deliver a noticeable time saving and when the cost of parking at the destination is not zero. Local authorities planning new corridors should prioritise routes where buses currently suffer the largest delays and consider corridor schemes together with workplace parking policy.' -1 @{ align = 3; first = 18 }
    P 'References' -1 @{ bold = 1; size = 12; before = 12; after = 6 }
    $refs = @(
      '[1]  Batty, P., Palacin, R. and Gonzalez-Gil, A. (2015) Challenges and opportunities in developing urban modal shift. Travel Behaviour and Society, 2(2), pp. 109-123.',
      '[2]  Cameron, A.C., Gelbach, J.B. and Miller, D.L. (2008) Bootstrap-based improvements for inference with clustered errors. Review of Economics and Statistics, 90(3), pp. 414-427.',
      '[3]  Currie, G. and Sarvi, M. (2012) New model for secondary benefits of transit priority. Transportation Research Record, 2276(1), pp. 63-71.',
      '[4]  Department for Transport (2021) Bus Back Better: National Bus Strategy for England. London: DfT.',
      '[5]  Wang, Y., Moudon, A.V. and Shen, Q. (2022) How does ride-hailing influence bus ridership? Transportation Research Part A, 158, pp. 1-15.')
    foreach ($r in $refs) { P $r -1 @{ size = 10; left = 22; first = -22; after = 3; align = 0 } }
    GoEnd
  }
  # ------------------------------------------------------------------ invoice
  'invoice_word' {
    Margins 50 50 56 56 -A4
    SetNormal 'Arial' 9.5 0 1.0
    $dark = C 38 50 56; $acc = C 0 121 107
    $t = Table 1 2 @(260, 223) $false
    InCell $t 1 1
    P 'Fenwick & Clarke Studio' -1 @{ size = 16; bold = 1; color = $dark; after = 2 }
    foreach ($l in '27 Market Place', 'Knaresborough HG5 8AL', 'United Kingdom', 'VAT GB 391 7720 46') { P $l -1 @{ size = 9; color = (C 90 90 90) } }
    InCell $t 1 2
    P 'INVOICE' -1 @{ size = 28; bold = 1; color = $acc; align = 2; after = 4 }
    P 'Invoice no.  FC-2026-0418' -1 @{ align = 2; size = 9.5 }
    P 'Issue date:  11 November 2026' -1 @{ align = 2; size = 9.5 }
    P 'Due date:  11 December 2026' -1 @{ align = 2; size = 9.5; bold = 1 }
    GoEnd
    P '' -1 @{ after = 14 }
    $t = Table 1 2 @(241, 242) $false
    InCell $t 1 1
    P 'BILL TO' -1 @{ size = 8; bold = 1; color = $acc; spacing = 1; after = 3 }
    foreach ($l in 'Moorside Veterinary Group', 'Attn: Accounts Payable', '5 Station Parade', 'Harrogate HG1 1TS') { P $l -1 @{} }
    InCell $t 1 2
    P 'PROJECT' -1 @{ size = 8; bold = 1; color = $acc; spacing = 1; after = 3 }
    P 'Website redesign and brand refresh' -1 @{}
    P 'Purchase order: MVG-PO-7731' -1 @{}
    P 'Account manager: Olivia Fenwick' -1 @{}
    GoEnd
    P '' -1 @{ after = 14 }
    $t = Table 8 5 @(30, 243, 50, 80, 80)
    $t.Borders.Enable = 0
    for ($r = 1; $r -le 8; $r++) { $b = $t.Rows.Item($r).Borders.Item(-3); $b.LineStyle = 1; $b.Color = (C 210 210 210) }
    $t.Rows.Item(1).Shading.BackgroundPatternColor = $dark
    $t.TopPadding = 3; $t.BottomPadding = 3
    Row $t 1 @('#', 'Description', 'Qty', 'Unit price', 'Amount') @{ bold = 1; color = (C 255 255 255); size = 9 } @(3, 4, 5)
    $items = @(
      @('1', 'Discovery workshop and stakeholder interviews', '1', '1,450.00', '1,450.00'),
      @('2', 'Brand refresh: logo, colour palette and type system', '1', '2,800.00', '2,800.00'),
      @('3', 'Website design, 9 page templates', '9', '420.00', '3,780.00'),
      @('4', 'Front-end build and content management setup', '38', '75.00', '2,850.00'),
      @('5', 'Photography, half-day shoot at two practices', '2', '550.00', '1,100.00'),
      @('6', 'Copy editing and search optimisation', '12', '65.00', '780.00'),
      @('7', 'Hosting and maintenance, first 12 months', '12', '35.00', '420.00'))
    for ($i = 0; $i -lt $items.Count; $i++) { Row $t ($i + 2) $items[$i] @{ size = 9 } @(3, 4, 5) }
    GoEnd
    P '' -1 @{ after = 10 }
    $t = Table 4 2 @(110, 90) $false
    $t.Rows.Alignment = 2
    Row $t 1 @('Subtotal', "${pound}13,180.00") @{} @(2)
    Row $t 2 @('Discount (5%)', "${nd}${pound}659.00") @{} @(2)
    Row $t 3 @('VAT at 20%', "${pound}2,504.20") @{} @(2)
    Row $t 4 @('Total due', "${pound}15,025.20") @{ bold = 1; size = 11; color = (C 255 255 255); shade = $acc } @(2)
    $t.TopPadding = 3; $t.BottomPadding = 3
    GoEnd
    P '' -1 @{ after = 18 }
    P 'Payment details' -1 @{ bold = 1; size = 10; color = $dark; after = 4 }
    P 'Bank transfer to: Fenwick & Clarke Studio Ltd   Sort code: 40-12-76   Account: 61442209' -1 @{ after = 2 }
    P 'Please quote invoice number FC-2026-0418 as the payment reference. Payment terms: 30 days from the issue date.' -1 @{ after = 14 }
    P 'Notes' -1 @{ bold = 1; size = 10; color = $dark; after = 4 }
    P 'Hosting renews automatically each November unless cancelled in writing 30 days before the renewal date. Additional design changes after sign-off are charged at 65.00 GBP per hour.' -1 @{ color = (C 90 90 90); after = 24 }
    P 'Thank you for your business.' -1 @{ align = 1; italic = 1; size = 11; color = $acc }
    $ft = $doc.Sections.Item(1).Footers.Item(1).Range
    $ft.Text = 'Fenwick & Clarke Studio Ltd  |  Registered in England no. 11920334  |  hello@fenwickclarke.example  |  01423 860 214'
    $ft.Font.Size = 8; $ft.Font.Color = (C 130 130 130); $ft.ParagraphFormat.Alignment = 1
    GoEnd
  }
  default { throw "unknown doc $Name" }
  }
  } finally {}

  $pages = $doc.ComputeStatistics(2)
  $pdf = [string]"$Out\$Name.pdf"
  $doc.ExportAsFixedFormat($pdf, 17)
  $tmpDocx = [string]"$Tmp\$Name.docx"
  if (Test-Path $tmpDocx) { Remove-Item $tmpDocx }
  $doc.SaveAs2($tmpDocx, 16)
  $doc.Close(0)
  Copy-Item $tmpDocx (Join-Path $Out "$Name.docx") -Force
  [ordered]@{ doc = $Name; ok = $true; pages = $pages } | ConvertTo-Json -Compress
} catch {
  [ordered]@{ doc = $Name; ok = $false; error = $_.Exception.Message; line = $_.InvocationInfo.ScriptLineNumber } | ConvertTo-Json -Compress
} finally {
  try { $word.Quit() } catch {}
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
