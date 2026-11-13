# Draws the PNG pictures used by the Word sample documents.
param([string]$Out)
Add-Type -AssemblyName System.Drawing
New-Item -ItemType Directory -Force $Out | Out-Null
function NewBmp($w, $h) { $b = New-Object System.Drawing.Bitmap $w, $h; $b.SetResolution(150, 150); $b }
function Col($r, $g, $b) { [System.Drawing.Color]::FromArgb(255, $r, $g, $b) }

# 1) Company logo: rounded square with initials
$b = NewBmp 300 300; $g = [System.Drawing.Graphics]::FromImage($b); $g.SmoothingMode = 'AntiAlias'; $g.TextRenderingHint = 'AntiAliasGridFit'
$g.Clear([System.Drawing.Color]::White)
$br = New-Object System.Drawing.Drawing2D.LinearGradientBrush((New-Object System.Drawing.Point 0, 0), (New-Object System.Drawing.Point 300, 300), (Col 18 78 140), (Col 20 140 160))
$g.FillEllipse($br, 10, 10, 280, 280)
$f = New-Object System.Drawing.Font('Segoe UI', 54, [System.Drawing.FontStyle]::Bold)
$sf = New-Object System.Drawing.StringFormat; $sf.Alignment = 'Center'; $sf.LineAlignment = 'Center'
$g.DrawString('HV', $f, [System.Drawing.Brushes]::White, (New-Object System.Drawing.RectangleF 0, 0, 300, 300), $sf)
$g.Dispose(); $b.Save("$Out\logo.png", [System.Drawing.Imaging.ImageFormat]::Png); $b.Dispose()

# 2) Bar chart for the report
$b = NewBmp 1200 640; $g = [System.Drawing.Graphics]::FromImage($b); $g.SmoothingMode = 'AntiAlias'; $g.TextRenderingHint = 'AntiAliasGridFit'
$g.Clear([System.Drawing.Color]::White)
$axis = New-Object System.Drawing.Pen((Col 120 120 120), 2); $grid = New-Object System.Drawing.Pen((Col 225 225 225), 1)
$lf = New-Object System.Drawing.Font('Calibri', 13); $tf = New-Object System.Drawing.Font('Calibri', 16, [System.Drawing.FontStyle]::Bold)
$g.DrawString('Average weekly hours saved per team, by quarter', $tf, (New-Object System.Drawing.SolidBrush (Col 40 40 40)), 80, 18)
$x0 = 110; $y0 = 560; $hgt = 440
for ($i = 0; $i -le 5; $i++) { $y = $y0 - $i * $hgt / 5; $g.DrawLine($grid, $x0, $y, 1150, $y); $g.DrawString(("{0}" -f ($i * 4)), $lf, [System.Drawing.Brushes]::DimGray, 60, $y - 12) }
$g.DrawLine($axis, $x0, $y0, 1150, $y0); $g.DrawLine($axis, $x0, $y0, $x0, 100)
$q = 'Q1 2025', 'Q2 2025', 'Q3 2025', 'Q4 2025', 'Q1 2026', 'Q2 2026'; $v1 = 3.1, 4.6, 6.2, 8.8, 11.4, 13.9; $v2 = 2.2, 2.9, 4.1, 5.0, 7.6, 9.3
$c1 = New-Object System.Drawing.SolidBrush (Col 31 92 153); $c2 = New-Object System.Drawing.SolidBrush (Col 237 125 49)
for ($i = 0; $i -lt 6; $i++) {
  $cx = $x0 + 40 + $i * 170
  $h1 = $v1[$i] / 20 * $hgt; $h2 = $v2[$i] / 20 * $hgt
  $g.FillRectangle($c1, $cx, $y0 - $h1, 55, $h1); $g.FillRectangle($c2, $cx + 60, $y0 - $h2, 55, $h2)
  $g.DrawString($q[$i], $lf, [System.Drawing.Brushes]::DimGray, $cx + 8, $y0 + 10)
}
$g.FillRectangle($c1, 820, 70, 20, 20); $g.DrawString('Operations', $lf, [System.Drawing.Brushes]::Black, 845, 68)
$g.FillRectangle($c2, 1010, 70, 20, 20); $g.DrawString('Finance', $lf, [System.Drawing.Brushes]::Black, 1035, 68)
$g.Dispose(); $b.Save("$Out\chart.png", [System.Drawing.Imaging.ImageFormat]::Png); $b.Dispose()

# 3) Landscape "photo" for the newsletter
$b = NewBmp 900 560; $g = [System.Drawing.Graphics]::FromImage($b); $g.SmoothingMode = 'AntiAlias'
$sky = New-Object System.Drawing.Drawing2D.LinearGradientBrush((New-Object System.Drawing.Point 0, 0), (New-Object System.Drawing.Point 0, 380), (Col 120 175 225), (Col 235 225 200))
$g.FillRectangle($sky, 0, 0, 900, 560)
$g.FillEllipse((New-Object System.Drawing.SolidBrush (Col 255 236 170)), 640, 70, 110, 110)
$hill = New-Object System.Drawing.Drawing2D.GraphicsPath
$hill.AddBezier(0, 360, 250, 230, 450, 420, 900, 300); $hill.AddLine(900, 300, 900, 560); $hill.AddLine(900, 560, 0, 560); $hill.CloseFigure()
$g.FillPath((New-Object System.Drawing.SolidBrush (Col 92 140 80)), $hill)
$hill2 = New-Object System.Drawing.Drawing2D.GraphicsPath
$hill2.AddBezier(0, 470, 300, 380, 600, 520, 900, 430); $hill2.AddLine(900, 430, 900, 560); $hill2.AddLine(900, 560, 0, 560); $hill2.CloseFigure()
$g.FillPath((New-Object System.Drawing.SolidBrush (Col 60 105 60)), $hill2)
# a building: the new community centre
$g.FillRectangle((New-Object System.Drawing.SolidBrush (Col 210 200 185)), 300, 300, 260, 150)
$roof = New-Object System.Drawing.Drawing2D.GraphicsPath; $roof.AddPolygon(@((New-Object System.Drawing.PointF 285, 300), (New-Object System.Drawing.PointF 430, 230), (New-Object System.Drawing.PointF 575, 300)))
$g.FillPath((New-Object System.Drawing.SolidBrush (Col 150 70 50)), $roof)
for ($i = 0; $i -lt 4; $i++) { $g.FillRectangle((New-Object System.Drawing.SolidBrush (Col 90 130 170)), 320 + $i * 60, 330, 35, 40) }
$g.FillRectangle((New-Object System.Drawing.SolidBrush (Col 110 75 50)), 410, 390, 40, 60)
$g.Dispose(); $b.Save("$Out\photo.png", [System.Drawing.Imaging.ImageFormat]::Png); $b.Dispose()

# 4) Handwritten signature (transparent background)
$b = NewBmp 600 160; $g = [System.Drawing.Graphics]::FromImage($b); $g.SmoothingMode = 'AntiAlias'; $g.TextRenderingHint = 'AntiAlias'
$g.Clear([System.Drawing.Color]::Transparent)
$f = New-Object System.Drawing.Font('Segoe Script', 30, [System.Drawing.FontStyle]::Italic)
$g.DrawString('Helen Marsh', $f, (New-Object System.Drawing.SolidBrush (Col 20 40 110)), 10, 20)
$g.DrawBezier((New-Object System.Drawing.Pen((Col 20 40 110), 3)), 20, 135, 200, 120, 380, 150, 560, 118)
$g.Dispose(); $b.Save("$Out\signature.png", [System.Drawing.Imaging.ImageFormat]::Png); $b.Dispose()
'ok'
