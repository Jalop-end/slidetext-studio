Add-Type -AssemblyName System.Drawing

$size = 256
$bmp = New-Object System.Drawing.Bitmap $size, $size
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

# Clear background (transparent)
$g.Clear([System.Drawing.Color]::Transparent)

# Draw rounded rectangle background with gradient (amber-600 to orange-500)
$rect = New-Object System.Drawing.Rectangle 12, 12, 232, 232
$brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
    (New-Object System.Drawing.Point 12, 244),
    (New-Object System.Drawing.Point 244, 12),
    [System.Drawing.Color]::FromArgb(217, 119, 6),   # amber-600
    [System.Drawing.Color]::FromArgb(249, 115, 22)   # orange-500
)

# Function to create rounded rectangle path
function Get-RoundedRectanglePath([System.Drawing.Rectangle]$r, [int]$radius) {
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $radius * 2
    $path.AddArc($r.X, $r.Y, $d, $d, 180, 90)
    $path.AddArc($r.X + $r.Width - $d, $r.Y, $d, $d, 270, 90)
    $path.AddArc($r.X + $r.Width - $d, $r.Y + $r.Height - $d, $d, $d, 0, 90)
    $path.AddArc($r.X, $r.Y + $r.Height - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    return $path
}

$path = Get-RoundedRectanglePath $rect 48
$g.FillPath($brush, $path)

# Subtle shadow / border
$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(60, 255, 255, 255), 4)
$g.DrawPath($pen, $path)

# Draw inner presentation slide sheet (white card)
$slideRect = New-Object System.Drawing.Rectangle 48, 48, 160, 115
$slidePath = Get-RoundedRectanglePath $slideRect 16
$whiteBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
$g.FillPath($whiteBrush, $slidePath)

# Slide header line (orange)
$headerPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(234, 88, 12), 12)
$headerPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$headerPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$g.DrawLine($headerPen, 72, 75, 120, 75)

# Slide text lines (slate/gray)
$linePen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(148, 163, 184), 8)
$linePen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$linePen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$g.DrawLine($linePen, 72, 102, 180, 102)
$g.DrawLine($linePen, 72, 122, 160, 122)
$g.DrawLine($linePen, 72, 142, 140, 142)

# Small AI sparkles / badge on bottom right
$badgeBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
    (New-Object System.Drawing.Point 145, 145),
    (New-Object System.Drawing.Point 215, 215),
    [System.Drawing.Color]::FromArgb(99, 102, 241),  # indigo-500
    [System.Drawing.Color]::FromArgb(168, 85, 247)  # purple-500
)
$badgeRect = New-Object System.Drawing.Rectangle 140, 140, 72, 72
$badgePath = Get-RoundedRectanglePath $badgeRect 20
$g.FillPath($badgeBrush, $badgePath)
$badgePen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 3)
$g.DrawPath($badgePen, $badgePath)

# Draw letter "AI" or Star in badge
$font = New-Object System.Drawing.Font("Segoe UI", 24, [System.Drawing.FontStyle]::Bold)
$textBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
$sf = New-Object System.Drawing.StringFormat
$sf.Alignment = [System.Drawing.StringAlignment]::Center
$sf.LineAlignment = [System.Drawing.StringAlignment]::Center
$g.DrawString("AI", $font, $textBrush, (New-Object System.Drawing.RectangleF 140, 140, 72, 72), $sf)

# Clean up graphics
$g.Flush()

# Convert bitmap to true ICO
$icoPath = "$PSScriptRoot\..\app.ico"
$icon = [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
$fs = [System.IO.File]::OpenWrite($icoPath)
$icon.Save($fs)
$fs.Close()
$icon.Dispose()
$bmp.Dispose()
$g.Dispose()

Write-Host "Icon created successfully: $icoPath"
