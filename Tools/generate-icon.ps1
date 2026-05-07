# Generates a multi-resolution .ico file for OOF Manager.
# Theme: blue rounded-square background + white envelope + accent dot (matches the email/OOF concept).
# Output: Resources\oofmanager.ico

Add-Type -AssemblyName System.Drawing

function New-EnvelopeBitmap([int]$size) {
    $bmp = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $g.Clear([System.Drawing.Color]::Transparent)

    # --- Rounded-square accent background ---
    $cornerRadius = [int]($size * 0.20)
    $bgRect = New-Object System.Drawing.Rectangle(0, 0, $size, $size)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $cornerRadius * 2
    $path.AddArc($bgRect.X, $bgRect.Y, $d, $d, 180, 90)
    $path.AddArc($bgRect.Right - $d, $bgRect.Y, $d, $d, 270, 90)
    $path.AddArc($bgRect.Right - $d, $bgRect.Bottom - $d, $d, $d, 0, 90)
    $path.AddArc($bgRect.X, $bgRect.Bottom - $d, $d, $d, 90, 90)
    $path.CloseFigure()

    # Vertical gradient (Microsoft accent blue family)
    $brushTop = [System.Drawing.Color]::FromArgb(255, 0, 120, 215)
    $brushBot = [System.Drawing.Color]::FromArgb(255, 0, 90, 175)
    $bgBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($bgRect, $brushTop, $brushBot, 90.0)
    $g.FillPath($bgBrush, $path)
    $bgBrush.Dispose()

    # --- Envelope ---
    $pad = [int]($size * 0.18)
    $envX = $pad
    $envY = [int]($size * 0.28)
    $envW = $size - 2 * $pad
    $envH = [int]($size * 0.46)
    $envRect = New-Object System.Drawing.RectangleF($envX, $envY, $envW, $envH)

    $envBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $envPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 0, 70, 140), [Math]::Max(1.0, $size / 64.0))
    $envPen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round

    # Body
    $g.FillRectangle($envBrush, $envRect)
    $g.DrawRectangle($envPen, $envRect.X, $envRect.Y, $envRect.Width, $envRect.Height)

    # Flap (V shape)
    $flap = New-Object 'System.Drawing.PointF[]' 3
    $flap[0] = New-Object System.Drawing.PointF($envRect.X, $envRect.Y)
    $flap[1] = New-Object System.Drawing.PointF(($envRect.X + $envRect.Width / 2.0), ($envRect.Y + $envRect.Height * 0.55))
    $flap[2] = New-Object System.Drawing.PointF(($envRect.X + $envRect.Width), $envRect.Y)
    $g.DrawLines($envPen, $flap)

    # --- Accent "@" badge (small red dot indicating notification / OOF active) ---
    $dotR = [int]($size * 0.16)
    $dotX = $size - $dotR - [int]($size * 0.06)
    $dotY = [int]($size * 0.06)
    $dotRect = New-Object System.Drawing.Rectangle($dotX, $dotY, $dotR, $dotR)
    $dotBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 232, 17, 35))
    $dotPen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, [Math]::Max(1.0, $size / 32.0))
    $g.FillEllipse($dotBrush, $dotRect)
    $g.DrawEllipse($dotPen, $dotRect)

    $envBrush.Dispose()
    $envPen.Dispose()
    $dotBrush.Dispose()
    $dotPen.Dispose()
    $g.Dispose()
    return $bmp
}

function Save-IcoFromBitmaps([System.Drawing.Bitmap[]] $bitmaps, [string] $outPath) {
    # Encode each bitmap as PNG, then assemble ICONDIR + ICONDIRENTRY[] + image data.
    $pngStreams = @()
    foreach ($bmp in $bitmaps) {
        $ms = New-Object System.IO.MemoryStream
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        $pngStreams += ,@{ Bytes = $ms.ToArray(); Width = $bmp.Width; Height = $bmp.Height }
        $ms.Dispose()
    }

    $count = $pngStreams.Count
    $headerSize = 6
    $entrySize = 16
    $offset = $headerSize + ($entrySize * $count)

    $out = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($out)

    # ICONDIR
    $bw.Write([UInt16]0)        # reserved
    $bw.Write([UInt16]1)        # type: icon
    $bw.Write([UInt16]$count)   # number of images

    foreach ($p in $pngStreams) {
        $w = if ($p.Width -ge 256) { 0 } else { [byte]$p.Width }
        $h = if ($p.Height -ge 256) { 0 } else { [byte]$p.Height }
        $bw.Write([byte]$w)
        $bw.Write([byte]$h)
        $bw.Write([byte]0)            # color count
        $bw.Write([byte]0)            # reserved
        $bw.Write([UInt16]1)          # planes
        $bw.Write([UInt16]32)         # bits per pixel
        $bw.Write([UInt32]$p.Bytes.Length)
        $bw.Write([UInt32]$offset)
        $offset += $p.Bytes.Length
    }

    foreach ($p in $pngStreams) {
        $bw.Write($p.Bytes)
    }

    $bw.Flush()
    [System.IO.File]::WriteAllBytes($outPath, $out.ToArray())
    $bw.Dispose()
    $out.Dispose()
}

$sizes = @(16, 24, 32, 48, 64, 128, 256)
$bitmaps = @()
foreach ($s in $sizes) { $bitmaps += ,(New-EnvelopeBitmap $s) }

$resourcesDir = Join-Path $PSScriptRoot "..\Resources"
if (-not (Test-Path $resourcesDir)) { New-Item -ItemType Directory -Path $resourcesDir -Force | Out-Null }
$icoPath = Join-Path $resourcesDir "oofmanager.ico"
Save-IcoFromBitmaps $bitmaps $icoPath

foreach ($b in $bitmaps) { $b.Dispose() }

Write-Host "Wrote $icoPath ($((Get-Item $icoPath).Length) bytes, $($sizes.Count) sizes)"

# Also write a high-resolution standalone PNG used by the login page splash.
# WPF's Image control picks the closest-matching frame from a multi-resolution
# .ico and then scales — that's what made the 96px login icon look blurry on
# HiDPI displays. A dedicated 512x512 PNG lets WPF downsample from a single
# crisp source instead.
$splashSize = 512
$splashBmp = New-EnvelopeBitmap $splashSize
$splashPath = Join-Path $resourcesDir "oofmanager-512.png"
$splashBmp.Save($splashPath, [System.Drawing.Imaging.ImageFormat]::Png)
$splashBmp.Dispose()
Write-Host "Wrote $splashPath ($((Get-Item $splashPath).Length) bytes, ${splashSize}x${splashSize})"
