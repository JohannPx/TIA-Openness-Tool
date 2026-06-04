# GenerateIcon.ps1 - Generate icon.ico from WPF drawing (used by CI)
# Produces a multi-resolution .ico file matching the app's "DB" badge icon
# (same design as New-AppIcon in scripts/modules/UI.ps1).

param(
    [string]$OutputPath = "icon.ico"
)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

function Render-DbIcon([int]$size) {
    $scale = $size / 32.0
    $dv = New-Object System.Windows.Media.DrawingVisual
    $dc = $dv.RenderOpen()

    # Background: rounded rectangle (Siemens teal gradient)
    $bgRect = New-Object System.Windows.Rect((1 * $scale), (1 * $scale), (30 * $scale), (30 * $scale))
    $bgBrush = New-Object System.Windows.Media.LinearGradientBrush(
        [System.Windows.Media.Color]::FromRgb(0, 120, 136),
        [System.Windows.Media.Color]::FromRgb(0, 90, 100),
        45)
    $bgPen = New-Object System.Windows.Media.Pen(
        (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(0, 70, 80))), (0.5 * $scale))
    $dc.DrawRoundedRectangle($bgBrush, $bgPen, $bgRect, (4 * $scale), (4 * $scale))

    # "DB" text (bold, white, centered)
    $typeface = New-Object System.Windows.Media.Typeface(
        (New-Object System.Windows.Media.FontFamily("Segoe UI")),
        [System.Windows.FontStyles]::Normal,
        [System.Windows.FontWeights]::Bold,
        [System.Windows.FontStretches]::Normal)
    $formattedText = New-Object System.Windows.Media.FormattedText(
        "DB", [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Windows.FlowDirection]::LeftToRight,
        $typeface, (14 * $scale), [System.Windows.Media.Brushes]::White)
    $textX = ((32 * $scale) - $formattedText.Width) / 2
    $textY = ((32 * $scale) - $formattedText.Height) / 2 - (1 * $scale)
    $dc.DrawText($formattedText, (New-Object System.Windows.Point($textX, $textY)))

    # Small export arrow (bottom-right corner, gold)
    $arrowGeo = New-Object System.Windows.Media.StreamGeometry
    $ctx = $arrowGeo.Open()
    $ctx.BeginFigure((New-Object System.Windows.Point((22 * $scale), (23 * $scale))), $true, $true)
    $ctx.LineTo((New-Object System.Windows.Point((28 * $scale), (23 * $scale))), $true, $false)
    $ctx.LineTo((New-Object System.Windows.Point((25 * $scale), (28 * $scale))), $true, $false)
    $ctx.Close()
    $arrowBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255, 200, 0))
    $dc.DrawGeometry($arrowBrush, $null, $arrowGeo)

    $dc.Close()

    $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($size, $size, 96, 96,
        [System.Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($dv)
    $rtb.Freeze()
    return $rtb
}

function ConvertTo-IcoBytes([System.Windows.Media.Imaging.BitmapSource[]]$Bitmaps) {
    # ICO file format: header + directory entries + PNG image data
    $ms = New-Object System.IO.MemoryStream

    # ICO Header: reserved(2) + type(2) + count(2)
    $ms.Write([byte[]](0, 0), 0, 2)                          # Reserved
    $ms.Write([BitConverter]::GetBytes([uint16]1), 0, 2)      # Type: 1 = ICO
    $ms.Write([BitConverter]::GetBytes([uint16]$Bitmaps.Count), 0, 2)

    # Pre-render each bitmap to PNG bytes
    $pngDataList = @()
    foreach ($bmp in $Bitmaps) {
        $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
        $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bmp))
        $pngMs = New-Object System.IO.MemoryStream
        $encoder.Save($pngMs)
        $pngDataList += , $pngMs.ToArray()
        $pngMs.Dispose()
    }

    # Directory entries offset starts after header (6 bytes) + entries (16 bytes each)
    $dataOffset = 6 + ($Bitmaps.Count * 16)

    for ($i = 0; $i -lt $Bitmaps.Count; $i++) {
        $bmp = $Bitmaps[$i]
        $pngData = $pngDataList[$i]
        $w = [byte]$(if ($bmp.PixelWidth -ge 256) { 0 } else { $bmp.PixelWidth })
        $h = [byte]$(if ($bmp.PixelHeight -ge 256) { 0 } else { $bmp.PixelHeight })

        $ms.WriteByte($w)                                                  # Width
        $ms.WriteByte($h)                                                  # Height
        $ms.WriteByte(0)                                                   # Color palette
        $ms.WriteByte(0)                                                   # Reserved
        $ms.Write([BitConverter]::GetBytes([uint16]1), 0, 2)               # Color planes
        $ms.Write([BitConverter]::GetBytes([uint16]32), 0, 2)              # Bits per pixel
        $ms.Write([BitConverter]::GetBytes([uint32]$pngData.Length), 0, 4) # Image size
        $ms.Write([BitConverter]::GetBytes([uint32]$dataOffset), 0, 4)     # Offset

        $dataOffset += $pngData.Length
    }

    # Write PNG image data
    foreach ($pngData in $pngDataList) {
        $ms.Write($pngData, 0, $pngData.Length)
    }

    return $ms.ToArray()
}

# Generate icons at multiple sizes
$sizes = @(16, 32, 48, 256)
$bitmaps = @()
foreach ($s in $sizes) {
    $bitmaps += Render-DbIcon $s
}

$icoBytes = ConvertTo-IcoBytes $bitmaps
$resolvedPath = if ([System.IO.Path]::IsPathRooted($OutputPath)) { $OutputPath } else { Join-Path (Get-Location) $OutputPath }
[System.IO.File]::WriteAllBytes($resolvedPath, $icoBytes)

Write-Host "Icon generated: $resolvedPath ($($sizes -join ', ') px)"
