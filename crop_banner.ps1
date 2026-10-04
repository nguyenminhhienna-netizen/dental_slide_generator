Add-Type -AssemblyName System.Drawing
$brainDir = "C:\Users\PC\.gemini\antigravity\brain\acfc5fe9-e6b9-4baf-ae0e-4fcb79847721"
$bannerPath = Join-Path $brainDir ".user_uploaded\media_1789790186457.png"
$cropPath = Join-Path $brainDir "banner_art_right.png"

$img = [System.Drawing.Image]::FromFile($bannerPath)
$w = $img.Width
$h = $img.Height
Write-Host "Original banner size: $w x $h"

# Crop the right 65% of the image (starts at x = w * 0.35)
$cropX = [int]($w * 0.38)
$cropW = $w - $cropX
$cropRect = New-Object System.Drawing.Rectangle($cropX, 0, $cropW, $h)

$bmp = New-Object System.Drawing.Bitmap($cropW, $h)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.DrawImage($img, (New-Object System.Drawing.Rectangle(0, 0, $cropW, $h)), $cropRect, [System.Drawing.GraphicsUnit]::Pixel)
$g.Dispose()
$img.Dispose()

$bmp.Save($cropPath, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
Write-Host "Saved cropped right art to: $cropPath"
