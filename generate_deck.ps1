[CmdletBinding()]
param (
    [string]$ConfigJsonPath,
    [string]$OutputPath
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
$OutputEncoding           = [System.Text.Encoding]::UTF8

$baseDir = "C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator"
$logoPath = "D:\For work\Logo\logo-ngang.png"
if (-not (Test-Path $logoPath)) {
    $logoPath = Join-Path $baseDir "ui\assets\logo.png"
}

if (-not (Test-Path $ConfigJsonPath)) {
    Write-Error "Lỗi: Không tìm thấy file cấu hình: $ConfigJsonPath"
    exit 1
}

$config = Get-Content $ConfigJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json

$localTemplate = Join-Path $baseDir "template\template.pptx"
$templatePptx = "D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\001\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi.pptx"
if (-not (Test-Path $templatePptx)) {
    $templatePptx = "D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\Ban đầu\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi.pptx"
}
if (-not (Test-Path $templatePptx)) {
    $templatePptx = $localTemplate
}

if (-not (Test-Path $templatePptx)) {
    Write-Error "Lỗi: Không tìm thấy file mẫu PowerPoint tại: $templatePptx"
    exit 1
}

if (-not $OutputPath) {
    if ($config.OutputPath) {
        $OutputPath = $config.OutputPath
    } else {
        $cleanName = if ($config.PatientName) { ($config.PatientName -replace '[\\/:*?"<>|]', '_').Trim() } else { "Bệnh Nhân" }
        $cleanYear = if ($config.BirthYear) { ($config.BirthYear -replace '[\\/:*?"<>|]', '_').Trim() } else { "1996" }
        $cleanDoc  = if ($config.DoctorName) { ($config.DoctorName -replace '[\\/:*?"<>|]', '_').Trim() } else { "Bác Sĩ" }
        
        $outFolder = if ($config.OutputFolder -and (Test-Path -LiteralPath $config.OutputFolder)) {
            $config.OutputFolder
        } else {
            $defaultDownloads = Join-Path $env:USERPROFILE "Downloads"
            if (Test-Path -LiteralPath $defaultDownloads) { $defaultDownloads } else { Join-Path $baseDir "output" }
        }
        $OutputPath = Join-Path $outFolder "$cleanName - $cleanYear - $cleanDoc.pptx"
    }
}

$outDir = [System.IO.Path]::GetDirectoryName($OutputPath)
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

# Check if target file is locked by PowerPoint
function Test-FileLocked([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return $false }
    try {
        $stream = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        if ($stream) { $stream.Close() }
        return $false
    } catch {
        return $true
    }
}

if (Test-FileLocked $OutputPath) {
    $dir = [System.IO.Path]::GetDirectoryName($OutputPath)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($OutputPath)
    $ext = [System.IO.Path]::GetExtension($OutputPath)
    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $OutputPath = Join-Path $dir "$name ($timestamp)$ext"
    Write-Host "File đích đang mở trong PowerPoint, đã tự động đổi tên thành: $OutputPath"
}

# Clean up any lingering PowerPoint processes
Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 1

Copy-Item -Path $templatePptx -Destination $OutputPath -Force
Write-Host "Copied template to: $OutputPath"

$ppt = New-Object -ComObject PowerPoint.Application
Start-Sleep -Seconds 1

try {
    # Open headless
    $pres = $ppt.Presentations.Open($OutputPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    
    # Colors (BGR for PowerPoint COM)
    $cWhite     = 0xFFFFFF
    $cNavy      = 0x663F0E   # #0E3F66
    $cGreen     = 0x59A800   # #00A859
    $cDarkGreen = 0x2A7300   # #00732A
    $cMint      = 0xEEF8E6   # #E6F8EE
    $cMuted     = 0x8B7464   # #64748B
    $cBorder    = 0xECD2D9   # #D9E2EC
    $cBodyText  = 0x334155   # Slate 700

    $pName   = if ($config.PatientName) { $config.PatientName } else { "Nguyễn Vũ Bảo Ngọc" }
    $pYear   = if ($config.BirthYear) { $config.BirthYear } else { "1996" }
    $pDoc    = if ($config.DoctorName) { $config.DoctorName } else { "Dr. Lê Võ Yến Nhi" }
    $pBranch = if ($config.ClinicBranch) { $config.ClinicBranch } else { "Cơ sở Tp. Hồ Chí Minh" }
    $pTagText = "$pName - $pYear"

    # ====================================================
    # 1. STYLE SLIDE 1: COVER
    # ====================================================
    Write-Host "Styling Cover Slide 1..."
    $s1 = $pres.Slides.Item(1)
    $s1.Background.Fill.Solid()
    $s1.Background.Fill.ForeColor.RGB = $cWhite

    foreach ($sh in $s1.Shapes) {
        if ($sh.HasTextFrame -eq [Microsoft.Office.Core.MsoTriState]::msoTrue) {
            $sh.TextFrame.TextRange.Font.Name = "Segoe UI"
            if ($sh.TextFrame.TextRange.Text -match "Nguyễn Vũ Bảo Ngọc|Bảo Ngọc") {
                $sh.TextFrame.TextRange.Text = "$pName - $pYear"
                $sh.TextFrame.TextRange.Font.Color.RGB = $cNavy
                $sh.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
                $sh.TextFrame.TextRange.Font.Size = 40
                $sh.Left = 60; $sh.Top = 175; $sh.Width = 840; $sh.Height = 65
                $sh.TextFrame.TextRange.ParagraphFormat.Alignment = 1
            } elseif ($sh.TextFrame.TextRange.Text -match "Dr\.|Cơ sở") {
                $sh.TextFrame.TextRange.Text = "$pDoc`r`n$pBranch"
                $sh.TextFrame.TextRange.Font.Color.RGB = $cMuted
                $sh.TextFrame.TextRange.Font.Size = 18
                $sh.Left = 60; $sh.Top = 255; $sh.Width = 840; $sh.Height = 85
                $sh.TextFrame.TextRange.ParagraphFormat.Alignment = 1
            }
        }
    }

    # Category Pill
    $covTag = $s1.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle, 60, 130, 290, 30)
    $covTag.Name = "BRAND_COVTAG"
    $covTag.Fill.Solid(); $covTag.Fill.ForeColor.RGB = $cMint
    $covTag.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $covTag.TextFrame.TextRange.Text = "HỒ SƠ BỆNH ÁN KHỚP THÁI DƯƠNG HÀM"
    $covTag.TextFrame.TextRange.Font.Name = "Segoe UI"
    $covTag.TextFrame.TextRange.Font.Size = 9.5
    $covTag.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $covTag.TextFrame.TextRange.Font.Color.RGB = $cDarkGreen
    $covTag.TextFrame.VerticalAnchor = [Microsoft.Office.Core.MsoVerticalAnchor]::msoAnchorMiddle

    if (Test-Path $logoPath) {
        $l1 = $s1.Shapes.AddPicture($logoPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 60, 45, 190, 48)
        $l1.Name = "BRAND_LOGO"
    }

    $covFoot = $s1.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 60, 485, 600, 25)
    $covFoot.Name = "BRAND_FOOT"
    $covFoot.TextFrame.TextRange.Text = "NHA KHOA THÙY ANH • HỆ THỐNG NHA KHOA CHUYÊN SÂU"
    $covFoot.TextFrame.TextRange.Font.Name = "Segoe UI"; $covFoot.TextFrame.TextRange.Font.Size = 9; $covFoot.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue; $covFoot.TextFrame.TextRange.Font.Color.RGB = 0x94A3B8

    # ====================================================
    # 2. STYLE SLIDE 2 & 3: INFO SLIDES
    # ====================================================
    foreach ($sNum in @(2, 3)) {
        $s = $pres.Slides.Item($sNum)
        $s.Background.Fill.Solid()
        $s.Background.Fill.ForeColor.RGB = $cWhite

        for ($i = 1; $i -le $s.Shapes.Count; $i++) {
            $sh = $s.Shapes.Item($i)
            if ($sh.HasTextFrame -eq [Microsoft.Office.Core.MsoTriState]::msoTrue) {
                $txt = $sh.TextFrame.TextRange.Text.Trim()
                $sh.TextFrame.TextRange.Font.Name = "Segoe UI"
                if ($sh.Name -like "Title*" -or $txt -match "Chief complain|Dental history") {
                    $sh.TextFrame.TextRange.Font.Color.RGB = $cNavy
                    $sh.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
                    $sh.TextFrame.TextRange.Font.Size = 20
                    $sh.Left = 175; $sh.Top = 13; $sh.Width = 500; $sh.Height = 32
                    $sh.TextFrame.MarginLeft = 0; $sh.TextFrame.MarginTop = 0
                } else {
                    $sh.TextFrame.TextRange.Font.Color.RGB = $cBodyText
                    $sh.TextFrame.TextRange.Font.Size = 15
                    $sh.Left = 60; $sh.Top = 75; $sh.Width = 840; $sh.Height = 400
                    if ($sNum -eq 2 -and $config.ChiefComplaint) {
                        $sh.TextFrame.TextRange.Text = $config.ChiefComplaint
                    }
                    if ($sNum -eq 3 -and $config.DentalHistory) {
                        $sh.TextFrame.TextRange.Text = $config.DentalHistory
                    }
                }
            }
        }

        # Logo
        if (Test-Path $logoPath) {
            $lPic = $s.Shapes.AddPicture($logoPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 25, 14, 130, 30)
            $lPic.Name = "BRAND_LOGO"
        }

        # Top hairline
        $line = $s.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRectangle, 25, 48, 910, 1)
        $line.Name = "BRAND_LINE"
        $line.Fill.Solid(); $line.Fill.ForeColor.RGB = 0xE2E8F0
        $line.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse

        # Footer
        $fL = $s.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 25, 514, 500, 18)
        $fL.Name = "BRAND_FL"
        $fL.TextFrame.TextRange.Text = "NHA KHOA THÙY ANH • CHUYÊN KHOA KHỚP THÁI DƯƠNG HÀM (TMD)"
        $fL.TextFrame.TextRange.Font.Name = "Segoe UI"; $fL.TextFrame.TextRange.Font.Size = 8; $fL.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue; $fL.TextFrame.TextRange.Font.Color.RGB = 0x94A3B8

        $fR = $s.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 830, 514, 105, 18)
        $fR.Name = "BRAND_FR"
        $fR.TextFrame.TextRange.Text = "Trang " + $sNum.ToString("00") + " / 31"
        $fR.TextFrame.TextRange.Font.Name = "Segoe UI"; $fR.TextFrame.TextRange.Font.Size = 8.5; $fR.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue; $fR.TextFrame.TextRange.Font.Color.RGB = 0x64748B
        $fR.TextFrame.TextRange.ParagraphFormat.Alignment = 3
    }

    # ====================================================
    # 3. HELPER: GET PATIENT PICTURES IN A SLIDE
    # ====================================================
    function Get-PatientPics($slide) {
        $res = @()
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $sh = $slide.Shapes.Item($i)
            if (($sh.Type -in @(11, 13, 16)) -and ($sh.Name -notlike "BRAND_*") -and ($sh.Top -ge 50)) {
                $res += $sh
            }
        }
        $sorted = @($res | Sort-Object { [double]$_.Left })
        return ,$sorted
    }

    function Format-PicShape($pic) {
        try {
            $pic.AutoShapeType = 5 # msoShapeRoundedRectangle
            if ($pic.Adjustments.Count -ge 1) {
                $pic.Adjustments.Item(1) = 0.04
            }
            $pic.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
            $pic.Line.ForeColor.RGB = 0xD9E2EC
            $pic.Line.Weight = 1.2
        } catch {}
    }

    function Replace-Pic($slide, $shape, $path, $slotId = "") {
        $l = $shape.Left; $t = $shape.Top; $w = $shape.Width; $h = $shape.Height
        $shape.Delete()

        # Load image at native dimensions to inspect natural aspect ratio
        $p = $slide.Shapes.AddPicture($path, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, $l, $t, -1, -1)
        if ($slotId) {
            $p.Name = "PATIENT_PIC_$slotId"
        } else {
            $p.Name = "PATIENT_PIC"
        }
        $p.LockAspectRatio = [Microsoft.Office.Core.MsoTriState]::msoTrue

        $origW = $p.Width
        $origH = $p.Height
        if ($origW -gt 0 -and $origH -gt 0) {
            # Scale strictly within allotted bounding box without distortion
            $scale = [math]::Min($w / $origW, $h / $origH)
            $newW = $origW * $scale
            $newH = $origH * $scale
            $newL = $l + ($w - $newW) / 2
            $newT = $t + ($h - $newH) / 2

            $p.Width = $newW
            $p.Height = $newH
            $p.Left = $newL
            $p.Top = $newT
        }
        Format-PicShape $p
        return $p
    }

    function Process-SlotShape($slide, $shape, $slotId) {
        $slotKey = [string]$slotId
        if ($slotFiles.ContainsKey($slotKey) -and (Test-Path -LiteralPath $slotFiles[$slotKey])) {
            Replace-Pic $slide $shape $slotFiles[$slotKey] $slotKey | Out-Null
        } else {
            # Empty slot: delete template picture so this position is completely left blank!
            $shape.Delete()
        }
    }

    # ====================================================
    # 4. MAP AND REPLACE PHOTOS BY SLOT
    # ====================================================
    $slotFiles = @{}
    if ($config.SlotFiles) {
        foreach ($prop in $config.SlotFiles.PSObject.Properties) {
            if ($prop.Value -and (Test-Path -LiteralPath $prop.Value)) {
                $slotFiles[$prop.Name] = $prop.Value
            }
        }
    }

    Write-Host "Replacing uploaded photos by slot (preserving aspect ratio & clearing empty slots)..."

    # Slide 4: 1.jpg, 2.jpg
    $s = $pres.Slides.Item(4); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '1' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '2' }

    # Slide 5: 3.jpg, 4.jpg, 5.jpg (5 = Nghiêng phải)
    $s = $pres.Slides.Item(5); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '3' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '4' }
    if ($p.Count -gt 2) { Process-SlotShape $s $p[2] '5' }

    # Slide 6: 5.jpg (TỰ ĐỘNG ĐỒNG BỘ TỪ 5.JPG - Bên trái), 6.jpg (Nghiêng trái - Bên phải)
    $s = $pres.Slides.Item(6); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '5' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '6' }

    # Slide 7: 7.jpg, 8.jpg
    $s = $pres.Slides.Item(7); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '7' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '8' }

    # Slide 8: 9.jpg, 10.jpg
    $s = $pres.Slides.Item(8); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '9' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '10' }

    # Slide 9: 11.jpg, 12.jpg
    $s = $pres.Slides.Item(9); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '11' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '12' }

    # Slide 10: 13.jpg, 14.jpg
    $s = $pres.Slides.Item(10); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '13' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '14' }

    # Slide 11: 15.jpg, 16.jpg
    $s = $pres.Slides.Item(11); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '15' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '16' }

    # Slide 12: 17.jpg (Mặt trong bên trái), 18.jpg (Mặt trong bên phải)
    $s = $pres.Slides.Item(12); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '17' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '18' }

    # Slide 13: 13.jpg (TỰ ĐỘNG ĐỒNG BỘ), 19.jpg
    $s = $pres.Slides.Item(13); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '13' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '19' }

    # Slide 14: 20.jpg, 21.jpg
    $s = $pres.Slides.Item(14); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '20' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '21' }

    # Slide 15: 22.jpg, 23.jpg
    $s = $pres.Slides.Item(15); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '22' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '23' }

    # Slide 16: Bên trái trượt phải (25.jpg), Trực diện trượt phải (24.jpg), Bên phải trượt phải (ĐỒNG BỘ 23.jpg)
    $s = $pres.Slides.Item(16); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '25' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '24' }
    if ($p.Count -gt 2) { Process-SlotShape $s $p[2] '23' }

    # Slide 17: Bên trái trượt trái (ĐỒNG BỘ 22.jpg), Trực diện trượt trái (26.jpg), Bên phải trượt trái (27.jpg)
    $s = $pres.Slides.Item(17); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '22' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '26' }
    if ($p.Count -gt 2) { Process-SlotShape $s $p[2] '27' }

    # Slide 18: 28.jpg, 29.jpg
    $s = $pres.Slides.Item(18); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '28' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '29' }

    # Slide 19: 30.jpg (Cung trên qua gương)
    $s = $pres.Slides.Item(19); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '30' }

    # Slide 20: 31.jpg (Cung dưới qua gương)
    $s = $pres.Slides.Item(20); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '31' }

    # Slide 21: 32.mp4 hoặc 32.jpg
    $s21 = $pres.Slides.Item(21)
    $mediaShape = $null
    for ($i = 1; $i -le $s21.Shapes.Count; $i++) {
        $sh = $s21.Shapes.Item($i)
        if ($sh.Type -in @(11, 13, 16) -and $sh.Top -ge 50) {
            $mediaShape = $sh
            break
        }
    }
    if ($mediaShape) {
        if ($slotFiles.ContainsKey('32') -and (Test-Path -LiteralPath $slotFiles['32'])) {
            $f32 = $slotFiles['32']
            $ext = [System.IO.Path]::GetExtension($f32).ToLower()
            $l = $mediaShape.Left; $t = $mediaShape.Top; $w = $mediaShape.Width; $h = $mediaShape.Height
            $mediaShape.Delete()
            if ($ext -in @('.mp4', '.avi', '.mov', '.wmv')) {
                $mObj = $s21.Shapes.AddMediaObject2($f32, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, $l, $t, $w, $h)
                $mObj.Name = "PATIENT_PIC_32"
            } else {
                $p = $s21.Shapes.AddPicture($f32, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, $l, $t, -1, -1)
                $p.Name = "PATIENT_PIC_32"
                $p.LockAspectRatio = [Microsoft.Office.Core.MsoTriState]::msoTrue
                $scale = [math]::Min($w / $p.Width, $h / $p.Height)
                $newW = $p.Width * $scale; $newH = $p.Height * $scale
                $p.Width = $newW; $p.Height = $newH
                $p.Left = $l + ($w - $newW) / 2; $p.Top = $t + ($h - $newH) / 2
                Format-PicShape $p
            }
        } else {
            $mediaShape.Delete()
        }
    }

    # Slide 22: Composite 1 (11 ảnh)
    $s22 = $pres.Slides.Item(22)
    foreach ($sh in @($s22.Shapes)) {
        if ($sh.Type -in 11, 13) {
            $l = [math]::Round($sh.Left); $t = [math]::Round($sh.Top)
            if ($t -ge 50 -and $t -le 80) {
                if ($l -le 250) { Process-SlotShape $s22 $sh '5' }
                elseif ($l -gt 250 -and $l -le 420) { Process-SlotShape $s22 $sh '3' }
                elseif ($l -gt 420 -and $l -le 580) { Process-SlotShape $s22 $sh '4' }
                elseif ($l -gt 580) { Process-SlotShape $s22 $sh '33' }
            } elseif ($l -ge 350 -and $l -le 450) {
                if ($t -le 280) { Process-SlotShape $s22 $sh '13' }
                elseif ($t -gt 280 -and $t -le 380) { Process-SlotShape $s22 $sh '14' }
                elseif ($t -gt 380) { Process-SlotShape $s22 $sh '28' }
            } elseif ($t -ge 200 -and $t -le 300) {
                if ($l -le 300) { Process-SlotShape $s22 $sh '30' }
                elseif ($l -gt 500) { Process-SlotShape $s22 $sh '31' }
            } elseif ($t -gt 380) {
                if ($l -le 300) { Process-SlotShape $s22 $sh '15' }
                elseif ($l -gt 500) { Process-SlotShape $s22 $sh '16' }
            }
        }
    }

    # Slide 23: Composite 2 (11 ảnh)
    $s23 = $pres.Slides.Item(23)
    foreach ($sh in @($s23.Shapes)) {
        if ($sh.Type -in 11, 13) {
            $l = [math]::Round($sh.Left); $t = [math]::Round($sh.Top)
            if ($t -ge 45 -and $t -le 75) {
                if ($l -le 250) { Process-SlotShape $s23 $sh '5' }
                elseif ($l -gt 250 -and $l -le 420) { Process-SlotShape $s23 $sh '6' }
                elseif ($l -gt 420 -and $l -le 580) { Process-SlotShape $s23 $sh '34' }
                elseif ($l -gt 580) { Process-SlotShape $s23 $sh '35' }
            } elseif ($l -ge 350 -and $l -le 450) {
                if ($t -le 280) { Process-SlotShape $s23 $sh '19' }
                elseif ($t -gt 280 -and $t -le 380) { Process-SlotShape $s23 $sh '24' }
                elseif ($t -gt 380) { Process-SlotShape $s23 $sh '26' }
            } elseif ($t -ge 170 -and $t -le 270) {
                if ($l -le 300) { Process-SlotShape $s23 $sh '17' }
                elseif ($l -gt 500) { Process-SlotShape $s23 $sh '18' }
            } elseif ($t -gt 380) {
                if ($l -le 300) { Process-SlotShape $s23 $sh '20' }
                elseif ($l -gt 500) { Process-SlotShape $s23 $sh '21' }
            }
        }
    }

    # Slide 24: 9 ảnh đường thở & khuôn miệng
    $s24 = $pres.Slides.Item(24)
    foreach ($sh in @($s24.Shapes)) {
        if ($sh.Type -in 11, 13) {
            $l = [math]::Round($sh.Left); $t = [math]::Round($sh.Top)
            if ($t -le 150) {
                if ($l -le 250) { Process-SlotShape $s24 $sh '40' }
                elseif ($l -gt 250 -and $l -le 500) { Process-SlotShape $s24 $sh '42' }
                elseif ($l -gt 500) { Process-SlotShape $s24 $sh '41' }
            } elseif ($t -gt 150 -and $t -le 320) {
                if ($l -le 250) { Process-SlotShape $s24 $sh '36' }
                elseif ($l -gt 250 -and $l -le 500) { Process-SlotShape $s24 $sh '37' }
                elseif ($l -gt 500) { Process-SlotShape $s24 $sh '38' }
            } elseif ($t -gt 320) {
                if ($l -le 250) { Process-SlotShape $s24 $sh '28' }
                elseif ($l -gt 250 -and $l -le 500) { Process-SlotShape $s24 $sh '29' }
                elseif ($l -gt 500) { Process-SlotShape $s24 $sh '39' }
            }
        }
    }

    # Slide 25 - 31 (Phim và đo đạc)
    $s = $pres.Slides.Item(25); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '40' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '41' }

    $s = $pres.Slides.Item(26); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '42' }

    $s = $pres.Slides.Item(27); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '43' }

    $s = $pres.Slides.Item(28); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '44' }

    $s = $pres.Slides.Item(29); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '45' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '46' }
    if ($p.Count -gt 2) { Process-SlotShape $s $p[2] '47' }

    $s = $pres.Slides.Item(30); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '48' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '49' }
    if ($p.Count -gt 2) { Process-SlotShape $s $p[2] '50' }

    $s = $pres.Slides.Item(31); $p = (Get-PatientPics $s)
    if ($p.Count -gt 0) { Process-SlotShape $s $p[0] '51' }
    if ($p.Count -gt 1) { Process-SlotShape $s $p[1] '52' }

    # ====================================================
    # 4.5. STRICT CLEANUP: PURGE ANY REMAINING TEMPLATE PHOTOS
    #    "tuyệt đối không sử dụng lại bất kì ảnh nào của slide mẫu.
    #     Nếu ô ảnh đưa vào trống thì hãy để vị trí tương ứng ấy trống không"
    # ====================================================
    Write-Host "Strict cleanup: Purging any unreplaced template photos from all slides..."
    for ($sNum = 4; $sNum -le 31; $sNum++) {
        $s = $pres.Slides.Item($sNum)
        $toDelete = @()
        for ($i = 1; $i -le $s.Shapes.Count; $i++) {
            $sh = $s.Shapes.Item($i)
            if ($sh.Type -in @(11, 13, 16) -and $sh.Name -notlike "BRAND_*" -and $sh.Name -notlike "PATIENT_PIC_*") {
                $toDelete += $sh
            }
        }
        foreach ($dSh in $toDelete) {
            try { $dSh.Delete() } catch {}
        }
    }

    # ====================================================
    # 5. AESTHETIC POLISH ON ALL CONTENT SLIDES 4..31
    #    (Preserve exact photo aspect ratio & coordinates!)
    # ====================================================
    Write-Host "Polishing typography and photo frames on slides 4..31..."
    for ($sNum = 4; $sNum -le 31; $sNum++) {
        $s = $pres.Slides.Item($sNum)
        $s.Background.Fill.Solid()
        $s.Background.Fill.ForeColor.RGB = $cWhite

        # 1. Format existing Slide Title text box
        for ($i = 1; $i -le $s.Shapes.Count; $i++) {
            $sh = $s.Shapes.Item($i)
            if ($sh.HasTextFrame -eq [Microsoft.Office.Core.MsoTriState]::msoTrue -and ($sh.Name -like "Title*" -or $sh.Type -eq 14)) {
                $sh.TextFrame.TextRange.Font.Name = "Segoe UI"
                $sh.TextFrame.TextRange.Font.Size = 16
                $sh.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
                $sh.TextFrame.TextRange.Font.Color.RGB = $cNavy
                $sh.Left = 175
                $sh.Top = 13
                $sh.Width = 500
                $sh.Height = 32
                $sh.TextFrame.MarginLeft = 0; $sh.TextFrame.MarginTop = 0
            }
        }

        # 2. Format ALL pictures to rounded corners & elegant border
        for ($i = 1; $i -le $s.Shapes.Count; $i++) {
            $sh = $s.Shapes.Item($i)
            if ($sh.Type -in 11, 13) {
                if ($sh.Top -ge 50 -and $sh.Name -notlike "BRAND_*") {
                    Format-PicShape $sh
                }
            }
        }

        # 3. Logo top-left
        if (Test-Path $logoPath) {
            $lPic = $s.Shapes.AddPicture($logoPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 25, 14, 130, 30)
            $lPic.Name = "BRAND_LOGO"
        }

        # 4. Top hairline divider
        $line = $s.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRectangle, 25, 48, 910, 1)
        $line.Name = "BRAND_LINE"
        $line.Fill.Solid(); $line.Fill.ForeColor.RGB = 0xE2E8F0
        $line.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse

        # 5. Patient Tag (top right)
        $pTag = $s.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 680, 16, 255, 24)
        $pTag.Name = "BRAND_PTAG"
        $pTag.TextFrame.TextRange.Text = $pTagText
        $pTag.TextFrame.TextRange.Font.Name = "Segoe UI"
        $pTag.TextFrame.TextRange.Font.Size = 9.5
        $pTag.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
        $pTag.TextFrame.TextRange.Font.Color.RGB = 0x64748B
        $pTag.TextFrame.TextRange.ParagraphFormat.Alignment = 3

        # 6. Subtle Footer
        $fL = $s.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 25, 514, 500, 18)
        $fL.Name = "BRAND_FL"
        $fL.TextFrame.TextRange.Text = "NHA KHOA THÙY ANH • CHUYÊN KHOA KHỚP THÁI DƯƠNG HÀM (TMD)"
        $fL.TextFrame.TextRange.Font.Name = "Segoe UI"; $fL.TextFrame.TextRange.Font.Size = 8; $fL.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue; $fL.TextFrame.TextRange.Font.Color.RGB = 0x94A3B8

        $fR = $s.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 830, 514, 105, 18)
        $fR.Name = "BRAND_FR"
        $fR.TextFrame.TextRange.Text = "Trang " + $sNum.ToString("00") + " / 31"
        $fR.TextFrame.TextRange.Font.Name = "Segoe UI"; $fR.TextFrame.TextRange.Font.Size = 8.5; $fR.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue; $fR.TextFrame.TextRange.Font.Color.RGB = 0x64748B
        $fR.TextFrame.TextRange.ParagraphFormat.Alignment = 3
    }

    # Save Presentation
    $pres.Save()
    $pres.Close()
    Write-Host "FINAL_OUTPUT_PATH:$OutputPath"
    Write-Host "SUCCESS: Redesigned presentation created at: $OutputPath"
}
catch {
    $errMsg = "Lỗi khi tạo slide: $($_.Exception.Message)"
    if ($_.InvocationInfo) {
        $errMsg += " (tại dòng $($_.InvocationInfo.ScriptLineNumber))"
    }
    Write-Host "ERROR_DETAIL:$errMsg" -ForegroundColor Red
    Write-Error $errMsg
    exit 1
}
finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
