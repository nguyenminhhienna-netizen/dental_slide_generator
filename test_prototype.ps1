$templatePptx = "D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\001\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi.pptx"
if (-not (Test-Path $templatePptx)) {
    $templatePptx = "D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\Ban đầu\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi.pptx"
}

$logoPath = "D:\For work\Logo\logo-ngang.png"
$bannerArt = "C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\assets\banner_art_right.png"
if (-not (Test-Path $bannerArt)) {
    $bannerArt = "C:\Users\PC\.gemini\antigravity\brain\acfc5fe9-e6b9-4baf-ae0e-4fcb79847721\banner_art_right.png"
}

$testPptx = "C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\test_aesthetic.pptx"
Copy-Item $templatePptx $testPptx -Force

$ppt = New-Object -ComObject PowerPoint.Application
try {
    $pres = $ppt.Presentations.Open($testPptx, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    
    # BGR Colors
    $cWhite = 0xFFFFFF
    $cNavy = 0x663F0E        # #0E3F66
    $cGreen = 0x59A800       # #00A859
    $cMuted = 0x8B7464       # #64748B
    $cBorder = 0xECD2D9      # #D9E2EC in BGR
    $cCardBg = 0xFAFCFF
    
    # --- SLIDE 1: COVER ---
    $s1 = $pres.Slides.Item(1)
    $s1.Background.Fill.Solid()
    $s1.Background.Fill.ForeColor.RGB = $cWhite
    
    foreach ($sh in $s1.Shapes) {
        if ($sh.HasTextFrame -eq [Microsoft.Office.Core.MsoTriState]::msoTrue) {
            $sh.TextFrame.TextRange.Font.Name = "Segoe UI"
            if ($sh.TextFrame.TextRange.Text -match "Nguyễn Vũ Bảo Ngọc") {
                $sh.TextFrame.TextRange.Font.Color.RGB = $cNavy
                $sh.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
                $sh.TextFrame.TextRange.Font.Size = 40
                $sh.Left = 60; $sh.Top = 170; $sh.Width = 840; $sh.Height = 70
                $sh.TextFrame.TextRange.ParagraphFormat.Alignment = 1 # Left align
            } elseif ($sh.TextFrame.TextRange.Text -match "Dr\.|Cơ sở") {
                $sh.TextFrame.TextRange.Font.Color.RGB = 0x64748B
                $sh.TextFrame.TextRange.Font.Size = 18
                $sh.Left = 60; $sh.Top = 250; $sh.Width = 840; $sh.Height = 90
                $sh.TextFrame.TextRange.ParagraphFormat.Alignment = 1 # Left align
            }
        }
    }
    
    # Add Tag Pill above name
    $covTag = $s1.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle, 60, 125, 290, 30)
    $covTag.Fill.Solid(); $covTag.Fill.ForeColor.RGB = 0xEEF8E6 # mint
    $covTag.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $covTag.TextFrame.TextRange.Text = "HỒ SƠ BỆNH ÁN KHỚP THÁI DƯƠNG HÀM"
    $covTag.TextFrame.TextRange.Font.Name = "Segoe UI"
    $covTag.TextFrame.TextRange.Font.Size = 9.5
    $covTag.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $covTag.TextFrame.TextRange.Font.Color.RGB = 0x2A7300 # green
    $covTag.TextFrame.VerticalAnchor = [Microsoft.Office.Core.MsoVerticalAnchor]::msoAnchorMiddle
    
    if (Test-Path $logoPath) {
        $s1.Shapes.AddPicture($logoPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 60, 45, 190, 48) | Out-Null
    }
    
    # Cover Footer
    $covFoot = $s1.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 60, 485, 600, 25)
    $covFoot.TextFrame.TextRange.Text = "NHA KHOA THÙY ANH • HỆ THỐNG NHA KHOA CHUYÊN SÂU"
    $covFoot.TextFrame.TextRange.Font.Name = "Segoe UI"
    $covFoot.TextFrame.TextRange.Font.Size = 9
    $covFoot.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $covFoot.TextFrame.TextRange.Font.Color.RGB = 0x94A3B8
    
    # --- SLIDE 2 & 3: INFO SLIDES ---
    foreach ($sNum in @(2, 3)) {
        $s = $pres.Slides.Item($sNum)
        $s.Background.Fill.Solid()
        $s.Background.Fill.ForeColor.RGB = $cWhite
        
        # Format existing text shapes first
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
                    $sh.TextFrame.TextRange.Font.Color.RGB = 0x334155
                    $sh.TextFrame.TextRange.Font.Size = 15
                    $sh.Left = 60; $sh.Top = 75; $sh.Width = 840; $sh.Height = 400
                }
            }
        }
        
        # Logo
        if (Test-Path $logoPath) {
            $s.Shapes.AddPicture($logoPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 25, 14, 130, 30) | Out-Null
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
    
    # --- SLIDES 4 TO 31: CLINICAL PHOTO SLIDES ---
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
        
        # 2. Rounded corners & borders for ALL pictures on this slide!
        for ($i = 1; $i -le $s.Shapes.Count; $i++) {
            $sh = $s.Shapes.Item($i)
            if ($sh.Type -in 11, 13) {
                # Ensure it's a patient photo, not the logo
                if ($sh.Top -ge 50) {
                    $sh.AutoShapeType = 5 # msoShapeRoundedRectangle
                    if ($sh.Adjustments.Count -ge 1) {
                        $sh.Adjustments.Item(1) = 0.04
                    }
                    $sh.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
                    $sh.Line.ForeColor.RGB = 0xD9E2EC
                    $sh.Line.Weight = 1.2
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
        $pTag.TextFrame.TextRange.Text = "Nguyễn Vũ Bảo Ngọc - 1996"
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
    
    $pres.Save()
    
    # Export samples to view
    $pres.Slides.Item(1).Export("C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\sample_slide_1.jpg", "JPG", 1920, 1080)
    $pres.Slides.Item(2).Export("C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\sample_slide_2.jpg", "JPG", 1920, 1080)
    $pres.Slides.Item(4).Export("C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\sample_slide_4.jpg", "JPG", 1920, 1080)
    $pres.Slides.Item(5).Export("C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\sample_slide_5.jpg", "JPG", 1920, 1080)
    $pres.Slides.Item(10).Export("C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\sample_slide_10.jpg", "JPG", 1920, 1080)
    $pres.Slides.Item(22).Export("C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\sample_slide_22.jpg", "JPG", 1920, 1080)
    $pres.Slides.Item(24).Export("C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\sample_slide_24.jpg", "JPG", 1920, 1080)
    
    $pres.Close()
    Write-Host "Exported sample slides successfully!"
} finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
