[CmdletBinding()]
param()

$baseDir = "C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator"
$brainDir = "C:\Users\PC\.gemini\antigravity\brain\acfc5fe9-e6b9-4baf-ae0e-4fcb79847721"
$jsonPath = Join-Path $baseDir "mockup_data.json"
$data = Get-Content $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json

$logoPath = "D:\For work\Logo\logo-ngang.png"
$bannerPath = Join-Path $brainDir ".user_uploaded\media_1789790186457.png"
$thumbDir = Join-Path $baseDir "thumbnails"

Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 1

$ppt = New-Object -ComObject PowerPoint.Application
Start-Sleep -Seconds 1

try {
    $mockupPptx = Join-Path $baseDir "design_mockups.pptx"
    if (Test-Path $mockupPptx) { Remove-Item $mockupPptx -Force }
    
    $pres = $ppt.Presentations.Add([Microsoft.Office.Core.MsoTriState]::msoFalse)
    $pres.PageSetup.SlideWidth = 960
    $pres.PageSetup.SlideHeight = 540
    
    # BGR Colors
    $cGreen = 0x59A800       # #00A859 Medical Green
    $cNavy = 0x663F0E        # #0E3F66 Navy/Deep Ocean
    $cTeal = 0xA68200        # #0082A6 Teal/Cyan
    $cMint = 0xEEF8E6        # #E6F8EE Soft mint pill
    $cDarkGreen = 0x2A7300   # #00732A
    $cBorder = 0xF0E8E2      # #E2E8F0
    $cCardBg = 0xFAFCFF
    $cDarkText = 0x1E293B
    $cMutedText = 0x8B7464
    $cWhite = 0xFFFFFF
    
    # ====================================================
    # SLIDE 1: COVER
    # ====================================================
    $s1 = $pres.Slides.Add(1, 12)
    $s1.Background.Fill.Solid()
    $s1.Background.Fill.ForeColor.RGB = $cWhite
    
    $bannerArt = Join-Path $brainDir "banner_art_right.png"
    if (Test-Path $bannerArt) {
        $bPic = $s1.Shapes.AddPicture($bannerArt, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 490, 0, 470, 540)
        $bPic.ZOrder([Microsoft.Office.Core.MsoZOrderCmd]::msoSendToBack)
    }
    
    # Logo Capsule
    $logoCard = $s1.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle, 40, 32, 175, 46)
    $logoCard.Fill.Solid()
    $logoCard.Fill.ForeColor.RGB = $cWhite
    $logoCard.Line.ForeColor.RGB = 0xCBD5E1
    $logoCard.Line.Weight = 1.2
    if (Test-Path $logoPath) {
        $s1.Shapes.AddPicture($logoPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 48, 38, 159, 34) | Out-Null
    }
    
    # Category Tag
    $tag1 = $s1.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle, 40, 102, 280, 26)
    $tag1.Fill.Solid()
    $tag1.Fill.ForeColor.RGB = $cMint
    $tag1.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $tag1.TextFrame.TextRange.Text = $data.cover.tag
    $tag1.TextFrame.TextRange.Font.Name = "Segoe UI"
    $tag1.TextFrame.TextRange.Font.Size = 9.5
    $tag1.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $tag1.TextFrame.TextRange.Font.Color.RGB = $cDarkGreen
    $tag1.TextFrame.VerticalAnchor = [Microsoft.Office.Core.MsoVerticalAnchor]::msoAnchorMiddle
    
    # Main Deck Title
    $title1 = $s1.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 36, 136, 440, 95)
    $title1.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $title1.TextFrame.MarginLeft = 0; $title1.TextFrame.MarginTop = 0
    $tr1 = $title1.TextFrame.TextRange
    $tr1.Text = $data.cover.mainTitle + "`r`n" + $data.cover.subTitle
    $tr1.Font.Name = "Segoe UI"
    $tr1.Font.Size = 25
    $tr1.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $tr1.Font.Color.RGB = $cNavy
    $tr1.Lines(2).Font.Size = 13
    $tr1.Lines(2).Font.Color.RGB = $cTeal
    
    # Patient Info Container Card
    $card1 = $s1.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle, 40, 240, 425, 230)
    $card1.Fill.Solid()
    $card1.Fill.ForeColor.RGB = $cCardBg
    $card1.Line.ForeColor.RGB = 0xD9E2EC
    $card1.Line.Weight = 1.5
    
    # Card Header Strip
    $cardH = $s1.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle, 40, 240, 425, 36)
    $cardH.Fill.Solid()
    $cardH.Fill.ForeColor.RGB = $cGreen
    $cardH.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $cardH.TextFrame.TextRange.Text = "  " + $data.cover.cardTitle
    $cardH.TextFrame.TextRange.Font.Name = "Segoe UI"
    $cardH.TextFrame.TextRange.Font.Size = 10.5
    $cardH.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $cardH.TextFrame.TextRange.Font.Color.RGB = $cWhite
    $cardH.TextFrame.VerticalAnchor = [Microsoft.Office.Core.MsoVerticalAnchor]::msoAnchorMiddle
    
    # Card Body Text
    $cardB = $s1.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 58, 288, 395, 175)
    $cardB.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $cb = $cardB.TextFrame.TextRange
    $cbText = ($data.cover.nameLabel + ": " + $data.cover.nameVal + "`r`n`r`n" +
               $data.cover.birthLabel + ": " + $data.cover.birthVal + "`r`n" +
               $data.cover.doctorLabel + ": " + $data.cover.doctorVal + "`r`n" +
               $data.cover.branchLabel + ": " + $data.cover.branchVal + "`r`n" +
               $data.cover.dateLabel + ": " + $data.cover.dateVal)
    $cb.Text = $cbText
    $cb.Font.Name = "Segoe UI"
    $cb.Font.Size = 11.5
    $cb.Font.Color.RGB = $cDarkText
    $cb.Lines(1).Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $cb.Lines(1).Font.Size = 13
    $cb.Lines(1).Font.Color.RGB = $cNavy
    
    # Footer
    $foot1 = $s1.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 40, 488, 430, 25)
    $foot1.TextFrame.TextRange.Text = $data.cover.slogan
    $foot1.TextFrame.TextRange.Font.Name = "Segoe UI"
    $foot1.TextFrame.TextRange.Font.Size = 8.5
    $foot1.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $foot1.TextFrame.TextRange.Font.Color.RGB = 0x8A7868
    
    # ====================================================
    # SLIDE 2: KHAM NGOAI MAT (3 PHOTOS)
    # ====================================================
    $s2 = $pres.Slides.Add(2, 12)
    $s2.Background.Fill.Solid()
    $s2.Background.Fill.ForeColor.RGB = $cWhite
    
    # Top Bar
    $bar2 = $s2.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRectangle, 0, 0, 960, 56)
    $bar2.Fill.Solid()
    $bar2.Fill.ForeColor.RGB = 0xF8FAFC
    $bar2.Line.ForeColor.RGB = 0xE2E8F0
    $bar2.Line.Weight = 1
    
    # Pill Logo
    $pill2 = $s2.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle, 15, 8, 145, 40)
    $pill2.Fill.Solid()
    $pill2.Fill.ForeColor.RGB = $cWhite
    $pill2.Line.ForeColor.RGB = 0xCBD5E1
    $pill2.Line.Weight = 1
    if (Test-Path $logoPath) {
        $s2.Shapes.AddPicture($logoPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 22, 13, 131, 30) | Out-Null
    }
    
    # Category Tag
    $cTag2 = $s2.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle, 175, 14, 135, 26)
    $cTag2.Fill.Solid()
    $cTag2.Fill.ForeColor.RGB = $cMint
    $cTag2.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $cTag2.TextFrame.TextRange.Text = $data.slide5.category
    $cTag2.TextFrame.TextRange.Font.Name = "Segoe UI"
    $cTag2.TextFrame.TextRange.Font.Size = 9.5
    $cTag2.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $cTag2.TextFrame.TextRange.Font.Color.RGB = $cDarkGreen
    $cTag2.TextFrame.VerticalAnchor = [Microsoft.Office.Core.MsoVerticalAnchor]::msoAnchorMiddle
    
    # Slide Title
    $t2 = $s2.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 320, 12, 375, 32)
    $t2.TextFrame.MarginLeft = 0; $t2.TextFrame.MarginTop = 0
    $t2.TextFrame.TextRange.Text = $data.slide5.title
    $t2.TextFrame.TextRange.Font.Name = "Segoe UI"
    $t2.TextFrame.TextRange.Font.Size = 16.5
    $t2.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $t2.TextFrame.TextRange.Font.Color.RGB = $cNavy
    
    # Patient Tag
    $pTag2 = $s2.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle, 705, 14, 240, 26)
    $pTag2.Fill.Solid()
    $pTag2.Fill.ForeColor.RGB = $cWhite
    $pTag2.Line.ForeColor.RGB = 0xCBD5E1
    $pTag2.Line.Weight = 0.8
    $pTag2.TextFrame.TextRange.Text = $data.slide5.patientTag
    $pTag2.TextFrame.TextRange.Font.Name = "Segoe UI"
    $pTag2.TextFrame.TextRange.Font.Size = 9
    $pTag2.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $pTag2.TextFrame.TextRange.Font.Color.RGB = 0x64748B
    $pTag2.TextFrame.VerticalAnchor = [Microsoft.Office.Core.MsoVerticalAnchor]::msoAnchorMiddle
    
    # 3 Photo Containers
    $s5Pics = @(
        (Join-Path $thumbDir "slide_5_img_0.jpg"),
        (Join-Path $thumbDir "slide_5_img_1.jpg"),
        (Join-Path $thumbDir "slide_5_img_2.jpg")
    )
    $cardW = 302
    $cardH = 428
    $cardTop = 64
    for ($i = 0; $i -lt 3; $i++) {
        $cLeft = 15 + $i * ($cardW + 12)
        $card = $s2.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle, $cLeft, $cardTop, $cardW, $cardH)
        $card.Fill.Solid()
        $card.Fill.ForeColor.RGB = $cWhite
        $card.Line.ForeColor.RGB = 0xE2E8F0
        $card.Line.Weight = 1.2
        
        if (Test-Path $s5Pics[$i]) {
            $p = $s2.Shapes.AddPicture($s5Pics[$i], [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, $cLeft + 4, $cardTop + 4, $cardW - 8, $cardH - 44)
            $p.LockAspectRatio = [Microsoft.Office.Core.MsoTriState]::msoTrue
        }
        
        # Caption badge
        $cap = $s2.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle, $cLeft + 15, $cardTop + $cardH - 34, $cardW - 30, 26)
        $cap.Fill.Solid()
        $cap.Fill.ForeColor.RGB = 0xF1F5F9
        $cap.Line.ForeColor.RGB = 0xCBD5E1
        $cap.Line.Weight = 0.8
        $cap.TextFrame.TextRange.Text = $data.slide5.captions[$i]
        $cap.TextFrame.TextRange.Font.Name = "Segoe UI"
        $cap.TextFrame.TextRange.Font.Size = 10
        $cap.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
        $cap.TextFrame.TextRange.Font.Color.RGB = $cDarkText
        $cap.TextFrame.VerticalAnchor = [Microsoft.Office.Core.MsoVerticalAnchor]::msoAnchorMiddle
    }
    
    # Bottom Hairline Brand Accent
    $line2 = $s2.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRectangle, 0, 508, 960, 3)
    $line2.Fill.Solid()
    $line2.Fill.ForeColor.RGB = $cGreen
    $line2.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    
    $fL2 = $s2.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 15, 514, 550, 20)
    $fL2.TextFrame.MarginLeft = 0; $fL2.TextFrame.MarginTop = 0
    $fL2.TextFrame.TextRange.Text = $data.slide5.footerLeft
    $fL2.TextFrame.TextRange.Font.Name = "Segoe UI"
    $fL2.TextFrame.TextRange.Font.Size = 8.5
    $fL2.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $fL2.TextFrame.TextRange.Font.Color.RGB = 0x94A3B8
    
    $fR2 = $s2.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 840, 514, 105, 20)
    $fR2.TextFrame.MarginRight = 0; $fR2.TextFrame.MarginTop = 0
    $fR2.TextFrame.TextRange.Text = $data.slide5.page
    $fR2.TextFrame.TextRange.Font.Name = "Segoe UI"
    $fR2.TextFrame.TextRange.Font.Size = 8.5
    $fR2.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $fR2.TextFrame.TextRange.Font.Color.RGB = 0x64748B
    $fR2.TextFrame.TextRange.ParagraphFormat.Alignment = 3
    
    # ====================================================
    # SLIDE 3: KHAM TRONG MIENG (2 PHOTOS)
    # ====================================================
    $s3 = $pres.Slides.Add(3, 12)
    $s3.Background.Fill.Solid()
    $s3.Background.Fill.ForeColor.RGB = $cWhite
    
    # Top Bar
    $bar3 = $s3.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRectangle, 0, 0, 960, 56)
    $bar3.Fill.Solid()
    $bar3.Fill.ForeColor.RGB = 0xF8FAFC
    $bar3.Line.ForeColor.RGB = 0xE2E8F0
    $bar3.Line.Weight = 1
    
    # Pill Logo
    $pill3 = $s3.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle, 15, 8, 145, 40)
    $pill3.Fill.Solid()
    $pill3.Fill.ForeColor.RGB = $cWhite
    $pill3.Line.ForeColor.RGB = 0xCBD5E1
    $pill3.Line.Weight = 1
    if (Test-Path $logoPath) {
        $s3.Shapes.AddPicture($logoPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 22, 13, 131, 30) | Out-Null
    }
    
    # Category Tag
    $cTag3 = $s3.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle, 175, 14, 145, 26)
    $cTag3.Fill.Solid()
    $cTag3.Fill.ForeColor.RGB = $cMint
    $cTag3.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $cTag3.TextFrame.TextRange.Text = $data.slide10.category
    $cTag3.TextFrame.TextRange.Font.Name = "Segoe UI"
    $cTag3.TextFrame.TextRange.Font.Size = 9.5
    $cTag3.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $cTag3.TextFrame.TextRange.Font.Color.RGB = $cDarkGreen
    $cTag3.TextFrame.VerticalAnchor = [Microsoft.Office.Core.MsoVerticalAnchor]::msoAnchorMiddle
    
    # Slide Title
    $t3 = $s3.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 330, 12, 365, 32)
    $t3.TextFrame.MarginLeft = 0; $t3.TextFrame.MarginTop = 0
    $t3.TextFrame.TextRange.Text = $data.slide10.title
    $t3.TextFrame.TextRange.Font.Name = "Segoe UI"
    $t3.TextFrame.TextRange.Font.Size = 16.5
    $t3.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $t3.TextFrame.TextRange.Font.Color.RGB = $cNavy
    
    # Patient Tag
    $pTag3 = $s3.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle, 705, 14, 240, 26)
    $pTag3.Fill.Solid()
    $pTag3.Fill.ForeColor.RGB = $cWhite
    $pTag3.Line.ForeColor.RGB = 0xCBD5E1
    $pTag3.Line.Weight = 0.8
    $pTag3.TextFrame.TextRange.Text = $data.slide10.patientTag
    $pTag3.TextFrame.TextRange.Font.Name = "Segoe UI"
    $pTag3.TextFrame.TextRange.Font.Size = 9
    $pTag3.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $pTag3.TextFrame.TextRange.Font.Color.RGB = 0x64748B
    $pTag3.TextFrame.VerticalAnchor = [Microsoft.Office.Core.MsoVerticalAnchor]::msoAnchorMiddle
    
    # 2 Photo Containers
    $s10Pics = @(
        (Join-Path $thumbDir "slide_10_img_0.jpg"),
        (Join-Path $thumbDir "slide_10_img_1.jpg")
    )
    $w2 = 456
    $h2 = 428
    for ($i = 0; $i -lt 2; $i++) {
        $l2 = 18 + $i * ($w2 + 12)
        $c2 = $s3.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle, $l2, $cardTop, $w2, $h2)
        $c2.Fill.Solid()
        $c2.Fill.ForeColor.RGB = $cWhite
        $c2.Line.ForeColor.RGB = 0xE2E8F0
        $c2.Line.Weight = 1.2
        
        if (Test-Path $s10Pics[$i]) {
            $p2 = $s3.Shapes.AddPicture($s10Pics[$i], [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, $l2 + 4, $cardTop + 4, $w2 - 8, $h2 - 44)
            $p2.LockAspectRatio = [Microsoft.Office.Core.MsoTriState]::msoTrue
        }
        
        # Caption badge
        $cap2 = $s3.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle, $l2 + 40, $cardTop + $h2 - 34, $w2 - 80, 26)
        $cap2.Fill.Solid()
        $cap2.Fill.ForeColor.RGB = 0xF1F5F9
        $cap2.Line.ForeColor.RGB = 0xCBD5E1
        $cap2.Line.Weight = 0.8
        $cap2.TextFrame.TextRange.Text = $data.slide10.captions[$i]
        $cap2.TextFrame.TextRange.Font.Name = "Segoe UI"
        $cap2.TextFrame.TextRange.Font.Size = 10.5
        $cap2.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
        $cap2.TextFrame.TextRange.Font.Color.RGB = $cDarkText
        $cap2.TextFrame.VerticalAnchor = [Microsoft.Office.Core.MsoVerticalAnchor]::msoAnchorMiddle
    }
    
    # Bottom Hairline Brand Accent
    $line3 = $s3.Shapes.AddShape([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRectangle, 0, 508, 960, 3)
    $line3.Fill.Solid()
    $line3.Fill.ForeColor.RGB = $cGreen
    $line3.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
    
    $fL3 = $s3.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 15, 514, 550, 20)
    $fL3.TextFrame.MarginLeft = 0; $fL3.TextFrame.MarginTop = 0
    $fL3.TextFrame.TextRange.Text = $data.slide10.footerLeft
    $fL3.TextFrame.TextRange.Font.Name = "Segoe UI"
    $fL3.TextFrame.TextRange.Font.Size = 8.5
    $fL3.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $fL3.TextFrame.TextRange.Font.Color.RGB = 0x94A3B8
    
    $fR3 = $s3.Shapes.AddTextbox([Microsoft.Office.Core.MsoTextOrientation]::msoTextOrientationHorizontal, 840, 514, 105, 20)
    $fR3.TextFrame.MarginRight = 0; $fR3.TextFrame.MarginTop = 0
    $fR3.TextFrame.TextRange.Text = $data.slide10.page
    $fR3.TextFrame.TextRange.Font.Name = "Segoe UI"
    $fR3.TextFrame.TextRange.Font.Size = 8.5
    $fR3.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $fR3.TextFrame.TextRange.Font.Color.RGB = 0x64748B
    $fR3.TextFrame.TextRange.ParagraphFormat.Alignment = 3
    
    # Save Presentation
    $pres.SaveAs($mockupPptx)
    Write-Host "Mockup PPTX saved to: $mockupPptx"
    
    # Export High-Res Images (1920x1080)
    $out1 = Join-Path $thumbDir "mockup_slide_1_cover.jpg"
    $out2 = Join-Path $thumbDir "mockup_slide_5_content.jpg"
    $out3 = Join-Path $thumbDir "mockup_slide_10_intraoral.jpg"
    
    $s1.Export($out1, "JPG", 1920, 1080)
    $s2.Export($out2, "JPG", 1920, 1080)
    $s3.Export($out3, "JPG", 1920, 1080)
    
    Copy-Item $out1 (Join-Path $brainDir "mockup_slide_1_cover.jpg") -Force
    Copy-Item $out2 (Join-Path $brainDir "mockup_slide_5_content.jpg") -Force
    Copy-Item $out3 (Join-Path $brainDir "mockup_slide_10_intraoral.jpg") -Force
    
    Write-Host "Exported all 3 high-res mockup slides successfully!"
    $pres.Close()
}
catch {
    Write-Error $_
}
finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
