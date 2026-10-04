$templatePptx = "D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\001\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi.pptx"
if (-not (Test-Path $templatePptx)) {
    $templatePptx = "D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\Ban đầu\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi.pptx"
}

$testPptx = "C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\test_rounded.pptx"
Copy-Item $templatePptx $testPptx -Force

$ppt = New-Object -ComObject PowerPoint.Application
try {
    $pres = $ppt.Presentations.Open($testPptx, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    
    # Test on Slide 5 (3 pictures) and Slide 10 (2 intraoral pictures)
    foreach ($sNum in @(5, 10)) {
        $slide = $pres.Slides.Item($sNum)
        Write-Host "Slide $sNum Shapes: $($slide.Shapes.Count)"
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $sh = $slide.Shapes.Item($i)
            if ($sh.Type -in 11, 13) {
                Write-Host "Picture $i : Type=$($sh.Type), Left=$($sh.Left), Top=$($sh.Top), W=$($sh.Width), H=$($sh.Height)"
                try {
                    $sh.AutoShapeType = 5 # msoShapeRoundedRectangle
                    Write-Host "  -> Successfully set AutoShapeType to 5 (Rounded Rectangle)!"
                    $sh.Line.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
                    $sh.Line.ForeColor.RGB = 0xD9E2EC # Soft slate border
                    $sh.Line.Weight = 1.5
                    # Try adjustments (corner radius)
                    if ($sh.Adjustments.Count -ge 1) {
                        $sh.Adjustments.Item(1) = 0.04
                        Write-Host "  -> Adjusted corner radius to 0.04"
                    }
                } catch {
                    Write-Warning "  -> AutoShapeType error: $_"
                }
            }
        }
    }
    
    $pres.Save()
    # Export slide 5 and 10 to check
    $slide5 = $pres.Slides.Item(5)
    $slide5.Export("C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\test_slide_5.jpg", "JPG", 1920, 1080)
    $slide10 = $pres.Slides.Item(10)
    $slide10.Export("C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\test_slide_10.jpg", "JPG", 1920, 1080)
    
    $pres.Close()
    Write-Host "Exported test slides successfully!"
} finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
