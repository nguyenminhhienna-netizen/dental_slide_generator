$templatePptx = "D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\001\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi.pptx"
if (-not (Test-Path $templatePptx)) {
    $templatePptx = "D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\Ban đầu\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi.pptx"
}

$ppt = New-Object -ComObject PowerPoint.Application
try {
    $pres = $ppt.Presentations.Open($templatePptx, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    $pres.Slides.Item(1).Export("C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\raw_slide_1.jpg", "JPG", 1920, 1080)
    $pres.Slides.Item(2).Export("C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\raw_slide_2.jpg", "JPG", 1920, 1080)
    $pres.Slides.Item(3).Export("C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\raw_slide_3.jpg", "JPG", 1920, 1080)
    $pres.Slides.Item(4).Export("C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\raw_slide_4.jpg", "JPG", 1920, 1080)
    $pres.Slides.Item(22).Export("C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\raw_slide_22.jpg", "JPG", 1920, 1080)
    $pres.Slides.Item(24).Export("C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\raw_slide_24.jpg", "JPG", 1920, 1080)
    $pres.Close()
} finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
