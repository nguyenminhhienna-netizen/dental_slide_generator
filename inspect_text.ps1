$templatePptx = "D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\001\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi.pptx"
if (-not (Test-Path $templatePptx)) {
    $templatePptx = "D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\Ban đầu\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi.pptx"
}

$ppt = New-Object -ComObject PowerPoint.Application
try {
    $pres = $ppt.Presentations.Open($templatePptx, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    for ($s = 1; $s -le 10; $s++) {
        $slide = $pres.Slides.Item($s)
        Write-Host "--- Slide $s ---"
        foreach ($sh in $slide.Shapes) {
            if ($sh.HasTextFrame -eq [Microsoft.Office.Core.MsoTriState]::msoTrue) {
                Write-Host "  Text: $($sh.TextFrame.TextRange.Text.Trim().Replace("`r`n", " | "))"
                Write-Host "    Font=$($sh.TextFrame.TextRange.Font.Name), Size=$($sh.TextFrame.TextRange.Font.Size), Color=$($sh.TextFrame.TextRange.Font.Color.RGB)"
            }
        }
    }
    $pres.Close()
} finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
