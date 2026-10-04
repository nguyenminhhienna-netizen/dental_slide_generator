$templatePptx = "D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\001\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi.pptx"
if (-not (Test-Path $templatePptx)) {
    $templatePptx = "D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\Ban đầu\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi.pptx"
}

$baseDir = "C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator"
Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 1

$ppt = New-Object -ComObject PowerPoint.Application
Start-Sleep -Seconds 1

try {
    $pres = $ppt.Presentations.Open($templatePptx, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    
    foreach ($sNum in @(4, 5, 8, 10, 15, 22)) {
        $s = $pres.Slides.Item($sNum)
        Write-Host "=== SLIDE $sNum (Shapes: $($s.Shapes.Count)) ==="
        for ($i = 1; $i -le $s.Shapes.Count; $i++) {
            $sh = $s.Shapes.Item($i)
            Write-Host "  Shape $i : Name=$($sh.Name), Type=$($sh.Type), Left=$([math]::Round($sh.Left,1)), Top=$([math]::Round($sh.Top,1)), W=$([math]::Round($sh.Width,1)), H=$([math]::Round($sh.Height,1)), Ratio=$([math]::Round($sh.Width / [math]::Max(1,$sh.Height), 2))"
        }
    }
    
    $pres.Close()
} finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
