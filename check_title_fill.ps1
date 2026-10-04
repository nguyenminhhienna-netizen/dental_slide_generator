$templatePptx = "D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\001\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi.pptx"
if (-not (Test-Path $templatePptx)) {
    $templatePptx = "D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\Ban đầu\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi.pptx"
}
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($templatePptx, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
$s = $pres.Slides.Item(5)
$sh = $s.Shapes.Item(1)
Write-Host "Shape 1 Name: $($sh.Name)"
Write-Host "Shape 1 Fill Type: $($sh.Fill.Type)"
Write-Host "Shape 1 Fill Visible: $($sh.Fill.Visible)"
Write-Host "Shape 1 Fill ForeColor RGB: $($sh.Fill.ForeColor.RGB.ToString('X'))"
Write-Host "Shape 1 Text: $($sh.TextFrame.TextRange.Text)"
Write-Host "Shape 1 Font Color: $($sh.TextFrame.TextRange.Font.Color.RGB.ToString('X'))"
$pres.Close()
$ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
