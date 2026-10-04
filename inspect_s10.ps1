$templatePptx = 'D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\001\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi.pptx'
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($templatePptx, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
$s10 = $pres.Slides.Item(10)
Write-Host "Slide 10 Shape Count: $($s10.Shapes.Count)"
for ($i = 1; $i -le $s10.Shapes.Count; $i++) {
    $sh = $s10.Shapes.Item($i)
    Write-Host "Shape $i : Name=$($sh.Name), Type=$($sh.Type), Left=$($sh.Left), Top=$($sh.Top), Width=$($sh.Width), Height=$($sh.Height)"
}
$pres.Close()
$ppt.Quit()
