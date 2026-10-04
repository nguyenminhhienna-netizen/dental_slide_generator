$pptx = "C:\Users\PC\Downloads\Nguyễn Vũ Bảo Ngọc - 1996 - Lê Võ Yến Nhi.pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($pptx, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

$s4 = $pres.Slides.Item(4)
Write-Host "Slide 4 Shapes Count:" $s4.Shapes.Count
for ($i = 1; $i -le $s4.Shapes.Count; $i++) {
    $sh = $s4.Shapes.Item($i)
    Write-Host ("Shape " + $i + " Type=" + $sh.Type + " Name=" + $sh.Name + " L=" + [math]::Round($sh.Left) + " T=" + [math]::Round($sh.Top) + " W=" + [math]::Round($sh.Width) + " H=" + [math]::Round($sh.Height))
}

$pres.Slides.Item(4).Export("C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\s4_generated.jpg", "JPG", 1920, 1080)
$pres.Close()
$ppt.Quit()
