$baseDir = "C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator"
$templatePptx = Join-Path $baseDir "template\template.pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($templatePptx, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

$s21 = $pres.Slides.Item(21)
Write-Host "Slide 21 shapes count: " $s21.Shapes.Count
for ($i = 1; $i -le $s21.Shapes.Count; $i++) {
    $sh = $s21.Shapes.Item($i)
    Write-Host ("Shape " + $i + " Type=" + $sh.Type + " Name=" + $sh.Name + " L=" + [math]::Round($sh.Left) + " T=" + [math]::Round($sh.Top) + " W=" + [math]::Round($sh.Width) + " H=" + [math]::Round($sh.Height))
}

$pres.Close()
$ppt.Quit()
