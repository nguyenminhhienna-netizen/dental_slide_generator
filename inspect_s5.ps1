$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open('C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\test_aesthetic.pptx', [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
$s5 = $pres.Slides.Item(5)
Write-Host "Total shapes on Slide 5: $($s5.Shapes.Count)"
for ($i = 1; $i -le $s5.Shapes.Count; $i++) {
    $sh = $s5.Shapes.Item($i)
    $txt = ""
    if ($sh.HasTextFrame -eq [Microsoft.Office.Core.MsoTriState]::msoTrue) { $txt = $sh.TextFrame.TextRange.Text.Trim() }
    Write-Host "Shape $i : Name=$($sh.Name), Type=$($sh.Type), Left=$([math]::Round($sh.Left)), Top=$([math]::Round($sh.Top)), W=$([math]::Round($sh.Width)), H=$([math]::Round($sh.Height)), Text='$txt'"
}
$pres.Close()
$ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
