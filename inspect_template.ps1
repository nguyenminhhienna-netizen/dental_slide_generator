$baseDir = "C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator"
$templatePptx = Join-Path $baseDir "template\template.pptx"
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($templatePptx, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

foreach ($sNum in @(15, 16, 17, 22, 23, 24)) {
    $s = $pres.Slides.Item($sNum)
    Write-Host ("=== SLIDE " + $sNum + " ===") -ForegroundColor Yellow
    for ($i = 1; $i -le $s.Shapes.Count; $i++) {
        $sh = $s.Shapes.Item($i)
        if ($sh.HasTextFrame -eq [Microsoft.Office.Core.MsoTriState]::msoTrue) {
            $txt = $sh.TextFrame.TextRange.Text.Trim()
            if ($txt) {
                Write-Host ("  TXT [L=" + [math]::Round($sh.Left) + ", T=" + [math]::Round($sh.Top) + "]: " + $txt)
            }
        }
    }
}
$pres.Close()
$ppt.Quit()
