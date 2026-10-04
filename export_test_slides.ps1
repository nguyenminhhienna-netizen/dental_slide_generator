$baseDir = "C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator"
$pptxPath = Join-Path $baseDir "output\test_output_redesign.pptx"
$thumbDir = Join-Path $baseDir "thumbnails"
$brainDir = "C:\Users\PC\.gemini\antigravity\brain\acfc5fe9-e6b9-4baf-ae0e-4fcb79847721"

Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 1

$ppt = New-Object -ComObject PowerPoint.Application
Start-Sleep -Seconds 1

try {
    $pres = $ppt.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    
    $slidesToExport = @(1, 2, 5, 10, 22, 24, 29)
    foreach ($sNum in $slidesToExport) {
        $slide = $pres.Slides.Item($sNum)
        $outJpg = Join-Path $thumbDir "gen_slide_$sNum.jpg"
        $slide.Export($outJpg, "JPG", 1920, 1080)
        Copy-Item $outJpg (Join-Path $brainDir "gen_slide_$sNum.jpg") -Force
        Write-Host "Exported slide $sNum -> $outJpg"
    }
    
    $pres.Close()
} finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
