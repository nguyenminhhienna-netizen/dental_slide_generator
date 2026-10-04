$baseDir = "C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator"
$pptxPath = Join-Path $baseDir "output\test_run.pptx"
$uiThumbsDir = Join-Path $baseDir "ui\thumbs"
$artifactDir = "C:\Users\PC\.gemini\antigravity\brain\acfc5fe9-e6b9-4baf-ae0e-4fcb79847721"

Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 1

$ppt = New-Object -ComObject PowerPoint.Application
Start-Sleep -Seconds 1

try {
    $pres = $ppt.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
        $slide = $pres.Slides.Item($i)
        $outJpg = Join-Path $uiThumbsDir "slide_$i.jpg"
        $slide.Export($outJpg, "JPG", 1280, 720)
        $artJpg = Join-Path $artifactDir "slide_$i.jpg"
        $slide.Export($artJpg, "JPG", 1280, 720)
        Write-Host "Exported UI thumbnail for slide $i -> $outJpg"
    }
    
    $pres.Close()
    Write-Host "All 31 UI thumbnails updated successfully!"
} finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
