try {
    Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 1
    
    $ppt = New-Object -ComObject PowerPoint.Application
    Start-Sleep -Seconds 1
    
    # Try adding presentation with WithWindow = msoFalse
    $pres = $ppt.Presentations.Add([Microsoft.Office.Core.MsoTriState]::msoFalse)
    Write-Host "Success! Presentation created headless."
    Write-Host "Slides count: $($pres.Slides.Count)"
    $pres.Close()
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    Write-Host "Cleaned up."
} catch {
    Write-Error $_
}
