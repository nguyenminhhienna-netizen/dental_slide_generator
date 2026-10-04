$assetsDir = "C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\assets"
if (-not (Test-Path $assetsDir)) { New-Item -ItemType Directory -Path $assetsDir -Force | Out-Null }
$src = "C:\Users\PC\.gemini\antigravity\brain\acfc5fe9-e6b9-4baf-ae0e-4fcb79847721\banner_art_right.png"
$dst = Join-Path $assetsDir "banner_art_right.png"
Copy-Item $src $dst -Force
Write-Host "Copied asset successfully to $dst"
