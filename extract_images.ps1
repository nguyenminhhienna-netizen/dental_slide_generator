Add-Type -AssemblyName System.IO.Compression.FileSystem
$baseDir = "C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator"
$jsonPath = Join-Path $baseDir "mockup_data.json"
$data = Get-Content $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json

$thumbDir = Join-Path $baseDir "thumbnails"
if (-not (Test-Path $thumbDir)) { New-Item -ItemType Directory -Path $thumbDir -Force | Out-Null }

$zip = [System.IO.Compression.ZipFile]::OpenRead($data.templatePptx)
Write-Host "Total entries in PPTX:" $zip.Entries.Count
$media = @($zip.Entries | Where-Object { $_.FullName -like "ppt/media/*" })
Write-Host "Total media images in PPTX:" $media.Count

# Extract first 15 images to thumbDir
for ($i = 0; $i -lt [Math]::Min(15, $media.Count); $i++) {
    $entry = $media[$i]
    $outName = "extracted_$($i)_" + [System.IO.Path]::GetFileName($entry.FullName)
    $dest = Join-Path $thumbDir $outName
    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $dest, $true)
    Write-Host "Extracted $outName ($($entry.Length) bytes)"
}
$zip.Dispose()
