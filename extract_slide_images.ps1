Add-Type -AssemblyName System.IO.Compression.FileSystem
$baseDir = "C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator"
$jsonPath = Join-Path $baseDir "mockup_data.json"
$data = Get-Content $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
$thumbDir = Join-Path $baseDir "thumbnails"

$zip = [System.IO.Compression.ZipFile]::OpenRead($data.templatePptx)

foreach ($slideNum in @(5, 10)) {
    $relName = "ppt/slides/_rels/slide$slideNum.xml.rels"
    $entry = $zip.GetEntry($relName)
    if ($entry) {
        $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
        $content = $reader.ReadToEnd()
        $reader.Dispose()
        Write-Host "--- Slide $slideNum Relationships ---"
        # Parse XML
        [xml]$x = $content
        $imgIdx = 0
        foreach ($rel in $x.Relationships.Relationship) {
            if ($rel.Type -like "*image*") {
                $target = $rel.Target # e.g. ../media/image12.jpeg
                $mediaFile = "ppt/media/" + [System.IO.Path]::GetFileName($target)
                $mediaEntry = $zip.GetEntry($mediaFile)
                if ($mediaEntry) {
                    $outPath = Join-Path $thumbDir "slide_${slideNum}_img_${imgIdx}.jpg"
                    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($mediaEntry, $outPath, $true)
                    Write-Host ("Slide " + $slideNum + " Image " + $imgIdx + " -> " + $outPath)
                    $imgIdx++
                }
            }
        }
    }
}

$zip.Dispose()
