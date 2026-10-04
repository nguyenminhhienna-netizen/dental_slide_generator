param([string]$FilePath)
$bytes = [System.IO.File]::ReadAllBytes($FilePath)
if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    Write-Host "Already has BOM: $FilePath"
    exit 0
}
$bom = [byte[]]@(0xEF, 0xBB, 0xBF)
$newBytes = New-Object byte[] ($bytes.Length + 3)
[System.Buffer]::BlockCopy($bom, 0, $newBytes, 0, 3)
[System.Buffer]::BlockCopy($bytes, 0, $newBytes, 3, $bytes.Length)
[System.IO.File]::WriteAllBytes($FilePath, $newBytes)
Write-Host "Added UTF-8 BOM to: $FilePath"
