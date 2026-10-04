$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "=== TEST 1: GET /api/slides ==="
$r1 = Invoke-RestMethod -Uri "http://localhost:8999/api/slides" -Method Get
Write-Host "Slide count returned:" $r1.Count

Write-Host "`n=== TEST 2: GET /api/default_folder ==="
$rFolder = Invoke-RestMethod -Uri "http://localhost:8999/api/default_folder" -Method Get
Write-Host "Default folder returned:" $rFolder.defaultFolder
if (-not $rFolder.defaultFolder -or $rFolder.defaultFolder -notlike "*Downloads*") {
    Write-Error "Default folder should contain Downloads!"
}

Write-Host "`n=== TEST 3: POST /api/upload ==="
$dummyBytes = [System.IO.File]::ReadAllBytes("C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\ui\assets\logo.png")
$uploadUrl = "http://localhost:8999/api/upload?slot=1&filename=slot1_test.png"
$r2 = Invoke-RestMethod -Uri $uploadUrl -Method Post -Body $dummyBytes -ContentType "application/octet-stream"
Write-Host "Upload result:" ($r2 | ConvertTo-Json)

Write-Host "`n=== TEST 4: POST /api/generate (With OutputFolder in Downloads) ==="
$downloadsDir = Join-Path $env:USERPROFILE "Downloads"
$genBody = @{
    PatientName = "Nguyễn Vũ Bảo Ngọc"
    BirthYear = "1996"
    DoctorName = "Lê Võ Yến Nhi"
    ClinicBranch = "Tp. Hồ Chí Minh"
    ChiefComplaint = "Đau khớp thái dương hàm bên trái khi ăn nhai."
    DentalHistory = "Đã từng điều trị chỉnh nha."
    OutputFolder = $downloadsDir
    SlotFiles = @{
        "1" = $r2.savedPath
    }
} | ConvertTo-Json

$r4 = Invoke-RestMethod -Uri "http://localhost:8999/api/generate" -Method Post -Body $genBody -ContentType "application/json; charset=utf-8"
Write-Host "Generate result:" ($r4 | ConvertTo-Json)

if ($r4.success -and (Test-Path -LiteralPath $r4.outputPath)) {
    Write-Host "`n>>> ALL API TESTS PASSED! Output file created at: $($r4.outputPath) <<<" -ForegroundColor Green
} else {
    Write-Error "Generate failed: $($r4.error)"
}
