$src = "D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\001\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi.pptx"
if (-not (Test-Path $src)) {
    $src = "D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\Ban đầu\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi.pptx"
}

$dstDir = "C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\template"
if (-not (Test-Path $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null }
$dst = Join-Path $dstDir "template.pptx"

if (Test-Path $src) {
    Copy-Item -Path $src -Destination $dst -Force
    Write-Host "Template copied to: $dst, Size: $((Get-Item $dst).Length)"
} else {
    Write-Warning "Source template not found: $src"
}
