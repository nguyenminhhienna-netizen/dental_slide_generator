[CmdletBinding()]
param (
    [int]$Port = 8999
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$uiDir = Join-Path $baseDir 'ui'
$slidesJsonPath = Join-Path $baseDir 'slides_data.json'
$slotsJsonPath = Join-Path $baseDir 'slots_definition.json'
$generatorScript = Join-Path $baseDir 'generate_deck.ps1'
$outputDir = Join-Path $baseDir 'output'
$uploadDir = Join-Path $baseDir 'uploads'

if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }
if (-not (Test-Path $uploadDir)) { New-Item -ItemType Directory -Path $uploadDir -Force | Out-Null }

$listener = New-Object System.Net.HttpListener
$prefix = "http://localhost:$Port/"
$listener.Prefixes.Add($prefix)

function Send-JsonResponse($response, $data, [int]$statusCode = 200) {
    $json = $data | ConvertTo-Json -Depth 5
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $response.StatusCode = $statusCode
    $response.ContentType = 'application/json; charset=utf-8'
    $response.ContentLength64 = $bytes.Length
    $response.OutputStream.Write($bytes, 0, $bytes.Length)
}

try {
    $listener.Start()
    Write-Host "=========================================================" -ForegroundColor Green
    Write-Host "  NHA KHOA THUY ANH - PHAN MEM TAO SLIDE LÂM SÀNG" -ForegroundColor Cyan
    Write-Host "  Đang chạy tại: $prefix" -ForegroundColor Yellow
    Write-Host "=========================================================" -ForegroundColor Green
    
    Start-Process $prefix

    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        $urlPath = $request.Url.LocalPath

        try {
            if ($request.HttpMethod -eq 'GET') {
                if ($urlPath -eq '/' -or $urlPath -eq '/index.html') {
                    $htmlPath = Join-Path $uiDir 'index.html'
                    $bytes = [System.IO.File]::ReadAllBytes($htmlPath)
                    $response.ContentType = 'text/html; charset=utf-8'
                    $response.ContentLength64 = $bytes.Length
                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                }
                elseif ($urlPath -eq '/api/slides') {
                    $bytes = [System.IO.File]::ReadAllBytes($slidesJsonPath)
                    $response.ContentType = 'application/json; charset=utf-8'
                    $response.ContentLength64 = $bytes.Length
                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                }
                elseif ($urlPath -eq '/api/slots') {
                    $src = if (Test-Path $slotsJsonPath) { $slotsJsonPath } else { $slidesJsonPath }
                    $bytes = [System.IO.File]::ReadAllBytes($src)
                    $response.ContentType = 'application/json; charset=utf-8'
                    $response.ContentLength64 = $bytes.Length
                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                }
                elseif ($urlPath -eq '/api/default_folder') {
                    $downloads = Join-Path $env:USERPROFILE "Downloads"
                    Send-JsonResponse $response @{
                        success = $true
                        defaultFolder = $downloads
                    }
                }
                elseif ($urlPath.StartsWith('/assets/')) {
                    $fileName = $urlPath.Substring(8)
                    $filePath = Join-Path $uiDir (Join-Path 'assets' $fileName)
                    if (Test-Path $filePath) {
                        $bytes = [System.IO.File]::ReadAllBytes($filePath)
                        $response.ContentType = 'image/png'
                        $response.ContentLength64 = $bytes.Length
                        $response.OutputStream.Write($bytes, 0, $bytes.Length)
                    } else {
                        $response.StatusCode = 404
                    }
                }
                elseif ($urlPath.StartsWith('/thumbs/')) {
                    $fileName = $urlPath.Substring(8)
                    $filePath = Join-Path $uiDir (Join-Path 'thumbs' $fileName)
                    if (Test-Path $filePath) {
                        $bytes = [System.IO.File]::ReadAllBytes($filePath)
                        $response.ContentType = 'image/jpeg'
                        $response.ContentLength64 = $bytes.Length
                        $response.OutputStream.Write($bytes, 0, $bytes.Length)
                    } else {
                        $response.StatusCode = 404
                    }
                }
                else {
                    $response.StatusCode = 404
                }
            }
            elseif ($request.HttpMethod -eq 'POST') {
                if ($urlPath -eq '/api/upload') {
                    # Handle raw binary upload for a specific slot
                    $slotId = $request.QueryString['slot']
                    $rawFileName = $request.QueryString['filename']
                    if (-not $rawFileName) { $rawFileName = "$slotId.jpg" }
                    $safeFileName = [System.IO.Path]::GetFileName($rawFileName)
                    
                    $destPath = Join-Path $uploadDir "$slotId`_$safeFileName"
                    $outStream = [System.IO.File]::Create($destPath)
                    $request.InputStream.CopyTo($outStream)
                    $outStream.Close()

                    Write-Host "Nạp ảnh thành công: Slot $slotId -> $destPath" -ForegroundColor DarkGreen
                    Send-JsonResponse $response @{
                        success = $true
                        slotId = $slotId
                        savedPath = $destPath
                    }
                }
                elseif ($urlPath -eq '/api/scan_folder') {
                    $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
                    $bodyText = $reader.ReadToEnd()
                    $reader.Close()

                    $payload = $bodyText | ConvertFrom-Json
                    $folderPath = $payload.folderPath

                    if (-not $folderPath -or -not (Test-Path $folderPath)) {
                        Send-JsonResponse $response @{
                            success = $false
                            error = "Thư mục không tồn tại: $folderPath"
                        }
                    } else {
                        $files = Get-ChildItem -Path $folderPath -File
                        $slotFiles = @{}
                        $count = 0
                        foreach ($f in $files) {
                            if ($f.BaseName -match '^(\d+)') {
                                $num = [int]$matches[1]
                                if ($num -ge 1 -and $num -le 52) {
                                    $slotFiles[$num.ToString()] = $f.FullName
                                    $count++
                                }
                            }
                        }
                        Send-JsonResponse $response @{
                            success = $true
                            folderPath = $folderPath
                            matchedCount = $count
                            slotFiles = $slotFiles
                        }
                    }
                }
                elseif ($urlPath -eq '/api/generate') {
                    $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
                    $bodyText = $reader.ReadToEnd()
                    $reader.Close()

                    $payload = $bodyText | ConvertFrom-Json
                    
                    $cleanName = ($payload.PatientName -replace '[\\/:*?"<>|]', '_').Trim()
                    $cleanYear = ($payload.BirthYear -replace '[\\/:*?"<>|]', '_').Trim()
                    $cleanDoc = ($payload.DoctorName -replace '[\\/:*?"<>|]', '_').Trim()
                    $fileName = "$cleanName - $cleanYear - $cleanDoc.pptx"

                    # Target folder: User specified folder or Downloads folder by default
                    $targetFolder = Join-Path $env:USERPROFILE "Downloads"
                    if ($payload.OutputFolder -and $payload.OutputFolder.Trim() -ne "") {
                        $trimmed = $payload.OutputFolder.Trim()
                        if (-not (Test-Path -LiteralPath $trimmed)) {
                            New-Item -ItemType Directory -Path $trimmed -Force | Out-Null
                        }
                        $targetFolder = $trimmed
                    } elseif (-not (Test-Path -LiteralPath $targetFolder)) {
                        $targetFolder = $outputDir
                    }

                    $targetPptx = Join-Path $targetFolder $fileName

                    # Inject OutputPath & OutputFolder directly into JSON so no Unicode characters get lost in CLI arguments
                    $payload | Add-Member -NotePropertyName "OutputPath" -NotePropertyValue $targetPptx -Force
                    $payload | Add-Member -NotePropertyName "OutputFolder" -NotePropertyValue $targetFolder -Force
                    $jsonToSave = $payload | ConvertTo-Json -Depth 10

                    $tempCfgPath = Join-Path $baseDir 'temp_run_config.json'
                    Set-Content -Path $tempCfgPath -Value $jsonToSave -Encoding UTF8

                    Write-Host "Bắt đầu xuất slide: $cleanName ($cleanYear)..." -ForegroundColor Cyan
                    
                    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
                    $pinfo.FileName = "powershell.exe"
                    $pinfo.Arguments = "-ExecutionPolicy Bypass -NoProfile -File `"$generatorScript`" -ConfigJsonPath `"$tempCfgPath`""
                    $pinfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8
                    $pinfo.StandardErrorEncoding  = [System.Text.Encoding]::UTF8
                    $pinfo.UseShellExecute = $false
                    $pinfo.RedirectStandardOutput = $true
                    $pinfo.RedirectStandardError = $true
                    $pinfo.CreateNoWindow = $true

                    $proc = [System.Diagnostics.Process]::Start($pinfo)
                    $stdout = $proc.StandardOutput.ReadToEnd()
                    $stderr = $proc.StandardError.ReadToEnd()
                    $proc.WaitForExit()

                    # Check for final output path in stdout
                    $finalPath = $targetPptx
                    if ($stdout -match "FINAL_OUTPUT_PATH:(.*)") {
                        $candidate = $matches[1].Trim()
                        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
                            $finalPath = $candidate
                        }
                    }

                    if ($proc.ExitCode -eq 0 -and (Test-Path -LiteralPath $finalPath)) {
                        Write-Host "Xuất slide thành công: $finalPath" -ForegroundColor Green
                        Send-JsonResponse $response @{
                            success = $true
                            outputPath = $finalPath
                            log = $stdout
                        }
                    } else {
                        Write-Host "Lỗi khi xuất slide: $stderr" -ForegroundColor Red
                        
                        # Extract friendly error message
                        $errMsg = "Không thể tạo file PowerPoint."
                        if ($stdout -match "ERROR_DETAIL:(.*)") {
                            $errMsg = $matches[1].Trim()
                        } elseif ($stderr) {
                            $errMsg = $stderr.Trim()
                        }

                        Send-JsonResponse $response @{
                            success = $false
                            error = $errMsg
                            detail = $stderr
                            log = $stdout
                        }
                    }
                }
                elseif ($urlPath -eq '/api/open') {
                    $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
                    $bodyText = $reader.ReadToEnd()
                    $reader.Close()

                    $payload = $bodyText | ConvertFrom-Json
                    if ($payload.path -and (Test-Path -LiteralPath $payload.path)) {
                        Start-Process -LiteralPath $payload.path
                        Send-JsonResponse $response @{ success = $true }
                    } else {
                        Send-JsonResponse $response @{
                            success = $false
                            error = "File không tồn tại trên ổ cứng: $($payload.path)"
                        }
                    }
                }
                elseif ($urlPath -eq '/api/open_folder') {
                    $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
                    $bodyText = $reader.ReadToEnd()
                    $reader.Close()

                    $payload = $bodyText | ConvertFrom-Json
                    if ($payload.folder -and (Test-Path -LiteralPath $payload.folder)) {
                        Start-Process -LiteralPath "explorer.exe" -ArgumentList "`"$($payload.folder)`""
                        Send-JsonResponse $response @{ success = $true }
                    } else {
                        Send-JsonResponse $response @{
                            success = $false
                            error = "Thư mục không tồn tại: $($payload.folder)"
                        }
                    }
                }
            }
        }
        catch {
            Write-Host "Error handling request: $_" -ForegroundColor Red
            Send-JsonResponse $response @{
                success = $false
                error = "Lỗi hệ thống máy chủ: $($_.Exception.Message)"
                detail = $_.ScriptStackTrace
            }
        }
        finally {
            $response.Close()
        }
    }
}
finally {
    $listener.Stop()
    $listener.Close()
}
