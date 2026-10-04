$data = Get-Content 'C:\Users\PC\.gemini\antigravity\scratch\dental_slide_generator\slides_data.json' -Raw -Encoding UTF8 | ConvertFrom-Json
$data | Select-Object slideNum, category, title, @{Name='SlotCount';Expression={$_.slots.Count}} | Format-Table -AutoSize
