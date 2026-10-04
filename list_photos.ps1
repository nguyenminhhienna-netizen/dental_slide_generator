$dir = "D:\Tổng Quát\Khớp thái dương hàm\2026\T9.2026\Nguyễn Vũ Bảo Ngọc - 1996 - Bs Nhi\001"
Get-ChildItem -Path $dir -Filter "*.JPG" | Select-Object -Property Name, Length | Format-Table -AutoSize
