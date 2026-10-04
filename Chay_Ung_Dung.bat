@echo off
chcp 65001 > nul
title NHA KHOA THUY ANH - TU DONG TAO SLIDE BENH AN
cd /d %~dp0

echo ========================================================
echo   PHAN MEM TU DONG TAO SLIDE LAM SANG - NHA KHOA THUY ANH
echo ========================================================
echo Dang khoi dong he thong va mo giao dien tren trinh duyet...

powershell -ExecutionPolicy Bypass -NoProfile -File %~dp0server.ps1
pause
