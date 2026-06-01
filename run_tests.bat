@echo off
chcp 65001 >nul
title IFR Automation V3 - 运行测试

cd /d "%~dp0"

echo.
echo ============================================================
echo    IFR Automation V3 测试套件
echo ============================================================
echo.

python test_v3.py

echo.
pause
