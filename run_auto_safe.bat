@echo off
chcp 65001 >nul
title IFR Automation V3 - 自动安全模式

cd /d "%~dp0"

echo.
echo ============================================================
echo    IFR Automation V3.0 - 自动处理安全项目
echo ============================================================
echo.

python ifr_automation_v3.py --root "C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)" --auto-safe-only

echo.
pause
