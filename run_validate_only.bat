@echo off
chcp 65001 >nul
title IFR Automation V3 - 仅验证模式

cd /d "%~dp0"

echo.
echo ============================================================
echo    IFR Automation V3.0 - 项目验证
echo ============================================================
echo.

python ifr_automation_v3.py --root "C:\Users\ACE\GREEN GOLD ENERGY Dropbox\Projects\Project (EPC)" --validate-only --export-report validation_report.json

echo.
pause
