@echo off
chcp 65001 >nul
title IFR Automation V3 - 智能项目同步

cd /d "%~dp0"

echo.
echo ============================================================
echo    IFR Automation V3.0 - 智能项目同步工具
echo ============================================================
echo.

python ifr_automation_v3.py

echo.
pause
