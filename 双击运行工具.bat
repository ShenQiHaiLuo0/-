@echo off
cd /d "%~dp0"
if not exist "%~dp0图形复制校验工具.ps1" (echo [错误] 未找到图形复制校验工具.ps1 & echo 请确保本BAT与PS1文件放在同一目录下。 & echo. & pause & exit /b 1)
title 图形版 智能复制校验工具
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0图形复制校验工具.ps1"
echo.
echo 工具已退出。
pause