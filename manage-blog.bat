@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo   Hugo Blog Post Manager
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0manage-posts.ps1"
