@echo off
setlocal

if "%~1"=="" (
    echo Usage: update_portable.cmd "C:\Path\To\Old\SEOParser_Viki"
    exit /b 1
)

powershell -ExecutionPolicy Bypass -File "%~dp0update_portable.ps1" -TargetPath "%~1"
exit /b %errorlevel%
