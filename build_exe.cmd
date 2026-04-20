@echo off
setlocal

set "BUILD_SCRIPT=%~dp0build_exe.ps1"

if not exist "%BUILD_SCRIPT%" (
    echo Build script not found: "%BUILD_SCRIPT%"
    exit /b 1
)

powershell -ExecutionPolicy Bypass -File "%BUILD_SCRIPT%" %*
exit /b %errorlevel%
