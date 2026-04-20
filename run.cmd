@echo off
setlocal
set "PYTHON_PATH=%~dp0.venv\Scripts\python.exe"
set "MAIN_PATH=%~dp0main.py"

if not exist "%PYTHON_PATH%" (
    echo Python not found: "%PYTHON_PATH%"
    exit /b 1
)

if not exist "%MAIN_PATH%" (
    echo Main script not found: "%MAIN_PATH%"
    exit /b 1
)

"%PYTHON_PATH%" "%MAIN_PATH%" %*
