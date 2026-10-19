@echo off
rem ============================================================
rem  vts2mp3 launcher
rem    - double click        -> open the web UI
rem    - drag a folder/file  -> convert it right away (CLI mode)
rem ============================================================
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo   Python was not found in PATH.
    echo   Please install Python 3 first, then run this file again.
    echo.
    pause
    exit /b 1
)

if "%~1"=="" (
    python "%~dp0vts2mp3.py"
) else (
    python "%~dp0vts2mp3.py" --cli %*
)

if errorlevel 1 pause
endlocal
