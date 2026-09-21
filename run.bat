@echo off
REM ================================================
REM Business English Thesis Assistant (Windows)
REM ================================================
chcp 65001 >nul 2>&1

set DIR=%~dp0
cd /d "%DIR%"

echo.
echo ================================================
echo   Business English Thesis Assistant
echo   Working dir: %DIR%
echo ================================================
echo.

REM --- Step 1: find Python ---
set PY=
where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set PY=python
) else (
    where py >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        set PY=py -3
    ) else (
        echo [ERROR] Python not found!
        echo.
        echo Please install Python 3.10 or newer:
        echo   1. Go to https://www.python.org/downloads/windows/
        echo   2. Download and install the latest Python
        echo   3. IMPORTANT: check "Add Python to PATH" during install
        echo   4. Re-run this script after install
        echo.
        cmd /k
        exit /b 1
    )
)
echo [1/4] Python OK: %PY%

REM --- Step 2: check tkinter ---
%PY% -c "import tkinter; import sys; sys.exit(0)" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] tkinter not available!
    echo.
    echo Your Python install is missing tkinter (the GUI library).
    echo Re-run the Python installer and tick "tcl/tk and IDLE".
    echo.
    cmd /k
    exit /b 1
)
echo [2/4] tkinter OK

REM --- Step 3: venv + deps ---
set VENV=%DIR%venv
if not exist "%VENV%\Scripts\python.exe" (
    echo [3/4] First run: creating virtual environment...
    %PY% -m venv "%VENV%"
    if %ERRORLEVEL% NEQ 0 (
        echo [ERROR] Failed to create venv!
        echo.
        cmd /k
        exit /b 1
    )
    "%VENV%\Scripts\pip.exe" install --upgrade pip >nul 2>&1
    "%VENV%\Scripts\pip.exe" install -r "%DIR%requirements.txt"
    if %ERRORLEVEL% NEQ 0 (
        echo [ERROR] Failed to install dependencies!
        echo Try manually:
        echo   "%VENV%\Scripts\pip.exe" install -r "%DIR%requirements.txt"
        echo.
        cmd /k
        exit /b 1
    )
) else (
    echo [3/4] venv OK
)

REM --- Step 4: launch app ---
echo [4/4] Launching app...
echo.
"%VENV%\Scripts\python.exe" "%DIR%app.py"
set APP_EXIT=%ERRORLEVEL%

if %APP_EXIT% NEQ 0 (
    echo.
    echo ================================================
    echo [ERROR] App exited with code %APP_EXIT%
    echo ================================================
    echo.
    echo Common causes:
    echo   - A Word document is locked (close Word and retry)
    echo   - data\topics.json is missing
    echo   - Encoding issue on Windows
    echo.
    echo Full error shown above. Press any key to close.
    cmd /k
)

endlocal