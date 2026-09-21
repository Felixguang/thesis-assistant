@echo off
cd /d "%~dp0"
echo Current dir: %CD%
echo.
echo Step 1: Checking Python...
python --version
if errorlevel 1 (
    echo.
    echo [ERROR] python not found in PATH
    echo Please install Python 3.10+ and check "Add Python to PATH"
    echo.
    pause
    exit /b 1
)
echo.
echo Step 2: Checking venv...
if not exist "venv\Scripts\python.exe" (
    echo First run: creating venv...
    python -m venv venv
    call venv\Scripts\activate.bat
    pip install --upgrade pip
    pip install -r requirements.txt
) else (
    call venv\Scripts\activate.bat
)
echo.
echo Step 3: Launching app...
python app.py
echo.
echo App exited. Press any key to close.
pause >nul