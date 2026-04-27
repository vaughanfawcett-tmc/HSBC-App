@echo off
REM HSBC SLA Report launcher.
REM
REM First run: creates a venv under .venv\ and installs dependencies.
REM Every run: starts the Streamlit app and opens a browser to it.

setlocal
cd /d "%~dp0"

REM Look for a Python 3.11+ on PATH. If `py -3.11` is available (the
REM Windows Python launcher), prefer that for predictable versioning.
set PYTHON_CMD=
where py >nul 2>nul
if %errorlevel%==0 (
    py -3.11 --version >nul 2>nul
    if %errorlevel%==0 set PYTHON_CMD=py -3.11
)
if "%PYTHON_CMD%"=="" (
    where python >nul 2>nul
    if %errorlevel%==0 set PYTHON_CMD=python
)
if "%PYTHON_CMD%"=="" (
    echo.
    echo Python 3.11 or newer is required.
    echo Download it from https://www.python.org/downloads/ and tick
    echo "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Setting up the virtual environment. This happens once.
    %PYTHON_CMD% -m venv .venv
    if errorlevel 1 (
        echo Failed to create virtual environment.
        pause
        exit /b 1
    )
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Failed to install dependencies.
        pause
        exit /b 1
    )
)

echo.
echo Opening HSBC SLA Report at http://localhost:8501
echo Leave this window open while the app is running.
echo Close this window when you're done.
echo.

".venv\Scripts\python.exe" -m streamlit run app.py --browser.gatherUsageStats=false

endlocal
