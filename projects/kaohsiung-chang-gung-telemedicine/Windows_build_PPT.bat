@echo off
setlocal EnableExtensions

rem Always run from the directory that contains this BAT file.
cd /d "%~dp0"
if errorlevel 1 goto :project_dir_failed

set "PROJECT_DIR=%CD%"
set "VENV_DIR=%PROJECT_DIR%\.venv"
set "OUTPUT_FILE=%PROJECT_DIR%\output\kaohsiung-chang-gung-telemedicine-v1.pptx"

set "PYTHON_LAUNCHER=py -3"
py -3 --version >nul 2>nul
if errorlevel 1 (
    set "PYTHON_LAUNCHER=python"
    python --version >nul 2>nul
    if errorlevel 1 (
        echo [ERROR] Python 3 was not found.
        echo Install Python from https://www.python.org/downloads/windows/
        echo During installation, select "Add python.exe to PATH".
        pause
        exit /b 1
    )
)

echo [1/4] Creating or reusing the local virtual environment...
if not exist "%VENV_DIR%\Scripts\python.exe" (
    %PYTHON_LAUNCHER% -m venv "%VENV_DIR%"
    if errorlevel 1 goto :failed
)

echo [2/4] Installing required Python packages...
"%VENV_DIR%\Scripts\python.exe" -m pip install --disable-pip-version-check -r "%PROJECT_DIR%\requirements.txt"
if errorlevel 1 goto :failed

echo [3/4] Building the 15-slide Clinical Calm presentation...
"%VENV_DIR%\Scripts\python.exe" "%PROJECT_DIR%\build_deck.py"
if errorlevel 1 goto :failed

echo [4/4] Validating the generated presentation...
"%VENV_DIR%\Scripts\python.exe" "%PROJECT_DIR%\validate_deck.py"
if errorlevel 1 goto :failed

if not exist "%OUTPUT_FILE%" goto :missing_output

echo.
echo [SUCCESS] PowerPoint created:
echo %OUTPUT_FILE%
echo.
pause
exit /b 0

:missing_output
echo [ERROR] The build completed without creating the expected PPTX:
echo %OUTPUT_FILE%
pause
exit /b 1

:project_dir_failed
echo [ERROR] Could not switch to the project directory:
echo %~dp0
pause
exit /b 1

:failed
echo.
echo [ERROR] Build failed. Review the message above.
pause
exit /b 1
