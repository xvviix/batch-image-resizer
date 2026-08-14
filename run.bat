@echo off
chcp 65001 >nul
title Batch Image Resizer - Launcher
cd /d "%~dp0"

echo ============================================================
echo    Batch Image Resizer v11  -  One-Click Launcher
echo ============================================================
echo.
echo This script will:
echo -----------------------------------------------
echo  [1] Check that Python is installed
echo  [2] Install all required libraries
echo      from requirements.txt - first run only
echo  [3] Launch the program automatically
echo.
echo Note: First run takes a few minutes.
echo       Next runs take only seconds.
echo.
pause

echo.
echo ============================================================
echo  [1/3] Checking Python ...
echo ============================================================
python --version >nul 2>&1
if errorlevel 1 goto no_python
for /f "delims=" %%v in ('python --version 2^>^&1') do set "PYVER=%%v"
echo  [OK] Python found: %PYVER%

REM Check Python version is suitable - 3.8 to 3.12
python -c "import sys; sys.exit(0 if (3,8) <= sys.version_info[:2] <= (3,12) else 1)" >nul 2>&1
if errorlevel 1 goto py_version_warn
goto install_deps

:no_python
echo.
echo  [ERROR] Python was not found on this system!
echo.
echo  Please do the following:
echo   1 - Go to python.org
echo   2 - Download Python 3.10 to 3.12
echo   3 - During install, CHECK the option
echo      'Add Python to PATH'
echo   4 - After install, run this file again
echo.
pause
exit /b 1

:py_version_warn
echo  [WARNING] Your Python version is 3.13 or newer.
echo  Some features may not work correctly.
echo  It is recommended to install Python 3.12.
echo.

:install_deps
echo.
echo ============================================================
echo  [2/3] Installing required libraries ...
echo ============================================================
echo  This step only takes long on the first run.
echo  Already installed packages are skipped.
echo.
%PYCMD% -m pip install -r requirements.txt
if errorlevel 1 goto install_fail
echo.
echo  [OK] All libraries are ready.
echo.
goto choose_version

:install_fail
echo.
echo  [WARNING] Some libraries failed to install.
echo  The program will still open.
echo  If voice features do not work, check your
echo  internet connection and run again.
echo.

:choose_version
echo.
echo ============================================================
echo  [3/3] Which version do you want to run?
echo ============================================================
echo.
echo   [1] English version
echo   [2] Persian version
echo.
set /p CHOICE="Enter 1 or 2: "
if "%CHOICE%"=="1" goto run_en
if "%CHOICE%"=="2" goto run_fa
echo  Invalid choice, defaulting to English.

:run_en
%PYCMD% batch_image_resizer_v11_en.py
if errorlevel 1 goto run_fail
goto end

:run_fa
%PYCMD% batch_image_resizer_v11_fa.py
if errorlevel 1 goto run_fail
goto end

:run_fail
echo.
echo  [ERROR] The program closed with an error.
echo  If it is related to voice features, make sure
echo  all libraries installed correctly.
echo.

:end
echo.
echo  Program closed. Run this file again to restart.
echo.
pause
