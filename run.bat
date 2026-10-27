@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"
title SSQ Downloader - Diagnostic Launcher
if not exist "双色球数据" mkdir "双色球数据"
set "LOG=%~dp0双色球数据\启动诊断日志.txt"
echo ========= SSQ Launcher =========>"%LOG%"
echo Location: %CD%>>"%LOG%"
echo Date: %DATE% %TIME%>>"%LOG%"
echo.
echo =============================================
echo SSQ Downloader v3 - Windows Launcher
echo Folder: %CD%
echo =============================================
echo.
if not exist "ssq_downloader.py" goto missing_script
if not exist "requirements.txt" goto missing_reqs
where py >nul 2>nul
if not errorlevel 1 (
  set "PY=py"
  set "PY_ARG=-3"
  goto found_python
)
where python >nul 2>nul
if not errorlevel 1 (
  set "PY=python"
  set "PY_ARG="
  goto found_python
)
echo ERROR: Cannot find Python 3.
echo ERROR: Cannot find Python 3.>>"%LOG%"
goto failure
:found_python
echo [1/3] Checking Python...
%PY% %PY_ARG% --version
if errorlevel 1 goto python_error
%PY% %PY_ARG% --version >>"%LOG%" 2>&1
echo [2/3] Checking dependencies...
%PY% %PY_ARG% -c "import requests,bs4,openpyxl,selenium" >nul 2>nul
if not errorlevel 1 goto dependencies_ok
echo Missing packages. Installing from requirements.txt ...
%PY% %PY_ARG% -m pip install -r "requirements.txt" >>"%LOG%" 2>&1
if errorlevel 1 goto install_error
:dependencies_ok
echo [3/3] Downloading data... details in SSQ data folder.
%PY% %PY_ARG% -u "ssq_downloader.py"
set "RESULT=%ERRORLEVEL%"
echo Python exit code: %RESULT%>>"%LOG%"
if not "%RESULT%"=="0" goto failure
echo.
echo SUCCESS. Check the SSQ data folder next to run.bat.
goto finish
:missing_script
echo ERROR: ssq_downloader.py is missing. Extract the complete ZIP.
echo ERROR: Missing ssq_downloader.py>>"%LOG%"
goto failure
:missing_reqs
echo ERROR: requirements.txt is missing. Extract the complete ZIP.
echo ERROR: Missing requirements.txt>>"%LOG%"
goto failure
:python_error
echo ERROR: Python could not start.
echo ERROR: Python could not start.>>"%LOG%"
goto failure
:install_error
echo ERROR: Pip install failed. Read startup log.
echo ERROR: Pip install failed.>>"%LOG%"
goto failure
:failure
echo.
echo FAILED. Please open the SSQ data folder and send logs.
echo Startup log: "%LOG%"
:finish
echo.
echo This window will stay open until you press a key.
pause
endlocal
