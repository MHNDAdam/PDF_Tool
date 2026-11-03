@echo off
setlocal EnableExtensions
rem ======================================================================
rem  PDF Tool launcher - version 1.0.0 (2026-10-27)
rem  Finds Python (offers to install it if missing) and starts pdf_tool.py
rem  Keep this file in the same folder as pdf_tool.py.
rem ======================================================================
set "LAUNCHER_VERSION=1.0.0"
title PDF Tool Launcher v%LAUNCHER_VERSION%
echo PDF Tool Launcher v%LAUNCHER_VERSION%

set "SCRIPT=%~dp0pdf_tool.py"
if exist "%SCRIPT%" goto have_script
echo.
echo ERROR: pdf_tool.py was not found next to this launcher.
echo Expected: %SCRIPT%
pause
exit /b 1

:have_script
call :find_python
if defined PYEXE goto run

echo.
echo Python was not found on this PC.
choice /c YN /n /m "Install Python now? (Y/N): "
if errorlevel 2 goto no_python
call :install_python
call :find_python
if defined PYEXE goto run
echo.
echo Python may have been installed, but this window cannot see it yet.
echo Close this window and run run_pdf_tool.bat again.
pause
exit /b 1

:no_python
echo.
echo Python is required. Download it from https://www.python.org/downloads/
echo (tick "Add python.exe to PATH" during setup), then run this file again.
pause
exit /b 1

:run
"%PYEXE%" %PYARGS% "%SCRIPT%" %*
if errorlevel 1 pause
endlocal
exit /b 0


rem ---------------------------------------------------------------------
:find_python
set "PYEXE="
set "PYARGS="
py -3 --version >nul 2>&1
if not errorlevel 1 (
    set "PYEXE=py"
    set "PYARGS=-3"
    goto :eof
)
python --version >nul 2>&1
if not errorlevel 1 (
    set "PYEXE=python"
    goto :eof
)
rem Not on PATH (for example just installed): look in the usual folders.
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do if exist "%%~fD\python.exe" set "PYEXE=%%~fD\python.exe"
for /d %%D in ("%ProgramFiles%\Python3*") do if exist "%%~fD\python.exe" set "PYEXE=%%~fD\python.exe"
goto :eof


rem ---------------------------------------------------------------------
:install_python
where winget >nul 2>&1
if errorlevel 1 goto install_python_download
echo.
echo Installing Python with winget...
winget install --id Python.Python.3.12 -e --accept-package-agreements --accept-source-agreements
call :find_python
if defined PYEXE goto :eof
echo winget did not finish the install. Trying a direct download...

:install_python_download
set "PYVER=3.12.8"
set "ARCHSUF=-amd64"
if /i "%PROCESSOR_ARCHITECTURE%"=="ARM64" set "ARCHSUF=-arm64"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if not defined PROCESSOR_ARCHITEW6432 set "ARCHSUF="
set "PYURL=https://www.python.org/ftp/python/%PYVER%/python-%PYVER%%ARCHSUF%.exe"
set "PYINST=%TEMP%\python-%PYVER%-setup.exe"
echo.
echo Downloading Python %PYVER% ...
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PYURL%' -OutFile '%PYINST%' -UseBasicParsing"
if not exist "%PYINST%" goto install_python_failed
echo Installing Python for the current user (please wait)...
start "" /wait "%PYINST%" /passive InstallAllUsers=0 PrependPath=1 Include_launcher=1 Include_test=0
del "%PYINST%" >nul 2>&1
goto :eof

:install_python_failed
echo.
echo Could not download Python. Check the internet connection, or install it
echo manually from https://www.python.org/downloads/
goto :eof
