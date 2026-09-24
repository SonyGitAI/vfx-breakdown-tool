@echo off
rem VFX Script Breakdown - xlsx export for Windows
rem Drag your *_breakdown.json onto this file, or run:
rem     build_xlsx.bat my_breakdown.json
setlocal
if "%~1"=="" (
  echo Drag a breakdown .json file onto build_xlsx.bat to build the workbook.
  pause
  exit /b 1
)
set SCRIPT=%~dp0build_xlsx.py
where py >nul 2>nul
if %errorlevel%==0 (
  py "%SCRIPT%" %*
  goto :done
)
where python >nul 2>nul
if %errorlevel%==0 (
  python "%SCRIPT%" %*
  goto :done
)
echo Python 3 is needed for this step - it's a free install from:
echo     https://www.python.org/downloads/
echo Tick "Add python.exe to PATH" during install, then run this again.
:done
pause
