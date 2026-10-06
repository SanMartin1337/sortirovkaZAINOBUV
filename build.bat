@echo off
REM Build Sortirovka.exe. Run on Windows. Needs Python 3.11+ (python.org).
setlocal

set PY=
where py >nul 2>nul && set PY=py
if not defined PY ( where python >nul 2>nul && set PY=python )
if not defined PY (
    echo Python not found. Install it from python.org and tick "Add python.exe to PATH".
    pause
    exit /b 1
)

%PY% -m pip install -r requirements.txt
if errorlevel 1 goto fail
%PY% -m PyInstaller --noconsole --onefile --collect-all tkinterdnd2 --name Sortirovka app.py
if errorlevel 1 goto fail

echo.
echo Done: %CD%\dist\Sortirovka.exe
pause
exit /b 0

:fail
echo.
echo BUILD FAILED - see messages above.
pause
exit /b 1
