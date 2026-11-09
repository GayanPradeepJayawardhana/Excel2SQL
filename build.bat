@echo off
REM ================================================
REM  Excel2MySQL - one-click build
REM ================================================
cd /d "%~dp0"

echo Installing dependencies...
python -m pip install --upgrade pip
if errorlevel 1 goto :fail
python -m pip install -r requirements.txt
if errorlevel 1 goto :fail

echo.
echo Building .exe with PyInstaller...
python -m PyInstaller --onefile --noconsole --clean ^
    --name Excel2MySQL ^
    --collect-all tkinterdnd2 ^
    app.py
if errorlevel 1 goto :fail

echo.
echo ================================================
echo  Done!  Your exe is at:  dist\Excel2MySQL.exe
echo ================================================
pause
exit /b 0

:fail
echo.
echo BUILD FAILED - see the messages above.
pause
exit /b 1