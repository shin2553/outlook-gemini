@echo off
cd /d "%~dp0"
chcp 65001 > nul

echo ======================================
echo  OutlookGemini .exe Build Start
echo ======================================

:: Python Scripts 경로 자동 탐색
for /f "delims=" %%i in ('python -c "import sys, os; print(os.path.join(sys.prefix, 'Scripts'))"') do set SCRIPTS=%%i
set PYINSTALLER=%SCRIPTS%\pyinstaller.exe

if not exist "%PYINSTALLER%" (
    echo Installing PyInstaller...
    pip install pyinstaller
)

if exist dist\OutlookGemini rmdir /s /q dist\OutlookGemini
if exist build\OutlookGemini rmdir /s /q build\OutlookGemini

"%PYINSTALLER%" outlook_gemini.spec

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

echo.
echo Copying config, profiles and manual...
copy /Y config_template.ini "dist\OutlookGemini\config.ini"
copy /Y MANUAL.md "dist\OutlookGemini\MANUAL.md"
echo [] > "dist\OutlookGemini\profiles.json"

echo.
echo ======================================
echo  Build complete!
echo  Output: dist\OutlookGemini\
echo ======================================
pause
