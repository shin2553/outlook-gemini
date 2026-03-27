@echo off
cd /d "%~dp0"
chcp 65001 > nul

echo ======================================
echo  OutlookGemini .exe Build Start
echo ======================================

:: 실행 중인 EXE 종료 후 빌드
taskkill /f /im OutlookGemini.exe > /dev/null 2>&1
timeout /t 1 > nul

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
if not exist "dist\OutlookGemini\profiles.json" (
    echo [] > "dist\OutlookGemini\profiles.json"
)

echo.
echo Copying manual_index.py...
copy /Y "manual_index.py" "dist\OutlookGemini\manual_index.py"

echo.
echo Copying Manuals_txt...
xcopy /E /I /Y "Manuals_txt" "dist\OutlookGemini\Manuals_txt"

echo.
echo Copying Manuals (PDF originals)...
xcopy /E /I /Y "Manuals" "dist\OutlookGemini\Manuals"

echo.
echo ======================================
echo  Build complete!
echo  Output: dist\OutlookGemini\r
echo ======================================
pause
