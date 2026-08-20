@echo off
setlocal
echo Installing LaTeX Math Excel Add-in...

set TARGET_DIR=%LOCALAPPDATA%\ExcelLatexAddin
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

:: Copy or download manifest
copy /Y "%~dp0public\manifest.xml" "%TARGET_DIR%\manifest.xml" >nul 2>&1
if not exist "%TARGET_DIR%\manifest.xml" (
    powershell -Command "Invoke-WebRequest -Uri 'https://venketeshswamy.github.io/ExcelLatex/manifest.xml' -OutFile '%TARGET_DIR%\manifest.xml'"
)

:: Register as Trusted Catalog in Office WEF Registry
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Catalogs\ExcelLatexCatalog" /v "Url" /t REG_SZ /d "%TARGET_DIR%" /f >nul
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Catalogs\ExcelLatexCatalog" /v "Flags" /t REG_DWORD /d 1 /f >nul

echo ========================================================
echo Installation Successful!
echo 1. Open Excel -> Insert -> My Add-ins -> SHARED FOLDER
echo 2. Click 'LaTeX Math' and click Add.
echo All .xltx templates and formulas will now work without debug warnings.
echo ========================================================
pause