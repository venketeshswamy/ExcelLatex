@echo off
echo Removing LaTeX Math Add-in and clearing cache...

:: 1. Remove Trusted Catalog from Registry
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Catalogs\ExcelLatexCatalog" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /f >nul 2>&1

:: 2. Delete the local manifest folder
rmdir /s /q "%LOCALAPPDATA%\ExcelLatexAddin" >nul 2>&1

:: 3. Clear Office Web Add-in Cache
rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" >nul 2>&1

echo ========================================================
echo Uninstalled successfully! 
echo Restart Excel and the add-in will be completely gone.
echo ========================================================
pause