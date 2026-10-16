@echo off
rem ==========================================================================
rem  build.bat - creates InventoryTracker.accdb in this folder.
rem  Double-click it. Needs Microsoft Access and Excel (2010 or newer).
rem  Optional: build.bat "D:\Data\MyInventory.accdb"
rem ==========================================================================
setlocal
cd /d "%~dp0"

if not exist "%~dp0build.vbs" (
    echo build.vbs is missing next to this file.
    pause
    exit /b 1
)

if "%~1"=="" (
    set "TARGET=%~dp0InventoryTracker.accdb"
) else (
    set "TARGET=%~f1"
)

if exist "%TARGET%" (
    echo The file already exists:
    echo   %TARGET%
    echo Delete or rename it, then run build.bat again.
    pause
    exit /b 1
)

echo Building "%TARGET%" ...
echo Access will open and show a message when the build is finished.
cscript //nologo "%~dp0build.vbs" "%TARGET%"
if errorlevel 1 (
    echo.
    echo Build failed. See the message above.
    pause
    exit /b 1
)

echo Done.
endlocal
