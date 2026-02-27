@echo off
:: ============================================================
::  Build: Harvest Excel Export  ->  dist\HarvestExport.exe
:: ============================================================
::  Prerequisites:
::    pip install -r requirements.txt
::  Run this from the project root folder.
:: ============================================================

echo.
echo  Building HarvestExport.exe ...
echo.

python -m PyInstaller ^
  --onefile ^
  --console ^
  --name HarvestExport ^
  --distpath dist ^
  --workpath build\_pyinstaller ^
  --specpath build ^
  src\harvest_export.py

if %ERRORLEVEL% NEQ 0 (
  echo.
  echo  BUILD FAILED. Check the output above for errors.
  pause
  exit /b 1
)

echo.
echo  =====================================================
echo   Build complete!  ->  dist\HarvestExport.exe
echo  =====================================================
echo.
pause
