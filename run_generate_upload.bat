@echo off
setlocal
cd /d "%~dp0"
python generate_and_upload.py --year 1405 --month 6
if errorlevel 1 (
  echo.
  echo Report generation/upload failed.
  pause
  exit /b 1
)
echo.
echo Report generated and uploaded successfully.
pause
