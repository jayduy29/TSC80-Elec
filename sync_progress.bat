@echo off
chcp 65001 >nul
echo ========================================================
echo   DONG BO TIEN DO TU TCS80-C459-ELECTRICAL PROGRESS.xlsx
echo ========================================================
echo.
python sync_progress.py
echo.
pause
