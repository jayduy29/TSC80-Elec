@echo off
cd /d "g:\My Drive\_work\towt5\bao cao tuan"
echo [%date% %time%] === BAT DAU TU DONG DONG BO === >> auto_sync.log
"C:\Users\Admin\.local\bin\python.exe" sync_progress.py >> auto_sync.log 2>&1
echo [%date% %time%] === KET THUC === >> auto_sync.log
echo. >> auto_sync.log
