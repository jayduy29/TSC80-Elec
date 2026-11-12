@echo off
cd /d "g:\My Drive\_work\towt5\bao cao tuan"
echo [%date% %time%] === BAT DAU DONG BO === >> auto_sync.log
"C:\Users\Admin\AppData\Roaming\uv\python\cpython-3.12-windows-x86_64-none\python.exe" sync_progress.py >> auto_sync.log 2>&1
wscript.exe "run_watcher_silent.vbs"
echo [%date% %time%] === KET THUC === >> auto_sync.log
