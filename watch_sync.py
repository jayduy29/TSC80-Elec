# -*- coding: utf-8 -*-
"""
Realtime File Watcher cho file tiến độ điện TCS80 (C459)
Tự động phát hiện khi file Excel được lưu (Ctrl+S), chờ Excel ghi xong 4s,
rồi tự động chạy sync_progress.py và git push lên GitHub ngay lập tức!
"""
import os
import sys
import time
import subprocess
from datetime import datetime

DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(DIR)
MASTER_FILE = os.path.join(DIR, "TCS80-C459-ELECTRICAL PROGRESS.xlsx")
SYNC_SCRIPT = os.path.join(DIR, "sync_progress.py")
PYTHON_EXE = os.path.join(os.path.dirname(sys.executable), "python.exe")
if not os.path.exists(PYTHON_EXE):
    PYTHON_EXE = sys.executable

LOG_FILE = os.path.join(DIR, "auto_sync.log")
LOCK_FILE = os.path.join(DIR, "watcher.lock")

def log(msg):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{now}] [REALTIME_WATCHER] {msg}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

def run_sync():
    try:
        log("⚡ Bắt đầu đồng bộ tự động và đẩy lên GitHub...")
        proc = subprocess.run(
            [PYTHON_EXE, SYNC_SCRIPT],
            cwd=DIR,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace"
        )
        if proc.returncode == 0:
            log("🚀 ĐỒNG BỘ THÀNH CÔNG! Điện thoại và web đã có số liệu mới.")
        else:
            log(f"⚠️ Đồng bộ có lỗi (code {proc.returncode}):\n{proc.stderr}")
    except Exception as e:
        log(f"❌ Ngoại lệ khi chạy đồng bộ: {e}")

def acquire_lock():
    try:
        if os.path.exists(LOCK_FILE):
            # Check if previous process is still alive
            try:
                with open(LOCK_FILE, "r") as f:
                    old_pid = int(f.read().strip())
                # On Windows check if pid exists
                import ctypes
                kernel32 = ctypes.windll.kernel32
                SYNCHRONIZE = 0x00100000
                process = kernel32.OpenProcess(SYNCHRONIZE, False, old_pid)
                if process != 0:
                    kernel32.CloseHandle(process)
                    log(f"⚠️ Watcher đã đang chạy ở tiến trình PID {old_pid}. Không khởi động tiến trình thứ 2.")
                    return False
            except Exception:
                pass
        with open(LOCK_FILE, "w") as f:
            f.write(str(os.getpid()))
        return True
    except Exception as e:
        log(f"⚠️ Không thể tạo lock file: {e}")
        return True

def cleanup_lock():
    try:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)
    except Exception:
        pass

def main():
    if not acquire_lock():
        sys.exit(0)

    log(f"🟢 Realtime Watcher đã khởi động (PID: {os.getpid()})")
    log(f"📁 Đang theo dõi file: {MASTER_FILE}")

    if not os.path.exists(MASTER_FILE):
        log(f"❌ Không tìm thấy file master: {MASTER_FILE}")
        cleanup_lock()
        sys.exit(1)

    last_mtime = os.path.getmtime(MASTER_FILE)
    last_safety_check = time.time()

    try:
        while True:
            time.sleep(2)
            if not os.path.exists(MASTER_FILE):
                continue

            try:
                cur_mtime = os.path.getmtime(MASTER_FILE)
            except OSError:
                continue

            if cur_mtime > last_mtime:
                log(f"🔔 Phát hiện file Excel vừa được lưu lúc {datetime.fromtimestamp(cur_mtime).strftime('%H:%M:%S')}!")
                log("⏳ Chờ 4 giây để Excel giải phóng lock ghi file...")
                time.sleep(4)

                # Kiểm tra xem file có đang tiếp tục được ghi không
                try:
                    mtime2 = os.path.getmtime(MASTER_FILE)
                    if mtime2 != cur_mtime:
                        log("⏳ File vẫn đang được ghi tiếp, chờ thêm 2 giây...")
                        time.sleep(2)
                except OSError:
                    pass

                last_mtime = os.path.getmtime(MASTER_FILE)
                run_sync()
                last_safety_check = time.time()
            else:
                # Mỗi 15 phút kiểm tra trạng thái an toàn
                now_t = time.time()
                if now_t - last_safety_check > 900:
                    last_safety_check = now_t
    except KeyboardInterrupt:
        log("🛑 Watcher dừng lại bởi người dùng.")
    except Exception as e:
        log(f"❌ Lỗi vòng lặp watcher: {e}")
    finally:
        cleanup_lock()

if __name__ == '__main__':
    main()
