# -*- coding: utf-8 -*-
"""
Realtime File Watcher cho các file tiến độ điện TCS80 (C459)
Tự động phát hiện khi file master (TCS80-C459-ELECTRICAL PROGRESS.xlsx) hoặc
bất kỳ file tuần mới/cũ nào (TCS80-C459-ELECTRI* PROGRESS-W*.xlsx) trong folder
được thêm mới hoặc lưu (Ctrl+S), chờ Excel ghi xong 4s, rồi tự động chạy
sync_progress.py để cập nhật C459_ELECTRICAL_PROGRESS_OVERVIEW.xlsx và đẩy lên GitHub!
"""
import os
import re
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
WATCH_RE = re.compile(r'^TCS80-C459-ELECTRI[A-Z]*\s+PROGRESS(?:[-_\s]*W\d{1,2})?\.xlsx$', re.IGNORECASE)


def log(msg):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{now}] [REALTIME_WATCHER] {msg}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def scan_watched_files():
    """Trả về dict {filename: mtime} của tất cả file tiến độ trong folder."""
    state = {}
    try:
        for fname in os.listdir(DIR):
            if fname.startswith('~$') or not fname.lower().endswith('.xlsx'):
                continue
            if WATCH_RE.match(fname):
                fpath = os.path.join(DIR, fname)
                try:
                    state[fname] = os.path.getmtime(fpath)
                except OSError:
                    pass
    except OSError:
        pass
    return state


def run_sync():
    try:
        log("⚡ Bắt đầu đồng bộ tự động tất cả file tuần vào Overview và đẩy lên GitHub...")
        proc = subprocess.run(
            [PYTHON_EXE, SYNC_SCRIPT],
            cwd=DIR,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace"
        )
        if proc.returncode == 0:
            log("🚀 ĐỒNG BỘ THÀNH CÔNG! Overview, điện thoại và web đã có số liệu mới.")
        else:
            log(f"⚠️ Đồng bộ có lỗi (code {proc.returncode}):\n{proc.stderr}\n{proc.stdout}")
    except Exception as e:
        log(f"❌ Ngoại lệ khi chạy đồng bộ: {e}")


def acquire_lock():
    try:
        if os.path.exists(LOCK_FILE):
            try:
                with open(LOCK_FILE, "r") as f:
                    old_pid = int(f.read().strip())
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
    log(f"📁 Đang theo dõi toàn bộ file tiến độ tuần trong: {DIR}")

    last_state = scan_watched_files()
    log(f"📋 Đang theo dõi {len(last_state)} file tiến độ.")

    try:
        while True:
            time.sleep(2)
            cur_state = scan_watched_files()

            changed = [
                fname for fname, mtime in cur_state.items()
                if fname not in last_state or mtime > last_state[fname]
            ]

            if changed:
                log(f"🔔 Phát hiện thay đổi ở {len(changed)} file: {', '.join(changed)}")
                log("⏳ Chờ 4 giây để Excel / Google Drive ghi xong file...")
                time.sleep(4)

                # Chờ thêm nếu file vẫn đang tiếp tục được ghi
                check_state = scan_watched_files()
                if any(check_state.get(f) != cur_state.get(f) for f in changed):
                    log("⏳ File vẫn đang được ghi tiếp, chờ thêm 2 giây...")
                    time.sleep(2)
                    check_state = scan_watched_files()

                last_state = check_state
                run_sync()
    except KeyboardInterrupt:
        log("🛑 Watcher dừng lại bởi người dùng.")
    except Exception as e:
        log(f"❌ Lỗi vòng lặp watcher: {e}")
    finally:
        cleanup_lock()


if __name__ == '__main__':
    main()
