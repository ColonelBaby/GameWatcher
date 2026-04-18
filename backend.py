import sys
import os
import time
import json
import psutil
import win32gui
import win32process
from datetime import datetime

if getattr(sys, 'frozen', False):
    os.chdir(sys._MEIPASS)

appdata_dir = os.path.join(os.environ['APPDATA'], 'GameWatcher')
if not os.path.exists(appdata_dir):
    os.makedirs(appdata_dir)

log_path = os.path.join(appdata_dir, "log.txt")
config_path = os.path.join(appdata_dir, "config.json")
LOCK_FILE = os.path.join(appdata_dir, "backend.lock")

HEARTBEAT_PATH = os.path.join(appdata_dir, "heartbeat.txt")

def load_config():
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"blacklist": [], "aliases": {}}

def get_active_window_name():
    try:
        window = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(window)
        process = psutil.Process(pid)
        name = process.name().lower()

        ignored_systems = ["explorer.exe", "taskmgr.exe", "searchhost.exe", "lockapp.exe"]
        if name in ignored_systems:
            return None
            
        return name
    except Exception:
        return "unknown"

import os
import sys
import time
from datetime import datetime


def main():
    if os.path.exists(LOCK_FILE):
        try:
            with open(LOCK_FILE, "r") as f:
                old_pid = int(f.read())
            if psutil.pid_exists(old_pid):
                print("Backend already running.")
                sys.exit()
        except Exception as e:
            with open(os.path.join(appdata_dir, "backend_error.log"), "a") as f:
                f.write(f"{datetime.now()}: {e}\n")

    with open(LOCK_FILE, "w") as f:
        f.write(str(os.getpid()))

    current_app = None
    start_time = None
    
    print("Recording started... (Switch-based mode)")

    try:
        while True:
            
            active_app = get_active_window_name()
            now = datetime.now()

            if active_app != current_app:
                if current_app and start_time:
                    duration = (now - start_time).total_seconds()
                    if duration > 2:
                        with open(log_path, "a", encoding="utf-8") as f:
                            f.write(f"{current_app},{start_time.isoformat()},{now.isoformat()}\n")
                current_app = active_app
                start_time = now

            with open(HEARTBEAT_PATH, "w") as f:
                f.write(str(time.time()))

            time.sleep(1)
    finally:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)

if __name__ == "__main__":
    main()