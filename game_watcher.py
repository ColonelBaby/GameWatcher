import subprocess
import os
import sys
import time
import psutil
import threading
from pystray import Icon, MenuItem, Menu
from PIL import Image
from filelock import FileLock
from frontend import App 
import tkinter as tk
from datetime import datetime

lock_path = os.path.join(os.environ['APPDATA'], 'GameWatcher', 'game_watcher.lock')
lock = FileLock(lock_path, timeout=1)

if len(sys.argv) > 1 and sys.argv[1] == "--backend":
    from backend import main as run_backend
    run_backend()
    sys.exit()

try:
    lock.acquire()
except:
    sys.exit()

appdata_dir = os.path.join(os.environ['APPDATA'], 'GameWatcher')
if not os.path.exists(appdata_dir):
    os.makedirs(appdata_dir)
LOCK_FILE = os.path.join(appdata_dir, "backend.lock")

def start_backend():
    if os.path.exists(LOCK_FILE):
        try:
            with open(LOCK_FILE, "r") as f:
                pid = int(f.read())
            if psutil.pid_exists(pid):
                return 
        except: pass

    subprocess.Popen([sys.executable, "--backend"], 
                     creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)

def open_stats():
    stats_root = tk.Tk()
    app = App(stats_root)
    stats_root.mainloop()

def open_frontend():
    def _create():
        stats_root = tk.Tk()
        stats_root.attributes("-topmost", True)
        app = App(stats_root)
        stats_root.after(100, lambda: stats_root.attributes("-topmost", False))
        stats_root.mainloop()
    
    threading.Thread(target=_create, daemon=True).start()

def on_quit(icon, item):
    if os.path.exists(LOCK_FILE):
        try:
            os.remove(LOCK_FILE)
        except:
            pass
    icon.stop()
    os._exit(0)

# アイコン作成
image = Image.new('RGB', (64, 64), (40, 40, 40))

icon = Icon("GameWatcher", image, menu=Menu(
    MenuItem("統計を開く", open_frontend),
    MenuItem("終了", on_quit)
))

def get_resource_path(relative_path):
    """ exe化後の一時フォルダからのパス解決 """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

try:
    icon_image = Image.open(get_resource_path("app_icon.png"))
except Exception:
    icon_image = Image.new('RGB', (64, 64), (40, 40, 40))

icon = Icon("GameWatcher", icon_image, menu=Menu(
    MenuItem("統計を開く", open_frontend),
    MenuItem("終了", on_quit)
))

if __name__ == "__main__":
    start_backend()
    icon.run()