import subprocess
import os
import sys
import time
import psutil
import threading
from pystray import Icon, MenuItem, Menu
from PIL import Image
from filelock import FileLock
import tkinter as tk
from datetime import datetime
from frontend import App, load_config, TEXTS

APPDATA_DIR = os.path.join(os.environ['APPDATA'], 'GameWatcher')
if not os.path.exists(APPDATA_DIR):
    os.makedirs(APPDATA_DIR)

LOCK_PATH = os.path.join(APPDATA_DIR, 'game_watcher.lock')
BACKEND_LOCK = os.path.join(APPDATA_DIR, "backend.lock")

if len(sys.argv) > 1 and sys.argv[1] == "--backend":
    from backend import main as run_backend
    run_backend()
    sys.exit()

lock = FileLock(LOCK_PATH, timeout=1)
try:
    lock.acquire()
except:
    sys.exit()

def get_resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def start_backend():
    if os.path.exists(BACKEND_LOCK):
        try:
            with open(BACKEND_LOCK, "r") as f:
                pid = int(f.read())
            if psutil.pid_exists(pid):
                return 
        except: pass

    subprocess.Popen([sys.executable, "--backend"], 
                     creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)

def open_frontend():
    def _create():
        stats_root = tk.Tk()
        stats_root.attributes("-topmost", True)
        app = App(stats_root)
        stats_root.after(100, lambda: stats_root.attributes("-topmost", False))
        stats_root.mainloop()
    
    threading.Thread(target=_create, daemon=True).start()

def on_quit(icon, item):
    if os.path.exists(BACKEND_LOCK):
        try: os.remove(BACKEND_LOCK)
        except: pass
    icon.stop()
    os._exit(0)

if __name__ == "__main__":
    config = load_config()
    lang = config.get("language", "ja")

    menu_open_label = "統計を開く" if lang == "ja" else "Open Statistics"
    menu_quit_label = "終了" if lang == "ja" else "Quit"

    try:
        icon_image = Image.open(get_resource_path("app_icon.png"))
    except:
        icon_image = Image.new('RGB', (64, 64), (40, 40, 40))

    icon = Icon("GameWatcher", icon_image, menu=Menu(
        MenuItem(menu_open_label, open_frontend),
        MenuItem(menu_quit_label, on_quit)
    ))

    start_backend()
    icon.run()