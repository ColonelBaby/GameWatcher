import os
import json
import re
import time
import psutil
import subprocess
import winreg
import sys
import tkinter as tk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import ttk, simpledialog, messagebox
from datetime import datetime, timedelta
from collections import defaultdict

appdata_dir = os.path.join(os.environ['APPDATA'], 'GameWatcher')
if not os.path.exists(appdata_dir):
    os.makedirs(appdata_dir)
log_path = os.path.join(appdata_dir, "log.txt")
config_path = os.path.join(appdata_dir, "config.json")
HEARTBEAT_PATH = os.path.join(appdata_dir, "heartbeat.txt")

def get_resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

TEXTS = {
    "ja": {
        "title": "GameWatcher",
        "tab_stats": " 統計 ",
        "tab_settings": " 設定 ",
        "tab_total": "一週間の合計",
        "stats_title": "今週のプレイ時間",
        "col_name": "アプリ名 (ダブルクリックで名前変更)",
        "col_time": "プレイ時間",
        "btn_refresh": "更新",
        "lang_label": "言語設定 (Language)",
        "blacklist_label": "無視リスト（記録しないアプリ）",
        "btn_remove": "選択した項目を削除",
        "blacklist_note": "※統計画面のアプリを右クリックして追加できます",
        "msg_restart": "設定を保存しました。再起動後に反映されます。",
        "msg_confirm_add": "このアプリを無視リストに追加しますか？",
        "msg_confirm_delete": "無視リストから削除しますか？",
        "ctx_add_blacklist": "このアプリを無視リストに追加",
        "days": ["月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日", "日曜日"],
        "tab_pie": " 円グラフ ",
        "tab_bar": " 棒グラフ ",
        "label_no_data": "データがありません",
        "unit_min": "分",
        "menu_open": "統計を開く",
        "menu_quit": "終了",
    },
    "en": { 
        "title": "GameWatcher",
        "tab_stats": " Stats ",
        "tab_settings": " Settings ",
        "tab_total": "Weekly Total",
        "stats_title": "Playtime this week",
        "col_name": "App Name (Double-click to rename)",
        "col_time": "Play Time",
        "btn_refresh": "Refresh",
        "lang_label": "Language Setting",
        "blacklist_label": "Blacklist (Apps to ignore)",
        "btn_remove": "Remove Selected",
        "blacklist_note": "*Add apps via right-click in the Stats tab",
        "msg_restart": "Settings saved. Please restart to apply changes.",
        "msg_confirm_add": "Add this app to blacklist?",
        "msg_confirm_delete": "Remove from blacklist?",
        "ctx_add_blacklist": "Add this app to blacklist",
        "days": ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
        "tab_pie": " Pie Chart ",
        "tab_bar": " Bar Chart ",
        "label_no_data": "No Data",
        "unit_min": "min",
        "menu_open": "Open Stats",
        "menu_quit": "Quit",
    }
}

def load_config():
    config = {"aliases": {}, "language": "en", "blacklist": [], "first_run": True}
    
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                loaded_data = json.load(f)
                config.update(loaded_data)
                config["first_run"] = False
        except:
            pass        
    return config

def save_config(config):
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4, ensure_ascii=False)

class App:
    def __init__(self, root):
        self.root = root
        self.config = load_config()
        self.lang = self.config.get("language", "ja")
        self.root.title(TEXTS[self.lang]["title"])

        try:
            icon_path = get_resource_path("app_icon.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Icon error: {e}")

        if self.config.get("first_run"):
            self.autostart_var = tk.BooleanVar(value=True)
            self.toggle_autostart()

            self.config["first_run"] = False
            save_config(self.config)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.root.title(TEXTS[self.lang]["title"])

        window_width = 1200
        window_height = 900

        screen_width = self.root.winfo_screenwidth()
        x = (screen_width // 2) - (window_width // 2)

        y = 60 

        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        self.day_trees = {}
        self.day_canvases = {}

        self.tab_control = ttk.Notebook(self.root)
        self.tab_stats = ttk.Frame(self.tab_control)
        self.tab_settings = ttk.Frame(self.tab_control)
        
        self.tab_control.add(self.tab_stats, text=TEXTS[self.lang]["tab_stats"])
        self.tab_control.add(self.tab_settings, text=TEXTS[self.lang]["tab_settings"])
        self.tab_control.pack(expand=1, fill="both")
        
        self.create_stats_widgets()
        self.create_settings_widgets()

        self.cleanup_old_logs()

        self.root.update()
        self.refresh_ui()

    def create_stats_widgets(self):
        tk.Label(self.tab_stats, text=TEXTS[self.lang]["stats_title"], font=("MS Gothic", 20, "bold")).pack(pady=10)

        self.day_tabs = ttk.Notebook(self.tab_stats)
        self.day_tabs.pack(fill="both", expand=True, padx=10, pady=5)

        self.day_trees = {} 

        self.status_frame = tk.Frame(self.tab_stats)
        self.status_frame.pack(anchor="ne", padx=20)
        
        self.status_lamp = tk.Canvas(self.status_frame, width=12, height=12)
        self.status_lamp.pack(side="left", padx=5)
        self.status_label = tk.Label(self.status_frame, text="Backend: Checking...", font=("Consolas", 9))
        self.status_label.pack(side="left")

        self.update_status_indicator()

        total_frame = ttk.Frame(self.day_tabs)
        self.day_tabs.add(total_frame, text=TEXTS[self.lang]["tab_total"])
        self.day_trees["total"] = self.create_treeview(total_frame)

        for i, day_name in enumerate(TEXTS[self.lang]["days"]):
            day_frame = ttk.Frame(self.day_tabs)
            self.day_tabs.add(day_frame, text=day_name)
            self.day_trees[i] = self.create_treeview(day_frame)

        self.graph_tabs = ttk.Notebook(self.tab_stats)
        self.graph_tabs.pack(fill="both", expand=True, padx=10, pady=5)

        self.tab_pie = ttk.Frame(self.graph_tabs)
        self.tab_bar = ttk.Frame(self.graph_tabs)

        self.graph_tabs.add(self.tab_pie, text=TEXTS[self.lang]["tab_pie"])
        self.graph_tabs.add(self.tab_bar, text=TEXTS[self.lang]["tab_bar"])

        self.day_tabs.bind("<<NotebookTabChanged>>", lambda e: self.refresh_ui())
        self.graph_tabs.bind("<<NotebookTabChanged>>", lambda e: self.refresh_ui())

        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label=TEXTS[self.lang]["ctx_add_blacklist"], command=self.add_to_blacklist_from_menu)

        ttk.Button(self.tab_stats, text=TEXTS[self.lang]["btn_refresh"], command=self.refresh_ui).pack(pady=5)

    def create_day_content(self, parent):
        # グラフ
        canvas = tk.Canvas(parent, bg="#f8f9fa", height=350)
        canvas.pack(fill="both", expand=True, padx=10, pady=5)

        canvas.bind("<Configure>", lambda e: self.on_resize())
        
        # 表
        tree = ttk.Treeview(parent, columns=("Name", "Time"), show="headings")
        tree.heading("Name", text=TEXTS[self.lang]["col_name"])
        tree.heading("Time", text=TEXTS[self.lang]["col_time"])
        tree.column("Name", width=400)
        tree.column("Time", width=150)
        tree.pack(fill="both", expand=True, padx=10, pady=5)
        
        tree.bind("<Double-1>", self.rename_alias)
        tree.bind("<Button-3>", self.show_context_menu)
        
        return tree, canvas

    def on_resize(self):
        if hasattr(self, "_resize_timer"):
            self.root.after_cancel(self._resize_timer)
        
        self._resize_timer = self.root.after(200, self.refresh_ui)

    def on_closing(self):
        if hasattr(self, 'after_id'):
            self.root.after_cancel(self.after_id)
        self.root.destroy()

    def check_backend_status(self):
        if os.path.exists(HEARTBEAT_PATH):
            mtime = os.path.getmtime(HEARTBEAT_PATH)
            if time.time() - mtime < 5:
                self.status_label.config(text="backend Active", fg="green")
                return
        self.status_label.config(text="backend Inactive", fg="red")

    def refresh_ui(self):
        stats = self.get_week_data()
        keys = ["total"] + list(range(7))
        
        for k in keys:
            tree = self.day_trees[k]
            for item in tree.get_children(): tree.delete(item)
            
            current_data = stats["total"] if k == "total" else stats["days"][k]
            last_comp_data = stats["last_total"] if k == "total" else stats["last_days"][k]
            
            sorted_data = sorted(current_data.items(), key=lambda x: x[1], reverse=True)
            
            for name, sec_total in sorted_data:
                h, m = int(sec_total) // 3600, (int(sec_total) % 3600) // 60
                time_str = f"{h}h {m}m"
                
                last_sec = last_comp_data.get(name, 0)
                diff_sec = sec_total - last_sec
                diff_m = int(abs(diff_sec)) // 60
                
                if last_sec == 0:
                    diff_str = "NEW"
                else:
                    sign = "+" if diff_sec >= 0 else "-"
                    diff_str = f"{sign}{diff_m}m"

                tree.insert("", tk.END, values=(name, time_str, diff_str))

        current_tab_id = self.day_tabs.index(self.day_tabs.select())
        if current_tab_id == 0:
            self.draw_graph(stats["total"], stats["last_total"])
        else:
            idx = current_tab_id - 1
            self.draw_graph(stats["days"][idx], stats["last_days"][idx])

    def _get_pie_data(self, data_dict):
        """円グラフ用のラベルと値を整理する補助関数"""
        if not data_dict: return [], []
        sorted_items = sorted(data_dict.items(), key=lambda x: x[1], reverse=True)
        total_time = sum(data_dict.values())
        labels, values, others_time = [], [], 0
        threshold = 0.05 
        for name, duration in sorted_items:
            if (duration / total_time) < threshold:
                others_time += duration
            else:
                labels.append(name)
                values.append(duration / 60)
        if others_time > 0:
            labels.append("Others" if self.lang == "en" else "その他")
            values.append(others_time / 60)
        return labels, values

    def draw_graph(self, data, last_data=None):
        for widget in self.tab_pie.winfo_children(): widget.destroy()
        for widget in self.tab_bar.winfo_children(): widget.destroy()

        if not data and not last_data:
            tk.Label(self.tab_pie, text="No Data").pack(expand=True)
            return

        plt.rcParams['font.family'] = 'MS Gothic'
        selected_graph = self.graph_tabs.index(self.graph_tabs.select())

        if selected_graph == 0:
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4), dpi=100)
            
            if last_data:
                l_labels, l_vals = self._get_pie_data(last_data)
                ax1.pie(l_vals, labels=l_labels, autopct='%1.1f%%', startangle=90, counterclock=False)
                ax1.set_title("Last Week" if self.lang=="en" else "先週")
            else:
                ax1.text(0.5, 0.5, "No Data", ha='center')

            if data:
                c_labels, c_vals = self._get_pie_data(data)
                ax2.pie(c_vals, labels=c_labels, autopct='%1.1f%%', startangle=90, counterclock=False)
                ax2.set_title("This Week" if self.lang=="en" else "今週")
            else:
                ax2.text(0.5, 0.5, "No Data", ha='center')
            
            canvas = FigureCanvasTkAgg(fig, master=self.tab_pie)

        else:
            fig, ax = plt.subplots(figsize=(6, 4), dpi=100)
            
            all_apps = set(data.keys()) | set(last_data.keys() if last_data else [])
            sorted_labels = sorted(all_apps, key=lambda x: data.get(x, 0) or (last_data.get(x, 0) if last_data else 0), reverse=True)[:8][::-1]
            
            if last_data:
                ax.barh(sorted_labels, [last_data.get(l, 0)/60 for l in sorted_labels], color='gray', alpha=0.3, label="Last Week")
            if data:
                ax.barh(sorted_labels, [data.get(l, 0)/60 for l in sorted_labels], color='#4da6ff', alpha=0.7, label="This Week")
            
            ax.legend()
            plt.tight_layout()
            canvas = FigureCanvasTkAgg(fig, master=self.tab_bar)

        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def create_treeview(self, parent):
        tree = ttk.Treeview(parent, columns=("Name", "Time","Diff"), show="headings")
        tree.heading("Name", text=TEXTS[self.lang]["col_name"])
        tree.heading("Time", text=TEXTS[self.lang]["col_time"])
        tree.column("Name", width=400)
        tree.column("Time", width=150)
        tree.pack(fill="both", expand=True, padx=5, pady=5)
        tree.bind("<Double-1>", self.rename_alias)
        tree.bind("<Button-3>", self.show_context_menu)
        return tree

    def create_settings_widgets(self):
        container = tk.Frame(self.tab_settings, padx=20, pady=20)
        container.pack(fill="both", expand=True)
        
        tk.Label(container, text=TEXTS[self.lang]["lang_label"], font=("MS Gothic", 12, "bold")).pack(anchor="w")
        self.lang_combo = ttk.Combobox(container, values=["English (en)", "日本語 (ja)"], state="readonly")
        self.lang_combo.set("English (en)" if self.lang == "en" else "日本語 (jp)")
        self.lang_combo.pack(anchor="w", pady=5)
        self.lang_combo.bind("<<ComboboxSelected>>", self.change_language)
        
        self.autostart_var = tk.BooleanVar()
        self.autostart_var.set(self.is_autostart_enabled())
        
        autostart_label = "Run GameWatcher at Startup" if self.lang == "en" else "システム起動時に実行する"
        self.autostart_check = tk.Checkbutton(
            container, text=autostart_label, variable=self.autostart_var, 
            command=self.toggle_autostart, font=("MS Gothic", 10)
        )
        self.autostart_check.pack(anchor="w", pady=10)

        tk.Label(container, text=TEXTS[self.lang]["blacklist_label"], font=("MS Gothic", 12, "bold")).pack(anchor="w", pady=(15, 0))
        
        list_frame = tk.Frame(container)
        list_frame.pack(fill="both", expand=True, pady=10)
        self.blacklist_box = tk.Listbox(list_frame, font=("Consolas", 10))
        self.blacklist_box.pack(side="left", fill="both", expand=True)
        scroll = tk.Scrollbar(list_frame)
        scroll.pack(side="right", fill="y")
        self.blacklist_box.config(yscrollcommand=scroll.set)
        scroll.config(command=self.blacklist_box.yview)
        
        ttk.Button(container, text=TEXTS[self.lang]["btn_remove"], command=self.remove_from_blacklist).pack(anchor="w", pady=5)
        tk.Label(container, text=TEXTS[self.lang]["blacklist_note"], font=("MS Gothic", 8), fg="gray").pack(anchor="w", pady=10)

    def get_current_tree(self):
        tab_id = self.day_tabs.index(self.day_tabs.select())
        if tab_id == 0: return self.day_trees["total"]
        return self.day_trees[tab_id - 1]

    def show_context_menu(self, event):
        tree = event.widget
        item = tree.identify_row(event.y)
        if item:
            tree.selection_set(item)
            tree.focus(item)
            self.context_menu.post(event.x_root, event.y_root)

    def add_to_blacklist_from_menu(self):
        tree = self.get_current_tree()
        item = tree.focus()
        if not item: return
        display_name = tree.item(item, "values")[0]
        
        target_name = display_name
        for original_name, alias in self.config["aliases"].items():
            if alias == display_name:
                target_name = original_name
                break
        
        if messagebox.askyesno("Confirm", TEXTS[self.lang]["msg_confirm_add"]):
            if target_name not in self.config["blacklist"]:
                self.config["blacklist"].append(target_name)
                save_config(self.config)
                self.refresh_ui()

    def change_language(self, event):
        new_lang = "ja" if "日本語" in self.lang_combo.get() else "en"
        self.config["language"] = new_lang
        save_config(self.config)
        messagebox.showinfo("Setting", TEXTS[new_lang]["msg_restart"])

    def get_week_data(self):
        now = datetime.now()
        start_of_week = (now - timedelta(days=now.weekday())).replace(hour=0, minute=0, second=0, microsecond=0)
        start_of_last_week = start_of_week - timedelta(days=7)
        limit_date = now - timedelta(days=31)

        stats = {
            "total": defaultdict(float),
            "last_total": defaultdict(float),
            "days": {i: defaultdict(float) for i in range(7)},
            "last_days": {i: defaultdict(float) for i in range(7)}
        }
        
        if not os.path.exists(log_path): return stats
        
        processed_logs = set()
        try:
            with open(log_path, "r", encoding="utf-8") as f:
                for line in f:
                    parts = line.strip().split(",")
                    if len(parts) != 3: continue
                    name_raw, s_str, e_str = parts
                    
                    log_key = f"{name_raw}_{s_str}"
                    if log_key in processed_logs: continue
                    processed_logs.add(log_key)

                    try:
                        start = datetime.fromisoformat(s_str)
                        end = datetime.fromisoformat(e_str)
                    except: continue

                    if end < limit_date: continue

                    clean_key = re.sub(r'(?i)\.exe$', '', name_raw).strip().lower()
                    display_name = self.config.get("aliases", {}).get(name_raw, clean_key)

                    if name_raw in self.config.get("blacklist", []) or display_name in self.config.get("blacklist", []):
                        continue

                    duration = (end - start).total_seconds()
                    if duration <= 0: continue

                    if end >= start_of_week:
                        actual_dur = (end - max(start, start_of_week)).total_seconds()
                        stats["total"][display_name] += actual_dur
                        stats["days"][end.weekday()][display_name] += actual_dur
                    
                    elif end >= start_of_last_week:
                        actual_dur = (min(end, start_of_week) - max(start, start_of_last_week)).total_seconds()
                        stats["last_total"][display_name] += actual_dur
                        stats["last_days"][end.weekday()][display_name] += actual_dur

            for d_set in [stats["total"], stats["last_total"]]:
                to_del = [n for n, t in d_set.items() if t < 60]
                for n in to_del: del d_set[n]
            for d_map in [stats["days"], stats["last_days"]]:
                for i in range(7):
                    to_del = [n for n, t in d_map[i].items() if t < 60]
                    for n in to_del: del d_map[i][n]

        except Exception as e:
            print(f"Error: {e}")
        return stats

    def rename_alias(self, event):
        tree = event.widget
        item = tree.focus()
        if not item: return
        old_name = tree.item(item, "values")[0]
        title = "Rename" if self.lang=="en" else "名前変更"
        prompt = f"New name for '{old_name}':" if self.lang=="en" else f"'{old_name}' の新しい表示名:"
        new_name = simpledialog.askstring(title, prompt, parent=self.root)
        if new_name:
            self.config["aliases"][old_name] = new_name
            save_config(self.config)
            self.refresh_ui()

    def remove_from_blacklist(self):
        selected = self.blacklist_box.curselection()
        if not selected: return
        name = self.blacklist_box.get(selected[0])
        if messagebox.askyesno("Confirm", TEXTS[self.lang]["msg_confirm_delete"]):
            self.config["blacklist"].remove(name)
            save_config(self.config)
            self.refresh_ui()

    def cleanup_old_logs(self):
        """1年以上前のデータを物理的に削除する"""
        if not os.path.exists(log_path): return
        
        one_year_ago = datetime.now() - timedelta(days=365)
        new_lines = []
        modified = False

        try:
            with open(log_path, "r", encoding="utf-8") as f:
                for line in f:
                    parts = line.strip().split(",")
                    if len(parts) != 3: continue
                    try:
                        end_time = datetime.fromisoformat(parts[2])
                        if end_time > one_year_ago:
                            new_lines.append(line)
                        else:
                            modified = True
                    except: continue
            
            if modified:
                with open(log_path, "w", encoding="utf-8") as f:
                    f.writelines(new_lines)
        except Exception as e:
            print(f"Cleanup error: {e}")

    def update_status_indicator(self):
        heartbeat_path = os.path.join(appdata_dir, "heartbeat.txt")
        is_active = False
        
        if os.path.exists(heartbeat_path):
            last_heartbeat = os.path.getmtime(heartbeat_path)
            if time.time() - last_heartbeat < 10:
                is_active = True
        
        if is_active:
            self.status_lamp.delete("all")
            self.status_lamp.create_oval(2, 2, 12, 12, fill="#00ff00", outline="#00aa00")
            self.status_label.config(text="Backend: Active", fg="#008800")
            pass
        else:
            self.status_lamp.delete("all")
            self.status_lamp.create_oval(2, 2, 12, 12, fill="#ff0000", outline="#aa0000")
            self.status_label.config(text="Backend: Inactive", fg="#880000")
            pass

        self.after_id = self.root.after(5000, self.update_status_indicator)

    def is_autostart_enabled(self):
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ)
            value, _ = winreg.QueryValueEx(key, "GameWatcher")
            winreg.CloseKey(key)
            return True
        except WindowsError:
            return False

    def toggle_autostart(self):
        """チェックボックスの状態に合わせてレジストリを更新する"""
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        app_name = "GameWatcher"
        
        if getattr(sys, 'frozen', False):
            app_path = sys.executable
        else:
            main_script = os.path.join(appdata_dir, "game_watcher.py")
            app_path = f'"{sys.executable}" "{main_script}"'

        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)
            if self.autostart_var.get():
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, app_path)
            else:
                try:
                    winreg.DeleteValue(key, app_name)
                except FileNotFoundError:
                    pass
            winreg.CloseKey(key)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update startup setting: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()