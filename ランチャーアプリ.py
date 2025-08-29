import tkinter as tk
from tkinter import simpledialog, ttk, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import threading
import subprocess
import time
import json
import os
import pystray
from PIL import Image, ImageDraw

APP_JSON = "apps.json"

class AppLauncher(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("起動順＆一括タイマーランチャー")
        self.geometry("700x500")

        # --- データ管理 ---
        self.app_groups = {}
        self.running_processes = {}
        self.app_status = {}
        self.tab_timers = {}
        self.tab_running_flags = {}
        self.tab_tray_icons = {}
        self.status_tree = None
        self.tabs = {}
        self.tab_control = None

        # --- 初期化 ---
        self._load_apps()
        self._setup_ui()
        self._periodic_update()

    # -----------------------------
    # データ永続化
    # -----------------------------
    def _load_apps(self):
        try:
            if os.path.exists(APP_JSON):
                with open(APP_JSON, "r", encoding="utf-8") as f:
                    self.app_groups = json.load(f)
            else:
                self.app_groups = {}
        except (json.JSONDecodeError, TypeError) as e:
            messagebox.showerror("エラー", f"設定ファイル({APP_JSON})の読み込みに失敗しました。\n{e}")
            self.app_groups = {}

    def _save_apps(self):
        try:
            with open(APP_JSON, "w", encoding="utf-8") as f:
                json.dump(self.app_groups, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("エラー", f"設定ファイル({APP_JSON})の保存に失敗しました。\n{e}")

    # -----------------------------
    # UI構築
    # -----------------------------
    def _setup_ui(self):
        self.tab_control = ttk.Notebook(self)
        self.tab_control.pack(expand=1, fill='both')

        # 最初に固定タブを作成
        self._create_status_tab()

        # 保存されているアプリタブをstatus_tabの前に追加
        app_tab_names = [name for name in self.app_groups.keys() if name != '起動中一覧']
        for name in app_tab_names:
            self._create_app_tab(name)

        btn_add_tab = tk.Button(self, text="+ タブ追加", command=self._add_new_tab)
        btn_add_tab.pack(side="top", anchor="ne", padx=5, pady=5)

        self.tab_control.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _create_app_tab(self, name):
        frame = tk.Frame(self.tab_control) # tk.Frameを使用
        self.tabs[name] = frame
        
        # 「起動中一覧」タブの直前に挿入
        status_tab_frame = self.status_tree.master
        insert_pos = self.tab_control.index(status_tab_frame)
        self.tab_control.insert(insert_pos, frame, text=name)

        canvas = tk.Canvas(frame)
        scrollbar = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas)
        scroll_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self._refresh_tab_buttons(name)

        # Canvasにドロップイベントをバインド
        canvas.drop_target_register(DND_FILES)
        canvas.dnd_bind('<<Drop>>', lambda e, n=name, f=scroll_frame: self._on_drop_app(e, n, f))

    def _create_status_tab(self):
        status_tab = tk.Frame(self.tab_control) # tk.Frameを使用
        self.tab_control.add(status_tab, text="起動中一覧")
        cols = ('グループ名', 'アプリ名', '状態')
        self.status_tree = ttk.Treeview(status_tab, columns=cols, show='headings')
        for c in cols:
            self.status_tree.heading(c, text=c)
        self.status_tree.pack(expand=1, fill='both')

    def _add_app_button(self, parent_frame, tab_name, app_path):
        btn_frame = tk.Frame(parent_frame)
        
        btn = tk.Button(btn_frame, text=os.path.basename(app_path), width=30, anchor='w',
                        command=lambda t=tab_name: self._run_apps_in_tab(t))
        btn.pack(side="left", fill="x", expand=True)
        
        btn.bind("<Button-3>", lambda e, n=tab_name, a=app_path, b=btn_frame: self._on_app_right_click(e, n, a, b))
        btn_frame.pack(fill='x', padx=5, pady=2)

    def _refresh_tab_buttons(self, tab_name):
        frame = self.tabs[tab_name]
        try:
            scroll_frame = frame.winfo_children()[0].winfo_children()[0]
        except IndexError:
            return
            
        for w in scroll_frame.winfo_children():
            w.destroy()
        for app in self.app_groups.get(tab_name, []):
            self._add_app_button(scroll_frame, tab_name, app)
        
        btn_timer = tk.Button(scroll_frame, text="タイマー付き一括終了", command=lambda n=tab_name: self._set_tab_timer(n))
        btn_timer.pack(pady=10)

    # -----------------------------
    # イベントハンドラ
    # -----------------------------
    def _on_tab_changed(self, event):
        try:
            selected_tab_name = self.tab_control.tab(self.tab_control.select(), "text")
        except tk.TclError:
            return
        for tab_name, tray_icon in self.tab_tray_icons.items():
            if self.tab_running_flags.get(tab_name, False):
                color = "green" if tab_name == selected_tab_name else "blue"
                tray_icon.icon = self._create_tray_image(color)

    def _on_drop_app(self, event, tab_name, scroll_frame):
        files = self.tk.splitlist(event.data)
        for f in files:
            if os.path.isfile(f):
                if tab_name not in self.app_groups:
                    self.app_groups[tab_name] = []
                self.app_groups[tab_name].append(f)
                self._add_app_button(scroll_frame, tab_name, f)
        self._save_apps()

    def _on_app_right_click(self, event, tab_name, app_path, btn_frame):
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="削除", command=lambda: self._delete_app(tab_name, app_path, btn_frame))
        index = self.app_groups[tab_name].index(app_path)
        if index > 0:
            menu.add_command(label="↑上に移動", command=lambda: self._move_app(tab_name, index, -1))
        if index < len(self.app_groups[tab_name]) - 1:
            menu.add_command(label="↓下に移動", command=lambda: self._move_app(tab_name, index, 1))
        menu.post(event.x_root, event.y_root)

    # -----------------------------
    # アプリケーションロジック
    # -----------------------------
    def _add_new_tab(self):
        name = simpledialog.askstring("新規タブ", "タブ名を入力:")
        if name and name not in self.app_groups:
            self.app_groups[name] = []
            self._save_apps()
            self._create_app_tab(name)

    def _delete_app(self, tab_name, app_path, btn_frame):
        if messagebox.askyesno("確認", f"{os.path.basename(app_path)}を削除しますか？"):
            if app_path in self.app_groups.get(tab_name, []):
                self.app_groups[tab_name].remove(app_path)
                self._save_apps()
                btn_frame.destroy()

    def _move_app(self, tab_name, index, direction):
        apps = self.app_groups[tab_name]
        new_index = index + direction
        apps[index], apps[new_index] = apps[new_index], apps[index]
        self._save_apps()
        self._refresh_tab_buttons(tab_name)

    def _run_apps_in_tab(self, tab_name):
        self.running_processes[tab_name] = []
        self.app_status[tab_name] = []
        for app_path in self.app_groups.get(tab_name, []):
            try:
                proc = subprocess.Popen(app_path)
                self.running_processes[tab_name].append(proc)
                self.app_status[tab_name].append({'name': app_path, 'proc': proc, 'status': '起動中'})
            except Exception as e:
                self.app_status[tab_name].append({'name': app_path, 'proc': None, 'status': f'起動失敗: {e}'})
        self._update_status_table()

    def _kill_apps_in_tab(self, tab_name):
        for proc in self.running_processes.get(tab_name, []):
            try:
                proc.terminate()
            except Exception:
                pass
        for app in self.app_status.get(tab_name, []):
            if app['status'] == '起動中':
                app['status'] = 'タイマー終了'
        self.running_processes[tab_name] = []

    def _update_status_table(self):
        if not self.status_tree:
            return
        for item in self.status_tree.get_children():
            self.status_tree.delete(item)
        
        combined_list = []
        for tab, apps in self.app_status.items():
            for a in apps:
                if a.get('proc') and a['status'] == '起動中':
                    if a['proc'].poll() is not None:
                        a['status'] = 'ユーザー終了'
                combined_list.append({'tab': tab, 'name': a['name'], 'status': a['status']})
        
        combined_list.sort(key=lambda x: (x['status'] != '起動中', x['status'] != 'タイマー終了'))
        for app in combined_list:
            self.status_tree.insert('', 'end', values=(app['tab'], os.path.basename(app['name']), app['status']))

    def _periodic_update(self):
        self._update_status_table()
        self.after(1000, self._periodic_update)

    # -----------------------------
    # タイマーとトレイ関連
    # -----------------------------
    def _create_tray_image(self, color="blue"):
        img = Image.new('RGB', (64, 64), 'white')
        d = ImageDraw.Draw(img)
        d.rectangle((0, 0, 63, 63), fill=color)
        return img

    def _set_tab_timer(self, tab_name):
        if self.tab_running_flags.get(tab_name, False):
            messagebox.showinfo("確認", f"{tab_name}のタイマーは既に実行中です。")
            return
        seconds = simpledialog.askinteger("タブタイマー", f"{tab_name} タブの全アプリ終了までの時間（秒）:", minvalue=1)
        if seconds and seconds > 0:
            self._run_apps_in_tab(tab_name)
            threading.Thread(target=self._start_timer_thread, args=(tab_name, seconds), daemon=True).start()

    def _start_timer_thread(self, tab_name, seconds):
        icon = pystray.Icon(tab_name, self._create_tray_image("blue"), f"{tab_name} タイマー")
        self.tab_tray_icons[tab_name] = icon
        
        threading.Thread(target=icon.run, daemon=True).start()

        self.tab_timers[tab_name] = seconds
        self.tab_running_flags[tab_name] = True
        
        self.after(100, lambda: self._timer_countdown(tab_name))

    def _timer_countdown(self, tab_name):
        if self.tab_timers.get(tab_name, 0) > 0 and self.tab_running_flags.get(tab_name, False):
            self.tab_timers[tab_name] -= 1
            
            seconds = self.tab_timers[tab_name]
            mins, secs = divmod(seconds, 60)
            if self.tab_tray_icons.get(tab_name):
                self.tab_tray_icons[tab_name].title = f"{tab_name} 残り: {mins:02d}:{secs:02d}"

            self.after(1000, lambda: self._timer_countdown(tab_name))
        else:
            self._stop_timer(tab_name)

    def _stop_timer(self, tab_name):
        if self.tab_running_flags.get(tab_name, False):
            self.tab_running_flags[tab_name] = False
            if self.tab_tray_icons.get(tab_name):
                self.tab_tray_icons[tab_name].stop()
                del self.tab_tray_icons[tab_name]
            self._kill_apps_in_tab(tab_name)
            self._update_status_table()
            messagebox.showinfo("タイマー終了", f"{tab_name} のタイマーが終了しました。")

if __name__ == "__main__":
    app = AppLauncher()
    app.mainloop()
