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
import pylnk3
import psutil

APP_JSON = "apps.json"
APP_TITLE = "起動順＆一括タイマーランチャー"
FOLDER_PREFIX = "folder:"

class AppLauncher(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
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
                    data = json.load(f)
                    self.app_groups = data.get("app_groups", {})
                    self.last_active_tab = data.get("last_active_tab")
            else:
                self.app_groups = {}
                self.last_active_tab = None
        except (json.JSONDecodeError, TypeError) as e:
            messagebox.showerror("エラー", f"設定ファイル({APP_JSON})の読み込みに失敗しました。\n{e}")
            self.app_groups = {}
            self.last_active_tab = None

    def _save_apps(self):
        try:
            data = {
                "app_groups": self.app_groups,
                "last_active_tab": self.last_active_tab
            }
            with open(APP_JSON, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("エラー", f"設定ファイル({APP_JSON})の保存に失敗しました。\n{e}")

    # -----------------------------
    # UI構築
    # -----------------------------
    def _setup_ui(self):
        self.tab_control = ttk.Notebook(self)
        self.tab_control.pack(expand=1, fill='both')

        # 右クリックイベントをバインド
        self.tab_control.bind("<Button-3>", self._on_tab_right_click)

        # 最初に固定タブを作成
        self._create_status_tab()

        # 保存されているアプリタブをstatus_tabの前に追加
        app_tab_names = [name for name in self.app_groups.keys() if name != '起動中一覧']
        for name in app_tab_names:
            self._create_app_tab(name)

        btn_add_tab = tk.Button(self, text="+ タブ追加", command=self._add_new_tab)
        btn_add_tab.pack(side="top", anchor="ne", padx=5, pady=5)

        self.tab_control.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # 最後に使用していたタブを選択
        if self.last_active_tab and self.last_active_tab in self.tabs:
            self.tab_control.select(self.tabs[self.last_active_tab])
        else:
            # デフォルトで最初のタブを選択（「起動中一覧」以外）
            if self.tab_control.tabs():
                first_tab_id = self.tab_control.tabs()[0]
                if self.tab_control.tab(first_tab_id, "text") == "起動中一覧" and len(self.tab_control.tabs()) > 1:
                    self.tab_control.select(self.tab_control.tabs()[1])
                else:
                    self.tab_control.select(first_tab_id)

    def _create_app_tab(self, name):
        frame = tk.Frame(self.tab_control)
        self.tabs[name] = frame
        
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

        canvas.drop_target_register(DND_FILES)
        canvas.dnd_bind('<<Drop>>', lambda e, n=name, f=scroll_frame: self._on_drop_app(e, n, f))

    def _create_status_tab(self):
        status_tab = tk.Frame(self.tab_control)
        self.tab_control.add(status_tab, text="起動中一覧")
        cols = ('グループ名', 'アプリ名', '状態')
        self.status_tree = ttk.Treeview(status_tab, columns=cols, show='headings')
        for c in cols:
            self.status_tree.heading(c, text=c)
        self.status_tree.pack(expand=1, fill='both')

    def _add_app_button(self, parent_frame, tab_name, app_path):
        btn_frame = tk.Frame(parent_frame)
        
        display_name = os.path.basename(app_path)
        command_func = lambda p=app_path, t=tab_name: self._run_single_app(p, t)

        if app_path.startswith(FOLDER_PREFIX):
            folder_path = app_path[len(FOLDER_PREFIX):]
            display_name = os.path.basename(folder_path) + " (フォルダ)"
            command_func = lambda p=folder_path: self._open_folder(p)

        btn = tk.Button(btn_frame, text=display_name, width=30, anchor='w',
                        command=command_func)
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
    def _on_tab_right_click(self, event):
        clicked_tab = self.tab_control.identify(event.x, event.y)
        if not clicked_tab:
            return
        
        try:
            tab_index = self.tab_control.index(clicked_tab)
            tab_name = self.tab_control.tab(tab_index, "text")
        except tk.TclError:
            return # クリック位置がタブでない場合は何もしない

        # 「起動中一覧」タブは削除不可
        if tab_name == "起動中一覧":
            return

        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label=f"タブ '{tab_name}' を削除", command=lambda: self._delete_tab(tab_index, tab_name))
        menu.post(event.x_root, event.y_root)

    def _on_tab_changed(self, event):
        try:
            selected_tab_name = self.tab_control.tab(self.tab_control.select(), "text")
            self.last_active_tab = selected_tab_name # 最後にアクティブだったタブを保存
            self._save_apps() # 設定を保存

            # 「起動中一覧」タブが選択された時のみ更新
            if selected_tab_name == "起動中一覧":
                self._update_status_table()
        except tk.TclError:
            pass # ウィジェット破棄中のエラーを無視

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
            elif os.path.isdir(f): # フォルダの場合
                if tab_name not in self.app_groups:
                    self.app_groups[tab_name] = []
                self.app_groups[tab_name].append(FOLDER_PREFIX + f) # プレフィックスを付けて登録
                self._add_app_button(scroll_frame, tab_name, FOLDER_PREFIX + f)
        self._save_apps()

    def _on_app_right_click(self, event, tab_name, app_path, btn_frame):
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="削除", command=lambda: self._delete_app(tab_name, app_path, btn_frame))
        
        # フォルダの場合は移動メニューを表示しない
        if not app_path.startswith(FOLDER_PREFIX):
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

    def _delete_tab(self, tab_index, tab_name):
        if messagebox.askyesno("タブの削除", f"タブ '{tab_name}' と登録されたすべてのアプリを削除します。\nよろしいですか？"):
            self.tab_control.forget(tab_index)
            if tab_name in self.tabs:
                del self.tabs[tab_name]
            if tab_name in self.app_groups:
                del self.app_groups[tab_name]
            self._save_apps()

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

    def _run_single_app(self, app_path, tab_name):
        if tab_name not in self.app_status:
            self.app_status[tab_name] = []

        try:
            target_path = app_path
            if app_path.lower().endswith('.lnk'):
                link = pylnk3.Lnk(app_path)
                target_path = link.path
            
            subprocess.Popen([target_path])
            self.app_status[tab_name].append({
                'name': app_path, 
                'status': '起動中',
                'target_path': target_path
            })
        except Exception as e:
            self.app_status[tab_name].append({'name': app_path, 'status': f'起動失敗: {e}', 'target_path': None})
        self._update_status_table()

    def _open_folder(self, folder_path):
        try:
            subprocess.Popen(['explorer', folder_path])
        except Exception as e:
            messagebox.showerror("エラー", f"フォルダを開けませんでした: {e}")

    def _kill_apps_in_tab(self, tab_name):
        apps_to_kill = [
            app for app in self.app_status.get(tab_name, []) 
            if app['status'] == '起動中' and app.get('target_path')
        ]
        
        if not apps_to_kill:
            return

        target_paths = {os.path.normcase(app['target_path']) for app in apps_to_kill}

        for proc in psutil.process_iter(['pid', 'exe']):
            try:
                if proc.info['exe'] and os.path.normcase(proc.info['exe']) in target_paths:
                    parent = psutil.Process(proc.info['pid'])
                    children = parent.children(recursive=True)
                    for child in children:
                        child.kill()
                    parent.kill()
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
            except Exception as e:
                print(f"プロセス検索・終了中にエラー: {e}")

        for app in self.app_status.get(tab_name, []):
            if app['status'] == '起動中':
                app['status'] = 'タイマー終了'

    def _update_status_table(self):
        if not self.status_tree:
            return

        try:
            running_exe_paths = {os.path.normcase(p.info['exe']) for p in psutil.process_iter(['exe']) if p.info['exe']}
        except Exception:
            running_exe_paths = set()

        for tab, apps in self.app_status.items():
            for app_status in apps:
                if app_status['status'] == '起動中':
                    target_path = app_status.get('target_path')
                    if not target_path or os.path.normcase(target_path) not in running_exe_paths:
                        app_status['status'] = 'ユーザー終了'

        for item in self.status_tree.get_children():
            self.status_tree.delete(item)
        
        combined_list = []
        for tab, apps in self.app_status.items():
            for a in apps:
                combined_list.append({'tab': tab, 'name': a['name'], 'status': a['status']})
        
        combined_list.sort(key=lambda x: (x['status'] != '起動中', x['status'] != 'タイマー終了', x['status'] != 'ユーザー終了'))
        for app in combined_list:
            display_name = os.path.basename(app['name'])
            if app['name'].startswith(FOLDER_PREFIX):
                display_name = os.path.basename(app['name'][len(FOLDER_PREFIX):]) + " (フォルダ)"
            self.status_tree.insert('', 'end', values=(app['tab'], display_name, app['status']))

    def _periodic_update(self):
        # _update_status_table() は _on_tab_changed でのみ呼び出す
        # self._update_status_table()
        self.after(1000, self._periodic_update)

    # -----------------------------
    # タイマーとトレイ関連
    # -----------------------------
    def _show_window(self):
        self.deiconify()

    def _create_tray_image(self, color="blue"):
        img = Image.new('RGB', (64, 64), 'white')
        d = ImageDraw.Draw(img)
        d.rectangle((0, 0, 63, 63), fill=color)
        return img

    def _set_tab_timer(self, tab_name):
        if self.tab_running_flags.get(tab_name, False):
            messagebox.showinfo("確認", f"{tab_name}のタイマーは既に実行中です。")
            return
        
        is_running = any(app['status'] == '起動中' for app in self.app_status.get(tab_name, []))
        if not is_running:
            messagebox.showinfo("確認", f"{tab_name}で起動中のアプリがありません。先にアプリを起動してください。")
            return

        seconds = simpledialog.askinteger("タブタイマー", f"{tab_name} タブの全アプリ終了までの時間（秒）:", minvalue=1)
        if seconds and seconds > 0:
            threading.Thread(target=self._start_timer_thread, args=(tab_name, seconds), daemon=True).start()
            self.withdraw()

    def _start_timer_thread(self, tab_name, seconds):
        menu = pystray.Menu(pystray.MenuItem('再表示', self._show_window))
        icon = pystray.Icon(tab_name, self._create_tray_image("blue"), f"{tab_name} タイマー", menu=menu)
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
            
            title_text = f"残り: {mins:02d}:{secs:02d} - {APP_TITLE}"
            self.title(title_text)

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
            self.title(APP_TITLE)
            self.deiconify()
            self._update_status_table()
            messagebox.showinfo("タイマー終了", f"{tab_name} のタイマーが終了しました。")

if __name__ == "__main__":
    app = AppLauncher()
    app.mainloop()