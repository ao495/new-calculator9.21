import os
import sys
import time
import json
import winreg
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from pathlib import Path

# PyWin32が必要
try:
    import win32com.client
except ModuleNotFoundError:
    if messagebox.askokcancel("モジュールエラー", "pywin32 が見つかりません。\nインストールしますか？"):
        import subprocess
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
            import win32com.client
        except Exception as e:
            messagebox.showerror("インストールエラー", f"pywin32 のインストールに失敗しました。\n{e}")
            sys.exit(1)
    else:
        sys.exit(1)

# PyInstaller用パス取得
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return Path(base_path) / relative_path

# --- 設定 ---
APP_NAME = "HistoryGuardApp"
CONFIG_FILE = Path(os.getenv('APPDATA')) / APP_NAME / "config.json"
ICON_FILE = resource_path("icon.ico")
recent_path = Path(os.getenv('APPDATA')) / "Microsoft" / "Windows" / "Recent"

# --- 履歴非表示 ---
def suppress_history(active: bool):
    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                               r"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced")
        winreg.SetValueEx(key, "DisableRecentDocsHistory", 0, winreg.REG_DWORD, 1 if active else 0)
        return True, "非履歴モード ON" if active else "非履歴モード OFF"
    except Exception as e:
        return False, f"履歴設定エラー: {e}"

# --- タイマー ---
def start_timer(duration_min: int, callback):
    def run():
        time.sleep(duration_min * 60)
        callback()
    threading.Thread(target=run, daemon=True).start()

# --- 最近使ったファイル取得 ---
def get_recent_files():
    shell = win32com.client.Dispatch("WScript.Shell")
    files = []
    try:
        for item in recent_path.glob("*.lnk"):
            try:
                shortcut = shell.CreateShortcut(str(item))
                target = Path(shortcut.TargetPath)
                # ターゲットが存在しなくてもLNKファイル自体は表示する
                last_modified = item.stat().st_mtime
                last_modified_str = datetime.fromtimestamp(last_modified).strftime('%Y-%m-%d %H:%M')
                files.append((target.name, str(target), last_modified, last_modified_str, item))
            except Exception:
                continue
    except Exception as e:
        messagebox.showwarning("警告", f"履歴ファイル読み込みエラー: {e}")
    files.sort(key=lambda x: x[2], reverse=True)
    return files

# --- 設定読み書き ---
def load_config():
    default_config = {
        "manual_mode": True,
        "timer_minutes": 60,
        "show_history": False,
        "show_details": False,
        "auto_refresh": True,
        "refresh_interval": 60,
        "window_size": "520x720"
    }
    if not CONFIG_FILE.exists():
        CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
        save_config(default_config)
        return default_config
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)
        for k, v in default_config.items():
            if k not in config:
                config[k] = v
        return config
    except Exception as e:
        messagebox.showwarning("設定読み込みエラー", f"デフォルト設定を使用します。\n{e}")
        return default_config

def save_config(config):
    try:
        CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        messagebox.showwarning("設定保存エラー", f"設定保存失敗: {e}")
        return False

# --- GUI ---
class HistoryGuardApp:
    def __init__(self, master):
        self.master = master
        self.config = load_config()
        master.title("履歴非表示モード＋タイマー")
        master.geometry(self.config["window_size"])
        master.minsize(500, 600)
        try:
            if ICON_FILE.exists():
                master.iconbitmap(ICON_FILE)
        except Exception:
            pass

        # --- フレームと変数 ---
        self.bg_gradient_start = "#f7f9fc"
        self.bg_gradient_end = "#d9e4f5"
        self.btn_color = "#a0c4ff"
        self.btn_hover = "#6699ff"
        self.delete_color = "#ff6b6b"
        self.delete_hover = "#ff3b3b"
        master.configure(bg=self.bg_gradient_start)
        self.main_frame = tk.Frame(master, bg=self.bg_gradient_start)
        self.main_frame.pack(fill="both", expand=True)

        # 上部設定
        self.frame_top = tk.Frame(self.main_frame, bg=self.bg_gradient_start, padx=10, pady=10)
        self.frame_top.pack(fill="x")

        self.var_manual = tk.IntVar(value=1 if self.config["manual_mode"] else 0)
        self.chk_manual = tk.Checkbutton(self.frame_top, text="履歴残さないモード",
                                         variable=self.var_manual, command=self.toggle_manual,
                                         font=("Yu Gothic UI", 12), bg=self.bg_gradient_start)
        self.chk_manual.pack(anchor="w", pady=(0,5))

        timer_frame = tk.Frame(self.frame_top, bg=self.bg_gradient_start)
        timer_frame.pack(fill="x", pady=(5,10))
        tk.Label(timer_frame, text="タイマー（分）：", font=("Yu Gothic UI",12), bg=self.bg_gradient_start).pack(side="left")
        self.var_timer = tk.StringVar(value=str(self.config["timer_minutes"]))
        vcmd = (self.master.register(self.validate_number), '%P')
        self.entry_timer = tk.Entry(timer_frame, textvariable=self.var_timer, width=5, font=("Yu Gothic UI",12), validate="key", validatecommand=vcmd)
        self.entry_timer.pack(side="left", padx=(5,10))
        self.btn_timer = tk.Button(timer_frame, text="タイマー開始", command=self.start_timer_mode,
                                   font=("Yu Gothic UI",12), bg=self.btn_color, activebackground=self.btn_hover)
        self.btn_timer.pack(side="left")

        self.var_show = tk.IntVar(value=1 if self.config["show_history"] else 0)
        self.chk_show = tk.Checkbutton(self.frame_top, text="履歴一覧を表示",
                                       variable=self.var_show, command=self.toggle_show,
                                       font=("Yu Gothic UI",12), bg=self.bg_gradient_start)
        self.chk_show.pack(anchor="w")

        detail_frame = tk.Frame(self.frame_top, bg=self.bg_gradient_start)
        detail_frame.pack(fill="x", pady=(5,10))
        self.var_detail = tk.IntVar(value=1 if self.config["show_details"] else 0)
        self.chk_detail = tk.Checkbutton(detail_frame, text="詳細表示",
                                         variable=self.var_detail, command=self.toggle_detail,
                                         font=("Yu Gothic UI",12), bg=self.bg_gradient_start)
        self.chk_detail.pack(side="left")
        self.var_auto_refresh = tk.IntVar(value=1 if self.config["auto_refresh"] else 0)
        self.chk_auto_refresh = tk.Checkbutton(detail_frame, text="自動更新",
                                               variable=self.var_auto_refresh, command=self.toggle_auto_refresh,
                                               font=("Yu Gothic UI",12), bg=self.bg_gradient_start)
        self.chk_auto_refresh.pack(side="left", padx=(10,0))

        self.lbl_status = tk.Label(self.main_frame, text="ステータス： 読み込み中...", fg="blue",
                                   font=("Yu Gothic UI",14,"bold"), bg=self.bg_gradient_start)
        self.lbl_status.pack(pady=(5,10))

        self.frame_tree = tk.Frame(self.main_frame, bg=self.bg_gradient_start)
        self.setup_treeview()
        self.scrollbar = ttk.Scrollbar(self.frame_tree, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<Double-1>", self.on_double_click)

        # 下部ボタン
        self.frame_bottom = tk.Frame(self.main_frame, bg=self.bg_gradient_start, padx=10, pady=10)
        self.frame_bottom.pack(fill="x")
        self.btn_refresh = tk.Button(self.frame_bottom, text="更新", command=self.update_tree,
                                     font=("Yu Gothic UI",12), bg=self.btn_color, activebackground=self.btn_hover)
        self.btn_refresh.pack(side="left", padx=5)
        self.btn_select_all = tk.Button(self.frame_bottom, text="全選択", command=self.select_all,
                                        font=("Yu Gothic UI",12), bg=self.btn_color, activebackground=self.btn_hover)
        self.btn_select_all.pack(side="left", padx=5)
        self.btn_unselect_all = tk.Button(self.frame_bottom, text="全解除", command=self.unselect_all,
                                          font=("Yu Gothic UI",12), bg=self.btn_color, activebackground=self.btn_hover)
        self.btn_unselect_all.pack(side="left", padx=5)
        self.btn_delete = tk.Button(self.frame_bottom, text="選択削除", command=self.delete_selected,
                                    bg=self.delete_color, fg="white", font=("Yu Gothic UI",12,"bold"), activebackground=self.delete_hover)
        self.btn_delete.pack(side="left", padx=5)
        self.btn_exit = tk.Button(self.frame_bottom, text="終了", command=self.on_exit,
                                  font=("Yu Gothic UI",12), bg=self.btn_color, activebackground=self.btn_hover)
        self.btn_exit.pack(side="right", padx=5)

        # 初期状態設定
        success, msg = suppress_history(self.var_manual.get() == 1)
        self.lbl_status.config(text=f"ステータス： {msg}")

        if self.var_show.get() == 1:
            self.frame_tree.pack(fill="both", expand=True)
            self.update_tree()
        else:
            self.frame_tree.pack_forget()

        self.refresh_job = None
        if self.var_auto_refresh.get() == 1:
            self.start_auto_refresh()

        self.master.bind("<Configure>", self.on_window_resize)

    # --- メソッド実装 ---
    def validate_number(self, P):
        if P == "":
            return True
        try:
            int(P)
            return True
        except ValueError:
            return False

    def toggle_manual(self):
        active = self.var_manual.get() == 1
        suppress_history(active)
        self.config["manual_mode"] = active
        save_config(self.config)
        msg = "非履歴モード ON" if active else "非履歴モード OFF"
        self.lbl_status.config(text=f"ステータス： {msg}")

    def start_timer_mode(self):
        try:
            minutes = int(self.var_timer.get())
            if minutes <= 0:
                raise ValueError
        except:
            messagebox.showwarning("入力エラー", "有効な数字を入力してください。")
            return
        self.config["timer_minutes"] = minutes
        save_config(self.config)
        self.btn_timer.config(state="disabled")
        start_timer(minutes, self.timer_finished)
        self.lbl_status.config(text=f"ステータス： タイマー {minutes}分開始")

    def timer_finished(self):
        self._update_after_timer()

    def _update_after_timer(self):
        self.master.after(0, lambda: [
            suppress_history(False),
            self.btn_timer.config(state="normal"),
            self.lbl_status.config(text="ステータス： タイマー終了、非履歴モード OFF")
        ])

    def toggle_show(self):
        active = self.var_show.get() == 1
        self.config["show_history"] = active
        save_config(self.config)
        if active:
            self.frame_tree.pack(fill="both", expand=True)
            self.update_tree()
        else:
            self.frame_tree.pack_forget()

    def toggle_detail(self):
        self.config["show_details"] = self.var_detail.get() == 1
        save_config(self.config)
        # Re-setup the treeview with new columns and then update
        self.setup_treeview()
        self.update_tree()


    def toggle_auto_refresh(self):
        active = self.var_auto_refresh.get() == 1
        self.config["auto_refresh"] = active
        save_config(self.config)
        if active:
            self.start_auto_refresh()
        else:
            self.stop_auto_refresh()

    def start_auto_refresh(self):
        self.stop_auto_refresh()
        self.refresh_job = self.master.after(self.config.get("refresh_interval",60)*1000, self.auto_refresh)

    def stop_auto_refresh(self):
        if self.refresh_job is not None:
            self.master.after_cancel(self.refresh_job)
            self.refresh_job = None

    def auto_refresh(self):
        self.update_tree()
        if self.var_auto_refresh.get() == 1:
            self.start_auto_refresh()

    def setup_treeview(self):
        # Destroy existing treeview if it exists
        for widget in self.frame_tree.winfo_children():
            widget.destroy()
            
        columns = ("name", "path", "modified", "lnk_path") if self.var_detail.get() else ("name", "lnk_path")
        self.tree = ttk.Treeview(self.frame_tree, columns=columns, show="headings", selectmode="extended")
        
        self.tree.heading("name", text="ファイル名")
        if self.var_detail.get():
            self.tree.heading("path", text="パス")
            self.tree.heading("modified", text="更新日時")
        
        # Hide the lnk_path column
        self.tree.column("lnk_path", width=0, stretch=tk.NO)
        self.tree.column("name", width=200)
        if self.var_detail.get():
            self.tree.column("path", width=250)
            self.tree.column("modified", width=120)
            
        self.scrollbar = ttk.Scrollbar(self.frame_tree, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<Double-1>", self.on_double_click)


    def update_tree(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        files = get_recent_files()
        for f in files:
            # Always include the lnk_path (f[4]) in the values tuple
            values = (f[0], f[1], f[3], str(f[4])) if self.var_detail.get() else (f[0], str(f[4]))
            self.tree.insert("", "end", values=values)

    def select_all(self):
        for i in self.tree.get_children():
            self.tree.selection_add(i)

    def unselect_all(self):
        for i in self.tree.selection():
            self.tree.selection_remove(i)

    def delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        confirm = messagebox.askyesno("削除確認", f"{len(sel)} 件の履歴ショートカットを削除してよろしいですか？\n（元のファイルは削除されません）")
        if not confirm:
            return
        
        deleted_count = 0
        for i in sel:
            item = self.tree.item(i)
            # The lnk_path is always the last element in the values tuple
            lnk_path_index = -1
            try:
                lnk_path = item["values"][lnk_path_index]
                if lnk_path and Path(lnk_path).exists():
                    os.remove(lnk_path)
                    deleted_count += 1
            except (IndexError, Exception) as e:
                print(f"Could not delete shortcut: {e}")
                continue
        
        if deleted_count > 0:
            messagebox.showinfo("完了", f"{deleted_count} 件の履歴を削除しました。")

        self.update_tree()


    def on_tree_click(self, event):
        pass

    def on_double_click(self, event):
        item_id = self.tree.selection()
        if not item_id:
            return
        
        item = self.tree.item(item_id[0])
        # The target path is the second element if details are shown
        target_path_index = 1 
        lnk_path_index = -1

        try:
            # First, try to open the target file directly
            if self.var_detail.get():
                target_path = item["values"][target_path_index]
                if target_path and Path(target_path).exists():
                    os.startfile(target_path)
                    return
            
            # If that fails or details are off, fall back to the lnk file
            lnk_path = item["values"][lnk_path_index]
            if lnk_path and Path(lnk_path).exists():
                os.startfile(lnk_path)

        except (IndexError, Exception) as e:
            messagebox.showwarning("オープンエラー", f"ファイルを開けませんでした。\n{e}")


    def on_window_resize(self, event):
        # Avoid saving size too frequently
        if hasattr(self, '_resize_job'):
            self.master.after_cancel(self._resize_job)
        self._resize_job = self.master.after(500, self._save_window_size)

    def _save_window_size(self):
        self.config["window_size"] = f"{self.master.winfo_width()}x{self.master.winfo_height()}"
        save_config(self.config)


    def on_exit(self):
        self.stop_auto_refresh()
        self.master.destroy()

# --- 実行 ---
def main():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    root = tk.Tk()
    style = ttk.Style()
    app = HistoryGuardApp(root)
    # A more robust global error handler
    def global_exception_handler(exc_type, exc_value, exc_traceback):
        messagebox.showerror("致命的なエラー", f"予期せぬエラーが発生しました: {exc_value}")
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
    sys.excepthook = global_exception_handler
    root.mainloop()

if __name__ == "__main__":
    main()