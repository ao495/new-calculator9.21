import tkinter as tk
from tkinter import ttk, simpledialog, messagebox, font
import webbrowser
import json
import os
import threading
import pystray
from PIL import Image, ImageDraw
import re
from pynput import keyboard as pk
import ctypes # Add this line

# --- 追加: ダークテーマ対応のカスタムダイアログ ---
class CustomAskStringDialog(tk.Toplevel):
    def __init__(self, parent, title=None, prompt=None, initialvalue=None, colors=None):
        super().__init__(parent)
        self.transient(parent)
        self.title(title)
        self.parent = parent
        self.result = None
        self.colors = colors or {'bg': '#2E2E2E', 'fg': '#EAEAEA', 'entry_bg': '#3C3C3C', 'btn_bg': '#555555'}

        self.configure(background=self.colors['bg'])

        # このダイアログ用のスタイルを定義
        style = ttk.Style(self)
        style.configure('DarkDialog.TLabel', background=self.colors['bg'], foreground=self.colors['fg'])
        style.configure('DarkDialog.TButton', background=self.colors['btn_bg'], foreground=self.colors['fg'])
        style.map('DarkDialog.TButton', background=[('active', '#6A6A6A')])
        style.configure('DarkDialog.TEntry', fieldbackground=self.colors['entry_bg'], foreground=self.colors['entry_fg'], insertcolor=self.colors['fg'])
        style.configure('DarkDialog.TFrame', background=self.colors['bg'])


        ttk.Label(self, text=prompt, style='DarkDialog.TLabel').pack(padx=10, pady=(10, 5))
        
        self.entry = ttk.Entry(self, width=40, style='DarkDialog.TEntry')
        self.entry.pack(padx=10, pady=(0, 10))
        if initialvalue:
            self.entry.insert(0, initialvalue)
        self.entry.focus_set()

        button_frame = ttk.Frame(self, style='DarkDialog.TFrame')
        button_frame.pack(pady=5)

        ttk.Button(button_frame, text="OK", command=self._on_ok, style='DarkDialog.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=self._on_cancel, style='DarkDialog.TButton').pack(side=tk.LEFT, padx=5)

        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.bind("<Return>", self._on_ok)
        self.bind("<Escape>", self._on_cancel)

        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (self.winfo_width() // 2)
        y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")
        
        self.grab_set()
        self.wait_window(self)

    def _on_ok(self, event=None):
        self.result = self.entry.get()
        self.destroy()

    def _on_cancel(self, event=None):
        self.result = None
        self.destroy()

def ask_string_dark(parent, title, prompt, initialvalue=None, colors=None):
    """カスタムのダークテーマ対応askstringダイアログ"""
    dialog = CustomAskStringDialog(parent, title, prompt, initialvalue, colors)
    return dialog.result

CONFIG_FILE = "search_config.json"

class TraySearchApp(tk.Tk):
    def __init__(self):
        super().__init__()

        # --- カラー設定 ---
        self.colors = {
            'bg': '#2E2E2E', 'fg': '#EAEAEA',
            'entry_bg': '#3C3C3C', 'entry_fg': '#EAEAEA',
            'btn_bg': '#555555', 'btn_fg': '#EAEAEA',
            'tab_bg': '#2E2E2E', 'tab_fg': '#EAEAEA',
            'tab_selected_bg': '#4A4A4A', 'tab_selected_fg': '#FFFFFF'
        }

        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('.', background=self.colors['bg'], foreground=self.colors['fg'], fieldbackground=self.colors['entry_bg'])
        style.configure('TFrame', background=self.colors['bg'])
        style.configure('TLabel', background=self.colors['bg'], foreground=self.colors['fg'])
        style.configure('TButton', background=self.colors['btn_bg'], foreground=self.colors['btn_fg'])
        style.map('TButton', background=[('active', '#6A6A6A')])
        style.configure('TEntry', fieldbackground=self.colors['entry_bg'], foreground=self.colors['entry_fg'], insertcolor=self.colors['fg'])
        style.configure('TNotebook', background=self.colors['bg'])
        style.configure('TNotebook.Tab', background=self.colors['tab_bg'], foreground=self.colors['tab_fg'], padding=[5,2])
        style.map('TNotebook.Tab', background=[('selected', self.colors['tab_selected_bg'])], foreground=[('selected', self.colors['tab_selected_fg'])])
        style.configure("Treeview", background=self.colors['entry_bg'], fieldbackground=self.colors['entry_bg'], foreground=self.colors['fg'])
        style.map('Treeview', background=[('selected', self.colors['tab_selected_bg'])])
        style.configure("Treeview.Heading", background=self.colors['btn_bg'], foreground=self.colors['btn_fg'])
        style.map("Treeview.Heading", background=[('active', '#6A6A6A')])
        
        self.configure(background=self.colors['bg']) # メインウィンドウの背景色

        self.title("単語ボタン検索アプリ")
        self.geometry("700x280") # 高さを少し広げる
        self.protocol("WM_DELETE_WINDOW", self._hide_window)

        # --- 設定ロード ---
        self.config_data = self._load_config()
        self.common_words = self.config_data.get("common_words", ["Python","Tkinter","AI","GAS","Amazon"])
        self.search_engines = self.config_data.get("search_engines", {
            "Google":"https://www.google.com/search?q={}",
            "Bing":"https://www.bing.com/search?q={}",
            "DuckDuckGo":"https://duckduckgo.com/?q={}"
        })
        self.hotkey_combination = set(self.config_data.get('hotkey', ['ctrl_l','alt_l','space']))
        self.focus_delay = self.config_data.get('focus_delay', 500) # Default to 500ms

        # --- タブ ---
        self.tab_control = ttk.Notebook(self)
        self.tab_control.pack(fill="x", padx=10, pady=5)
        self.tabs = {}
        self._refresh_search_tabs()

        # --- 検索窓 ---
        search_frame = ttk.Frame(self)
        search_frame.pack(fill=tk.X, padx=10, pady=5)
        entry_font = font.Font(family="Yu Gothic UI", size=14)
        self.main_entry = ttk.Entry(search_frame, width=60, font=entry_font)
        self.main_entry.pack(side="left", expand=True, fill="x")
        self.main_entry.bind("<Return>", self._search_on_enter)
        ttk.Button(search_frame, text="検索", command=self._search).pack(side="left", padx=(5,0))
        ttk.Button(search_frame, text="⚙", width=3, command=self._open_settings_window).pack(side="left", padx=(5,0))

        # --- 単語ボタンフレーム ---
        self.words_frame = ttk.Frame(self)
        self.words_frame.pack(pady=5, padx=10, fill=tk.X)
        self._refresh_word_buttons()
        button_frame = ttk.Frame(self)
        button_frame.pack(pady=(5,10))
        ttk.Button(button_frame,text="単語追加",command=self._add_word).pack(side="left", padx=5)

        # --- トレイ ---
        self.icon = pystray.Icon("TraySearch", self._create_tray_image(), "検索アプリ")
        menu = pystray.Menu(
            pystray.MenuItem("表示", lambda:self._show_window()),
            pystray.MenuItem("終了", lambda:self._exit_app())
        )
        self.icon.menu = menu
        threading.Thread(target=self.icon.run, daemon=True).start()

        # --- IME対応ホットキー ---
        self.pressed_keys = set()
        self.is_setting_hotkey = False
        self.listener = pk.Listener(on_press=self._on_press, on_release=self._on_release)
        self.listener.start()

        self.withdraw()

    # -----------------------------
    # キーイベント処理
    # -----------------------------
    def _get_key_name(self, key):
        """pynputのキーオブジェクトを正規化された文字列に変換する"""
        if isinstance(key, pk.Key):
            return key.name
        elif hasattr(key, 'char') and key.char:
            if key.char == ' ': return 'space'
            return key.char.lower()
        return None

    def _on_press(self, key):
        if self.is_setting_hotkey: return # ホットキー設定中はメインの動作を止める
        key_name = self._get_key_name(key)
        if key_name:
            self.pressed_keys.add(key_name)
        if self.hotkey_combination.issubset(self.pressed_keys):
            self._toggle_window()

    def _on_release(self, key):
        if self.is_setting_hotkey: return
        key_name = self._get_key_name(key)
        if key_name:
            self.pressed_keys.discard(key_name)

    # -----------------------------
    # 単語ボタン操作
    # -----------------------------
    def _refresh_word_buttons(self):
        for w in self.words_frame.winfo_children(): w.destroy()
        for word in self.common_words:
            btn = ttk.Button(self.words_frame,text=word,command=lambda w=word:self._add_word_to_entry(w))
            btn.pack(side="left", padx=3, pady=3)
            btn.bind("<Button-3>", lambda e,w=word:self._remove_word(w))

    def _add_word_to_entry(self, word):
        current = self.main_entry.get()
        new_text = f"{current} {word}" if current else word
        self.main_entry.delete(0, tk.END)
        self.main_entry.insert(0,new_text+" ")

    def _add_word(self):
        w = ask_string_dark(self, "単語追加", "追加する単語を入力:", colors=self.colors)
        if w and w not in self.common_words:
            self.common_words.append(w)
            self._save_config()
            self._refresh_word_buttons()

    def _remove_word(self, word):
        if messagebox.askyesno("確認",f"単語 '{word}' を削除しますか？"):
            self.common_words.remove(word)
            self._save_config()
            self._refresh_word_buttons()

    # -----------------------------
    # 検索タブ
    # -----------------------------
    def _refresh_search_tabs(self):
        for t in self.tab_control.tabs(): self.tab_control.forget(t)
        self.tabs.clear()
        for name in self.search_engines.keys():
            frame = ttk.Frame(self.tab_control)
            self.tab_control.add(frame,text=name)
            self.tabs[name]=frame

    def _search_on_enter(self,event=None): self._search()
    def _search(self):
        selected_tab = self.tab_control.tab(self.tab_control.select(),"text")
        engine_url = self.search_engines.get(selected_tab)
        query = self.main_entry.get().strip()
        if query and engine_url:
            url = re.sub(r'\{.*?\}', query.replace(" ","+"), engine_url)
            webbrowser.open(url)
            # ブラウザが開いた後にメインウィンドウとエントリーにフォーカスを戻す
            self.after(self.focus_delay, self._show_window) # _show_windowがフォーカス処理を含む

    # -----------------------------
    # トレイ
    # -----------------------------
    def _create_tray_image(self):
        img = Image.new('RGBA',(64,64),(0,0,0,0))
        d = ImageDraw.Draw(img)
        d.ellipse([(18,18),(46,46)],fill=(200,200,200,255),outline=(100,100,100,255),width=2)
        d.line([(40,40),(55,55)],fill=(150,150,150,255),width=4)
        return img

    # -----------------------------
    # ウィンドウ表示/非表示
    # -----------------------------
    def _hide_window(self): self.withdraw()
    def _show_window(self):
        self.deiconify()
        self.state("normal")
        self.lift()
        self.update_idletasks()
        # Delay the focus calls slightly to ensure window is fully ready
        self.after(50, lambda: self._apply_focus())

    def _apply_focus(self):
        self.focus_force()
        self.main_entry.focus_force()
        self.wm_attributes('-topmost',1)
        self.after(100, lambda:self.wm_attributes('-topmost',0))

        # Attempt to force foreground using Win32 API for Windows
        if os.name == 'nt': # Check if OS is Windows
            try:
                # Define Win32 API functions
                SetForegroundWindow = ctypes.windll.user32.SetForegroundWindow
                GetForegroundWindow = ctypes.windll.user32.GetForegroundWindow
                AttachThreadInput = ctypes.windll.user32.AttachThreadInput
                GetCurrentThreadId = ctypes.windll.kernel32.GetCurrentThreadId
                
                # Get the window handle (HWND) of the Tkinter window
                hwnd = self.winfo_id()
                
                # Get current foreground window's thread ID
                foreground_hwnd = GetForegroundWindow()
                if foreground_hwnd:
                    foreground_thread_id = ctypes.windll.user32.GetWindowThreadProcessId(foreground_hwnd, 0)
                else:
                    foreground_thread_id = 0
                
                # Get current application's thread ID
                current_thread_id = GetCurrentThreadId()

                # Attach threads if they are different
                if foreground_thread_id != current_thread_id:
                    AttachThreadInput(current_thread_id, foreground_thread_id, True)
                
                # Bring the window to the foreground
                SetForegroundWindow(hwnd)
                
                # Detach threads
                if foreground_thread_id != current_thread_id:
                    AttachThreadInput(current_thread_id, foreground_thread_id, False)
            except Exception as e:
                print(f"Error setting foreground window: {e}")
    def _toggle_window(self):
        self._hide_window() if self.winfo_viewable() else self._show_window()

    # -----------------------------
    # 設定画面
    # -----------------------------
    def _open_settings_window(self):
        if hasattr(self,'settings_window') and self.settings_window.winfo_exists():
            self.settings_window.lift(); return
        self.settings_window = tk.Toplevel(self)
        self.settings_window.title("設定")
        self.settings_window.geometry("600x500")
        self.settings_window.configure(background=self.colors['bg'])
        self.settings_window.transient(self)
        self.settings_window.grab_set()
        main_frame = ttk.Frame(self.settings_window,padding=10); main_frame.pack(expand=True,fill=tk.BOTH)

        # --- 検索エンジン設定 ---
        se_frame = ttk.LabelFrame(main_frame,text="検索エンジン設定",padding=10); se_frame.pack(expand=True,fill=tk.BOTH, pady=(0, 5))
        left_frame = ttk.Frame(se_frame); left_frame.pack(side=tk.LEFT,expand=True,fill=tk.BOTH,padx=(0,5))
        self.se_tree = ttk.Treeview(left_frame,columns=("name","url"),show="headings",height=6)
        self.se_tree.heading("name",text="名前"); self.se_tree.heading("url",text="URL")
        self.se_tree.pack(expand=True,fill=tk.BOTH)
        for n,u in self.search_engines.items(): self.se_tree.insert("",tk.END,values=(n,u))
        right_frame = ttk.Frame(se_frame); right_frame.pack(side=tk.LEFT,fill=tk.Y)
        ttk.Button(right_frame,text="追加",command=self._add_search_engine).pack(pady=2)
        ttk.Button(right_frame,text="編集",command=self._edit_search_engine).pack(pady=2)
        ttk.Button(right_frame,text="削除",command=self._delete_search_engine).pack(pady=2)

        # --- ホットキー設定 ---
        hk_frame = ttk.LabelFrame(main_frame, text="ホットキー設定", padding=10)
        hk_frame.pack(fill=tk.X, pady=5)
        
        self.hotkey_label_var = tk.StringVar(value=f"現在のホットキー: {' + '.join(sorted(list(self.hotkey_combination)))}")
        ttk.Label(hk_frame, textvariable=self.hotkey_label_var).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        ttk.Button(hk_frame, text="ホットキーを設定", command=self._set_hotkey).pack(side=tk.LEFT, padx=5)

        # --- フォーカス遅延設定 ---
        fd_frame = ttk.LabelFrame(main_frame, text="フォーカス遅延設定 (ミリ秒)", padding=10)
        fd_frame.pack(fill=tk.X, pady=5)

        self.focus_delay_var = tk.StringVar(value=str(self.focus_delay))
        ttk.Label(fd_frame, text="遅延時間:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(fd_frame, textvariable=self.focus_delay_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(fd_frame, text="設定", command=self._save_focus_delay).pack(side=tk.LEFT, padx=5)

    def _add_search_engine(self):
        n = ask_string_dark(self.settings_window, "追加", "検索エンジン名:", colors=self.colors)
        if not n: return
        u = ask_string_dark(self.settings_window, "追加", "URL (例: https://www.google.com/search?q={}):", colors=self.colors)
        if not u: return
        self.search_engines[n] = u
        self._save_config()
        self._refresh_search_tabs()
        self.se_tree.insert("",tk.END,values=(n,u))

    def _edit_search_engine(self):
        sel = self.se_tree.selection()
        if not sel: return
        item = self.se_tree.item(sel[0])
        old_name,old_url = item['values']
        n = ask_string_dark(self.settings_window, "編集", "検索エンジン名:", initialvalue=old_name, colors=self.colors)
        if not n: return
        u = ask_string_dark(self.settings_window, "編集", "URL:", initialvalue=old_url, colors=self.colors)
        if not u: return
        del self.search_engines[old_name]
        self.search_engines[n] = u
        self._save_config()
        self._refresh_search_tabs()
        self.se_tree.item(sel[0],values=(n,u))

    def _delete_search_engine(self):
        sel = self.se_tree.selection()
        if not sel: return
        item = self.se_tree.item(sel[0])
        name = item['values'][0]
        if messagebox.askyesno("確認",f"{name} を削除しますか？"):
            del self.search_engines[name]
            self._save_config()
            self._refresh_search_tabs()
            self.se_tree.delete(sel[0])

    def _set_hotkey(self):
        if self.is_setting_hotkey: return
        self.is_setting_hotkey = True

        setter_window = tk.Toplevel(self.settings_window)
        setter_window.title("ホットキー設定")
        setter_window.transient(self.settings_window)
        setter_window.grab_set()
        setter_window.configure(background=self.colors['bg'])

        new_hotkeys = set()
        label_var = tk.StringVar(value="新しいホットキーを押してください...")
        info_label = ttk.Label(setter_window, textvariable=label_var, font=("Yu Gothic UI", 10))
        info_label.pack(pady=20, padx=20)

        setter_window.update_idletasks()
        x = self.settings_window.winfo_rootx() + (self.settings_window.winfo_width() // 2) - (setter_window.winfo_width() // 2)
        y = self.settings_window.winfo_rooty() + (self.settings_window.winfo_height() // 2) - (setter_window.winfo_height() // 2)
        setter_window.geometry(f"350x100+{x}+{y}")

        hotkey_listener = None

        def on_press(key):
            key_name = self._get_key_name(key)
            if key_name:
                new_hotkeys.add(key_name)
                label_var.set(' + '.join(sorted(list(new_hotkeys))))

        def on_release(key):
            nonlocal hotkey_listener
            non_modifier_keys = [k for k in new_hotkeys if not any(mod in k for mod in ['ctrl', 'alt', 'shift', 'cmd', 'win'])]
            if new_hotkeys and (non_modifier_keys or any(k in new_hotkeys for k in ['space', 'enter', 'tab'])):
                self.hotkey_combination = new_hotkeys
                self._save_config()
                self.hotkey_label_var.set(f"現在のホットキー: {' + '.join(sorted(list(self.hotkey_combination)))}")
                if hotkey_listener: hotkey_listener.stop()
                setter_window.destroy()
            elif isinstance(key, pk.Key) and key == pk.Key.esc:
                 if hotkey_listener: hotkey_listener.stop()
                 setter_window.destroy()

        def on_close_setter():
            nonlocal hotkey_listener
            if hotkey_listener: hotkey_listener.stop()
            self.is_setting_hotkey = False
            setter_window.destroy()

        setter_window.protocol("WM_DELETE_WINDOW", on_close_setter)
        
        hotkey_listener = pk.Listener(on_press=on_press, on_release=on_release)
        hotkey_listener.start()
        
        setter_window.wait_window()
        self.is_setting_hotkey = False

    def _save_focus_delay(self):
        try:
            new_delay = int(self.focus_delay_var.get())
            if new_delay >= 0:
                self.focus_delay = new_delay
                self._save_config()
                messagebox.showinfo("設定", "フォーカス遅延時間を保存しました。")
            else:
                messagebox.showerror("エラー", "遅延時間は0以上の整数で入力してください。")
        except ValueError:
            messagebox.showerror("エラー", "無効な入力です。整数を入力してください。")

    # -----------------------------
    # 設定保存
    # -----------------------------
    def _load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE,"r",encoding="utf-8") as f: return json.load(f)
            except: return {}
        return {}

    def _save_config(self):
        self.config_data["common_words"] = self.common_words
        self.config_data["search_engines"] = self.search_engines
        self.config_data["hotkey"] = list(self.hotkey_combination)
        self.config_data["focus_delay"] = self.focus_delay
        with open(CONFIG_FILE,"w",encoding="utf-8") as f: json.dump(self.config_data,f,ensure_ascii=False,indent=2)

    # -----------------------------
    # 終了
    # -----------------------------
    def _exit_app(self):
        self.icon.stop()
        if self.listener.is_alive():
            self.listener.stop()
        self.destroy()

if __name__ == "__main__":
    app = TraySearchApp()
    app.mainloop()