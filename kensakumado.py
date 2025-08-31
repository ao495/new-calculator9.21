import tkinter as tk
from tkinter import ttk, simpledialog, messagebox, font
import webbrowser
import json
import os
import threading
import pystray
from PIL import Image, ImageDraw
import keyboard  # pip install keyboard
import re

CONFIG_FILE = "search_config.json"

class TraySearchApp(tk.Tk):
    def __init__(self):
        super().__init__()

        # --- スタイルの設定 ---
        self.dark_mode = True # ダークモードの切り替え用フラグ (将来的に)
        self.colors = {
            'bg': '#2E2E2E',
            'fg': '#EAEAEA',
            'entry_bg': '#3C3C3C',
            'entry_fg': '#EAEAEA',
            'btn_bg': '#555555',
            'btn_fg': '#EAEAEA',
            'tab_bg': '#2E2E2E',
            'tab_fg': '#EAEAEA',
            'tab_selected_bg': '#4A4A4A',
            'tab_selected_fg': '#FFFFFF'
        }

        style = ttk.Style(self)
        style.theme_use('clam') # カスタマイズしやすいclamテーマをベースにする

        # --- 全体のスタイル ---
        style.configure('.', background=self.colors['bg'], foreground=self.colors['fg'], fieldbackground=self.colors['entry_bg'], lightcolor=self.colors['bg'], darkcolor=self.colors['bg'], bordercolor=self.colors['bg'])
        self.configure(background=self.colors['bg'])

        # --- 各ウィジェットのスタイル ---
        style.configure('TFrame', background=self.colors['bg'])
        style.configure('TLabel', background=self.colors['bg'], foreground=self.colors['fg'])
        style.configure('TButton', background=self.colors['btn_bg'], foreground=self.colors['btn_fg'])
        style.map('TButton', background=[('active', '#6A6A6A')])
        style.configure('TEntry', fieldbackground=self.colors['entry_bg'], foreground=self.colors['entry_fg'], insertcolor=self.colors['fg'])
        
        # --- タブのスタイル ---
        style.configure('TNotebook', background=self.colors['bg'])
        style.configure('TNotebook.Tab', background=self.colors['tab_bg'], foreground=self.colors['tab_fg'], padding=[5, 2])
        style.map('TNotebook.Tab', background=[('selected', self.colors['tab_selected_bg'])], foreground=[('selected', self.colors['tab_selected_fg'])])

        # --- Treeview (リスト) のスタイル ---
        style.configure("Treeview", background=self.colors['entry_bg'], fieldbackground=self.colors['entry_bg'], foreground=self.colors['fg'])
        style.map('Treeview', background=[('selected', self.colors['tab_selected_bg'])])
        style.configure("Treeview.Heading", background=self.colors['btn_bg'], foreground=self.colors['btn_fg'])
        style.map("Treeview.Heading", background=[('active', '#6A6A6A')])
        self.title("単語ボタン検索アプリ")
        self.geometry("700x250")
        self.protocol("WM_DELETE_WINDOW", self._hide_window)

        self.config_data = self._load_config()
        self.common_words = self.config_data.get("common_words", ["Python", "Tkinter", "AI", "GAS", "Amazon"])
        self.search_engines = self.config_data.get("search_engines", {
            "Google": "https://www.google.com/search?q={}",
            "Bing": "https://www.bing.com/search?q={}",
            "DuckDuckGo": "https://duckduckgo.com/?q={}"
        })

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
        self.search_btn = ttk.Button(search_frame, text="検索", command=self._search)
        self.search_btn.pack(side="left", padx=(5, 0))

        # --- 設定ボタン ---
        settings_btn = ttk.Button(search_frame, text="⚙", width=3, command=self._open_settings_window)
        settings_btn.pack(side="left", padx=(5, 0))

        # --- 単語ボタンフレーム ---
        self.words_frame = ttk.Frame(self)
        self.words_frame.pack(pady=5, padx=10, fill=tk.X)
        self._refresh_word_buttons()

        # --- ボタンフレーム ---
        button_frame = ttk.Frame(self)
        button_frame.pack(pady=(5, 10))

        # --- 単語追加ボタン ---
        add_word_btn = ttk.Button(button_frame, text="単語追加", command=self._add_word)
        add_word_btn.pack(side="left", padx=5)

        # --- トレイ ---
        self.icon = pystray.Icon("TraySearch", self._create_tray_image(), "検索アプリ")
        menu = pystray.Menu(
            pystray.MenuItem("表示", lambda: self._show_window()),
            pystray.MenuItem("終了", lambda: self._exit_app())
        )
        self.icon.menu = menu
        threading.Thread(target=self.icon.run, daemon=True).start()

        # --- ホットキー ---
        self.current_hotkey = self.config_data.get('hotkey', 'ctrl+alt+space')
        self._register_hotkey(self.current_hotkey)

        # 入力途中での非表示用ホットキー
        keyboard.add_hotkey("ctrl+m", self._hide_window, suppress=True)

        # 最初は非表示でトレイ常駐
        self.withdraw()

    # -----------------------------
    # 単語ボタン操作
    # -----------------------------
    def _refresh_word_buttons(self):
        for widget in self.words_frame.winfo_children():
            widget.destroy()
        for word in self.common_words:
            btn = ttk.Button(self.words_frame, text=word, command=lambda w=word: self._add_word_to_entry(w))
            btn.pack(side="left", padx=3, pady=3)
            btn.bind("<Button-3>", lambda e, w=word: self._remove_word(w))

    def _add_word_to_entry(self, word):
        current = self.main_entry.get()
        new_text = f"{current} {word}" if current else word
        self.main_entry.delete(0, tk.END)
        self.main_entry.insert(0, new_text)

    def _add_word(self):
        new_word = simpledialog.askstring("単語追加", "追加する単語を入力:")
        if new_word and new_word not in self.common_words:
            self.common_words.append(new_word)
            self._save_words()
            self._refresh_word_buttons()

    def _remove_word(self, word):
        if messagebox.askyesno("確認", f"単語 '{word}' を削除しますか？"):
            self.common_words.remove(word)
            self._save_config()
            self._refresh_word_buttons()

    def _save_words(self): # このメソッドは _add_word から呼ばれる
        self._save_config()

    def _save_config(self):
        self.config_data["common_words"] = self.common_words
        self.config_data["search_engines"] = self.search_engines
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.config_data, f, ensure_ascii=False, indent=2)

    def _refresh_search_tabs(self):
        # 既存のタブをすべて削除
        for tab in self.tab_control.tabs():
            self.tab_control.forget(tab)
        self.tabs.clear()

        # search_enginesに基づいてタブを再作成
        for name in self.search_engines.keys():
            frame = ttk.Frame(self.tab_control)
            self.tab_control.add(frame, text=name)
            self.tabs[name] = frame

    def _load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                return {}
        return {}

    # -----------------------------
    # 検索機能
    # -----------------------------
    def _search_on_enter(self, event=None):
        """Enterキーが押されたときに検索を実行する"""
        self._search()

    def _search(self):
        selected_tab = self.tab_control.tab(self.tab_control.select(), "text")
        engine_url = self.search_engines.get(selected_tab)
        query = self.main_entry.get().strip()
        if query and engine_url:
            # {} の中身に関わらず、プレースホルダーを置換する
            url = re.sub(r'\{.*?\}', query.replace(" ", "+"), engine_url)
            webbrowser.open(url)

    # -----------------------------
    # トレイ関連
    # -----------------------------
    def _create_tray_image(self):
        # 64x64の透明な画像を作成
        img = Image.new('RGBA', (64, 64), (0, 0, 0, 0)) # RGBAで透明背景
        d = ImageDraw.Draw(img)

        # レンズ (円)
        lens_color = (200, 200, 200, 255) # 明るいグレー
        lens_border_color = (100, 100, 100, 255) # 濃いグレー
        lens_bbox = [(18, 18), (46, 46)] # x1, y1, x2, y2
        d.ellipse(lens_bbox, fill=lens_color, outline=lens_border_color, width=2)

        # 柄 (線)
        handle_color = (150, 150, 150, 255) # グレー
        d.line([(40, 40), (55, 55)], fill=handle_color, width=4)

        return img

    def _hide_window(self):
        self.withdraw()

    def _show_window(self):
        self.deiconify()
        self.lift()
        self.focus_force()
        self.main_entry.focus_force() # 検索窓にフォーカスを当てる

    def _toggle_window(self):
        if self.winfo_viewable():
            self._hide_window()
        else:
            self._show_window()

    # -----------------------------
    # 設定画面
    # -----------------------------
    def _open_settings_window(self):
        if hasattr(self, 'settings_window') and self.settings_window.winfo_exists():
            self.settings_window.lift()
            return

        self.settings_window = tk.Toplevel(self)
        self.settings_window.title("設定")
        self.settings_window.geometry("600x450") # 高さを少し増やす
        self.settings_window.configure(background=self.colors['bg'])
        self.settings_window.transient(self) # メインウィンドウの上に表示
        self.settings_window.grab_set() # モーダルにする

        # --- メインフレーム ---
        main_frame = ttk.Frame(self.settings_window, padding=10)
        main_frame.pack(expand=True, fill=tk.BOTH)

        # --- 検索エンジン設定フレーム ---
        se_frame = ttk.LabelFrame(main_frame, text="検索エンジン設定", padding=10)
        se_frame.pack(expand=True, fill=tk.BOTH)

        # --- 左側 (リスト) ---
        left_frame = ttk.Frame(se_frame)
        left_frame.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=(0, 5))

        self.se_tree = ttk.Treeview(left_frame, columns=("name", "url"), show="headings")
        self.se_tree.heading("name", text="名前")
        self.se_tree.heading("url", text="URL")
        self.se_tree.column("name", width=100)
        self.se_tree.column("url", width=250)
        self.se_tree.pack(expand=True, fill=tk.BOTH)

        # --- 右側 (編集エリア) ---
        right_frame = ttk.Frame(se_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(5, 0))

        ttk.Label(right_frame, text="名前:").pack(anchor=tk.W)
        self.se_name_entry = ttk.Entry(right_frame, width=30)
        self.se_name_entry.pack()

        ttk.Label(right_frame, text="URL:").pack(anchor=tk.W, pady=(10, 0))
        ttk.Label(right_frame, text="検索キーワードを入れたい場所に {} を入力", foreground="gray").pack(anchor=tk.W)
        ttk.Label(right_frame, text="例: https://www.google.com/search?q={}", foreground="gray").pack(anchor=tk.W)
        self.se_url_entry = ttk.Entry(right_frame, width=30)
        self.se_url_entry.pack()

        button_frame = ttk.Frame(right_frame)
        button_frame.pack(pady=20)
        add_update_btn = ttk.Button(button_frame, text="追加 / 更新", command=self._add_update_search_engine)
        add_update_btn.pack(fill=tk.X)
        delete_btn = ttk.Button(button_frame, text="削除", command=self._delete_search_engine)
        delete_btn.pack(fill=tk.X, pady=5)

        self._populate_se_tree()
        self.se_tree.bind("<<TreeviewSelect>>", self._on_se_tree_select)

        # --- ホットキー設定 ---
        hotkey_frame = ttk.LabelFrame(main_frame, text="ホットキー設定", padding=10)
        hotkey_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Label(hotkey_frame, text="表示/非表示:").pack(side=tk.LEFT)

        self.hotkey_entry_var = tk.StringVar()
        self.hotkey_entry_var.set(self.config_data.get('hotkey', 'ctrl+shift+s'))
        hotkey_entry = ttk.Entry(hotkey_frame, textvariable=self.hotkey_entry_var, state="readonly", width=20)
        hotkey_entry.pack(side=tk.LEFT, padx=5)

        self.change_hotkey_btn = ttk.Button(hotkey_frame, text="変更", command=self._start_hotkey_capture)
        self.change_hotkey_btn.pack(side=tk.LEFT)

    def _start_hotkey_capture(self):
        self.change_hotkey_btn.config(text="新しいキーを入力...", state="disabled")
        threading.Thread(target=self._capture_hotkey_thread, daemon=True).start()

    def _capture_hotkey_thread(self):
        new_hotkey = keyboard.read_hotkey(suppress=False)
        self.after(0, self._set_new_hotkey, new_hotkey)

    def _set_new_hotkey(self, new_hotkey):
        # 古いホットキーをハンドルで正確に解除
        keyboard.remove_hotkey(self.hotkey_handler)

        self._register_hotkey(new_hotkey, update_config=True)

        self.change_hotkey_btn.config(text="変更", state="normal")

    def _register_hotkey(self, hotkey_string, update_config=False):
        try:
            # 新しいホットキーを登録し、新しいハンドルを保存
            self.hotkey_handler = keyboard.add_hotkey(hotkey_string, self._toggle_window, suppress=True)
            self.current_hotkey = hotkey_string
            if update_config:
                self.config_data['hotkey'] = hotkey_string
                self.hotkey_entry_var.set(hotkey_string)
                self._save_config()
                messagebox.showinfo("成功", f"ホットキーが '{hotkey_string}' に設定されました。", parent=self.settings_window)
        except (ValueError, KeyError) as e:
            messagebox.showerror("エラー", f"無効なキーの組み合わせです: {e}", parent=self.settings_window)
            # エラーが発生した場合、古いホットキーを再登録
            try:
                self.hotkey_handler = keyboard.add_hotkey(self.current_hotkey, self._toggle_window, suppress=True)
            except (ValueError, KeyError):
                pass # 再登録にも失敗した場合は何もしない

    def _populate_se_tree(self):
        # Treeviewをクリア
        for i in self.se_tree.get_children():
            self.se_tree.delete(i)
        # データを挿入
        for name, url in self.search_engines.items():
            self.se_tree.insert("", tk.END, values=(name, url))

    def _on_se_tree_select(self, event):
        selected_items = self.se_tree.selection()
        if not selected_items:
            return
        
        selected_item = selected_items[0]
        name, url = self.se_tree.item(selected_item, "values")
        self.se_name_entry.delete(0, tk.END)
        self.se_name_entry.insert(0, name)
        self.se_url_entry.delete(0, tk.END)
        self.se_url_entry.insert(0, url)

    def _add_update_search_engine(self):
        name = self.se_name_entry.get().strip()
        url = self.se_url_entry.get().strip()

        if not name or not url:
            messagebox.showerror("エラー", "名前とURLを入力してください。", parent=self.settings_window)
            return
        if not re.search(r'\{.*?\}', url):
            messagebox.showerror("エラー", "URLには検索語を表すプレースホルダー（例: {}）を含めてください。", parent=self.settings_window)
            return

        self.search_engines[name] = url
        self._save_config() # 設定を保存
        self._populate_se_tree() # Treeviewを更新
        self._refresh_search_tabs() # メインウィンドウのタブを更新

        # 入力欄をクリア
        self.se_name_entry.delete(0, tk.END)
        self.se_url_entry.delete(0, tk.END)

    def _delete_search_engine(self):
        selected_items = self.se_tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "削除する検索エンジンを選択してください。", parent=self.settings_window)
            return

        if not messagebox.askyesno("確認", "選択した検索エンジンを削除しますか？", parent=self.settings_window):
            return

        for selected_item in selected_items:
            name, _ = self.se_tree.item(selected_item, "values")
            if name in self.search_engines:
                del self.search_engines[name]

        self._save_config() # 設定を保存
        self._populate_se_tree() # Treeviewを更新
        self._refresh_search_tabs() # メインウィンドウのタブを更新

        # 入力欄をクリア
        self.se_name_entry.delete(0, tk.END)
        self.se_url_entry.delete(0, tk.END)

    def _exit_app(self):
        self.icon.stop()
        keyboard.unhook_all_hotkeys()
        self.destroy()

if __name__ == "__main__":
    app = TraySearchApp()
    app.mainloop()
