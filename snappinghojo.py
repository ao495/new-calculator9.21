import os

import sys

import time

import subprocess

from datetime import datetime

from pathlib import Path

import tkinter as tk

from tkinter import ttk, filedialog, messagebox, Scale

import threading

import keyboard

import pyperclip

from PIL import ImageGrab, Image

import win32clipboard

import win32con

from io import BytesIO

import pystray

from pystray import MenuItem as item

from PIL import Image as PilImage

import winsound

import json



class SnippingToolWrapper:

    def __init__(self):

        self.root = tk.Tk()

        self.root.title("スクリーンショット補助ツール")

        self.root.geometry("350x550")

        self.root.resizable(False, False)

        self.root.iconbitmap(self.resource_path("screenshot_icon.ico"))

        

        # 設定のデフォルト値

        self.default_config = {

            "save_folder": str(Path.home() / "Pictures" / "Screenshots"),

            "hotkey_fullscreen": "ctrl+shift+1",

            "hotkey_region": "ctrl+shift+2",

            "auto_save": True,

            "copy_to_clipboard": True,

            "play_sound": True,

            "fullscreen_delay": 0,

            "region_wait_time": 5,

            "sound_file": "default"

        }

        

        # 設定読み込み

        self.config_file = Path.home() / "AppData" / "Local" / "ScreenshotTool" / "config.json"

        self.load_config()

        

        # 設定をローカル変数に

        self.save_folder = Path(self.config["save_folder"])

        self.hotkey_fullscreen = self.config["hotkey_fullscreen"]

        self.hotkey_region = self.config["hotkey_region"]

        self.auto_save = tk.BooleanVar(value=self.config["auto_save"])

        self.copy_to_clipboard = tk.BooleanVar(value=self.config["copy_to_clipboard"])

        self.play_sound = tk.BooleanVar(value=self.config["play_sound"])

        self.fullscreen_delay = tk.IntVar(value=self.config["fullscreen_delay"])

        self.region_wait_time = tk.IntVar(value=self.config["region_wait_time"])

        self.sound_file = self.config["sound_file"]

        

        # 長押し検出用

        self.key_press_time = None

        self.long_press_threshold = 0.5  # 0.5秒以上で長押し

        

        # タスクトレイアイコン

        self.icon = None

        

        # スクリーンショット監視スレッド

        self.clipboard_check_running = False

        

        self.setup_ui()

        self.setup_hotkeys()

        

    def resource_path(self, relative_path):

        """リソースファイルへのパスを取得（exe化時に必要）"""

        try:

            base_path = sys._MEIPASS

        except AttributeError:

            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

        

    def load_config(self):

        """設定ファイルを読み込む"""

        self.config = self.default_config.copy()

        

        if self.config_file.exists():

            try:

                with open(self.config_file, 'r', encoding='utf-8') as f:

                    loaded_config = json.load(f)

                    self.config.update(loaded_config)

            except Exception as e:

                print(f"設定ファイル読み込みエラー: {e}")

                

        # 設定ディレクトリがなければ作成

        self.config_file.parent.mkdir(parents=True, exist_ok=True)

        

    def save_config(self):

        """設定ファイルを保存する"""

        # 現在の設定を取得

        self.config.update({

            "save_folder": str(self.save_folder),

            "hotkey_fullscreen": self.hotkey_fullscreen,

            "hotkey_region": self.hotkey_region,

            "auto_save": self.auto_save.get(),

            "copy_to_clipboard": self.copy_to_clipboard.get(),

            "play_sound": self.play_sound.get(),

            "fullscreen_delay": self.fullscreen_delay.get(),

            "region_wait_time": self.region_wait_time.get(),

            "sound_file": self.sound_file

        })

        

        try:

            with open(self.config_file, 'w', encoding='utf-8') as f:

                json.dump(self.config, f, ensure_ascii=False, indent=2)

        except Exception as e:

            print(f"設定ファイル保存エラー: {e}")

            messagebox.showerror("エラー", f"設定ファイルの保存に失敗しました: {e}")

            

    def setup_ui(self):

        """UIのセットアップ"""

        # スタイル設定

        style = ttk.Style()

        style.configure("TButton", padding=5)

        style.configure("TCheckbutton", padding=2)

        

        # タブコントロール

        tab_control = ttk.Notebook(self.root)

        

        # メインタブ

        main_tab = ttk.Frame(tab_control, padding="10")

        settings_tab = ttk.Frame(tab_control, padding="10")

        

        tab_control.add(main_tab, text="メイン")

        tab_control.add(settings_tab, text="設定")

        tab_control.pack(expand=1, fill="both")

        

        # ========== メインタブ ==========

        # タイトル

        ttk.Label(main_tab, text="スクリーンショット補助ツール", 

                    font=("", 12, "bold")).grid(row=0, column=0, columnspan=2, pady=10)

        

        # キャプチャボタン

        button_frame = ttk.LabelFrame(main_tab, text="キャプチャ", padding="10")

        button_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        

        ttk.Button(button_frame, text="全画面キャプチャ", 

                    command=self.capture_fullscreen, width=20).grid(row=0, column=0, pady=5)

        ttk.Button(button_frame, text="範囲選択キャプチャ", 

                    command=self.capture_region, width=20).grid(row=1, column=0, pady=5)

        

        # クイック設定

        quick_settings = ttk.LabelFrame(main_tab, text="クイック設定", padding="10")

        quick_settings.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        

        # 保存先フォルダ

        ttk.Label(quick_settings, text="保存先:").grid(row=0, column=0, sticky=tk.W)

        self.folder_label = ttk.Label(quick_settings, text=str(self.save_folder), 

                                        relief="sunken", width=25)

        self.folder_label.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5), pady=5)

        ttk.Button(quick_settings, text="変更", 

                    command=self.change_folder).grid(row=0, column=2)

        

        # オプション

        ttk.Checkbutton(quick_settings, text="自動保存", 

                        variable=self.auto_save).grid(row=1, column=0, sticky=tk.W, pady=5)

        ttk.Checkbutton(quick_settings, text="クリップボードにコピー", 

                        variable=self.copy_to_clipboard).grid(row=1, column=1, sticky=tk.W, pady=5)

        ttk.Checkbutton(quick_settings, text="効果音", 

                        variable=self.play_sound).grid(row=2, column=0, sticky=tk.W, pady=5)

        

        # ホットキー情報

        hotkey_frame = ttk.LabelFrame(main_tab, text="ホットキー", padding="10")

        hotkey_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        

        ttk.Label(hotkey_frame, text=f"全画面: {self.hotkey_fullscreen}").grid(row=0, column=0, sticky=tk.W)

        ttk.Label(hotkey_frame, text=f"範囲選択: {self.hotkey_region}").grid(row=1, column=0, sticky=tk.W)

        ttk.Label(hotkey_frame, text="PrintScreenキー: 単押し→全画面 / 長押し→範囲選択", 

                    foreground="blue").grid(row=2, column=0, sticky=tk.W, pady=5)

        

        # ステータスバー

        self.status_var = tk.StringVar(value="準備完了")

        status_bar = ttk.Label(main_tab, textvariable=self.status_var, 

                                relief="sunken", anchor="w")

        status_bar.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))

        

        # ========== 設定タブ ==========

        # 遅延時間設定

        delay_frame = ttk.LabelFrame(settings_tab, text="タイミング設定", padding="10")

        delay_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=10, padx=5)

        

        ttk.Label(delay_frame, text="全画面キャプチャ前の遅延 (秒):").grid(row=0, column=0, sticky=tk.W, pady=5)

        ttk.Scale(delay_frame, from_=0, to=10, variable=self.fullscreen_delay, 

                    orient="horizontal", length=200).grid(row=0, column=1, padx=5)

        ttk.Label(delay_frame, textvariable=self.fullscreen_delay).grid(row=0, column=2, padx=5)

        

        ttk.Label(delay_frame, text="範囲選択の最大待機時間 (秒):").grid(row=1, column=0, sticky=tk.W, pady=5)

        ttk.Scale(delay_frame, from_=1, to=30, variable=self.region_wait_time, 

                    orient="horizontal", length=200).grid(row=1, column=1, padx=5)

        ttk.Label(delay_frame, textvariable=self.region_wait_time).grid(row=1, column=2, padx=5)

        

        # 効果音設定

        sound_frame = ttk.LabelFrame(settings_tab, text="効果音設定", padding="10")

        sound_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=10, padx=5)

        

        ttk.Checkbutton(sound_frame, text="キャプチャ時に効果音を鳴らす", 

                        variable=self.play_sound).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=5)

        

        ttk.Label(sound_frame, text="効果音:").grid(row=1, column=0, sticky=tk.W, pady=5)

        self.sound_var = tk.StringVar(value=self.sound_file)

        sound_combo = ttk.Combobox(sound_frame, textvariable=self.sound_var, width=20)

        sound_combo['values'] = ('default', 'ding', 'chime', 'notify', 'recycle')

        sound_combo.grid(row=1, column=1, padx=5, pady=5)

        ttk.Button(sound_frame, text="テスト", 

                    command=self.test_sound).grid(row=1, column=2, padx=5)

        ttk.Button(sound_frame, text="カスタム音選択", 

                    command=self.select_custom_sound).grid(row=2, column=0, columnspan=3, pady=5)

        

        # ホットキー設定

        hotkey_settings = ttk.LabelFrame(settings_tab, text="ホットキー設定", padding="10")

        hotkey_settings.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=10, padx=5)

        

        ttk.Label(hotkey_settings, text="全画面キャプチャ:").grid(row=0, column=0, sticky=tk.W, pady=5)

        self.fullscreen_hotkey_var = tk.StringVar(value=self.hotkey_fullscreen)

        ttk.Entry(hotkey_settings, textvariable=self.fullscreen_hotkey_var, width=15).grid(row=0, column=1, padx=5)

        ttk.Button(hotkey_settings, text="設定", 

                    command=lambda: self.set_hotkey("fullscreen")).grid(row=0, column=2, padx=5)

        

        ttk.Label(hotkey_settings, text="範囲選択キャプチャ:").grid(row=1, column=0, sticky=tk.W, pady=5)

        self.region_hotkey_var = tk.StringVar(value=self.hotkey_region)

        ttk.Entry(hotkey_settings, textvariable=self.region_hotkey_var, width=15).grid(row=1, column=1, padx=5)

        ttk.Button(hotkey_settings, text="設定", 

                    command=lambda: self.set_hotkey("region")).grid(row=1, column=2, padx=5)

        

                # 動作設定

        action_frame = ttk.LabelFrame(settings_tab, text="動作設定", padding="10")

        action_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=10, padx=5)

        

        ttk.Checkbutton(action_frame, text="起動時にタスクトレイに最小化", 

                        variable=tk.BooleanVar(value=True)).grid(row=0, column=0, sticky=tk.W, pady=5)

        ttk.Checkbutton(action_frame, text="スクリーンショット後に通知を表示", 

                        variable=tk.BooleanVar(value=True)).grid(row=1, column=0, sticky=tk.W, pady=5)

        

        # 保存ボタン

        ttk.Button(settings_tab, text="設定を保存", 

                    command=self.save_settings).grid(row=4, column=0, pady=10)

    

    def save_settings(self):

        """設定を保存"""

        # 設定を更新

        self.sound_file = self.sound_var.get()

        

        # ホットキーの更新

        if self.hotkey_fullscreen != self.fullscreen_hotkey_var.get():

            keyboard.remove_hotkey(self.hotkey_fullscreen)

            self.hotkey_fullscreen = self.fullscreen_hotkey_var.get()

            keyboard.add_hotkey(self.hotkey_fullscreen, self.capture_fullscreen)

            

        if self.hotkey_region != self.region_hotkey_var.get():

            keyboard.remove_hotkey(self.hotkey_region)

            self.hotkey_region = self.region_hotkey_var.get()

            keyboard.add_hotkey(self.hotkey_region, self.capture_region)

        

        # 設定ファイルに保存

        self.save_config()

        self.update_status("設定を保存しました")

        messagebox.showinfo("設定保存", "設定を保存しました")

    

    def setup_hotkeys(self):

        """ホットキーの設定"""

        # 通常のホットキー

        keyboard.add_hotkey(self.hotkey_fullscreen, self.capture_fullscreen)

        keyboard.add_hotkey(self.hotkey_region, self.capture_region)

        

        # PrintScreenキーの特殊処理

        keyboard.on_press_key('print_screen', self.on_printscreen_press)

        keyboard.on_release_key('print_screen', self.on_printscreen_release)

    

    def set_hotkey(self, hotkey_type):

        """ホットキー設定モード"""

        dialog = tk.Toplevel(self.root)

        dialog.title("ホットキー設定")

        dialog.geometry("300x150")

        dialog.resizable(False, False)

        dialog.grab_set()

        

        ttk.Label(dialog, text="新しいホットキーを押してください...", 

                    font=("", 10)).pack(pady=20)

        

        key_var = tk.StringVar(value="")

        key_label = ttk.Label(dialog, textvariable=key_var, font=("", 12, "bold"))

        key_label.pack(pady=10)

        

        # キーを押したときの処理

        def on_key(e):

            keys = []

            if e.keycode == 16:  # Shift

                keys.append("shift")

            elif e.keycode == 17:  # Ctrl

                keys.append("ctrl")

            elif e.keycode == 18:  # Alt

                keys.append("alt")

            else:

                try:

                    key_name = e.char if e.char else e.keysym

                    keys.append(key_name)

                except:

                    keys.append(e.keysym)

            

            # 現在のキー組み合わせを表示

            current_keys = key_var.get().split("+") if key_var.get() else []

            for k in keys:

                if k and k not in current_keys:

                    current_keys.append(k)

            

            key_var.set("+".join(current_keys))

            

        # 確定ボタン

        def confirm():

            new_hotkey = key_var.get()

            if new_hotkey:

                if hotkey_type == "fullscreen":

                    self.fullscreen_hotkey_var.set(new_hotkey)

                else:

                    self.region_hotkey_var.set(new_hotkey)

            dialog.destroy()

            

        dialog.bind("<Key>", on_key)

        ttk.Button(dialog, text="確定", command=confirm).pack(pady=10)

    

    def on_printscreen_press(self, event):

        """PrintScreenキーが押された時"""

        self.key_press_time = time.time()

        

    def on_printscreen_release(self, event):

        """PrintScreenキーが離された時"""

        if self.key_press_time:

            press_duration = time.time() - self.key_press_time

            self.key_press_time = None

            

            if press_duration >= self.long_press_threshold:

                # 長押し → 範囲選択

                self.capture_region()

            else:

                # 単押し → 全画面

                self.capture_fullscreen()

    

    def capture_fullscreen(self):

        """全画面キャプチャ"""

        self.update_status("全画面キャプチャ準備中...")

        

        # 遅延がある場合は待機

        delay = self.fullscreen_delay.get()

        if delay > 0:

            self.update_status(f"キャプチャまで {delay} 秒待機中...")

            for i in range(delay, 0, -1):

                self.update_status(f"キャプチャまで {i} 秒...")

                time.sleep(1)

        

        self.update_status("全画面キャプチャ中...")

        

        # スクリーンショットを取得

        screenshot = ImageGrab.grab()

        

        # 保存とクリップボード処理

        self.process_screenshot(screenshot, "fullscreen")

        

    def capture_region(self):

        """範囲選択キャプチャ（Snipping Tool使用）"""

        self.update_status("範囲選択モード起動中...")

        

        # Snipping Toolを起動（Windows 11の場合）

        try:

            # Windows 11の新しいSnipping Tool

            subprocess.Popen(["ms-screenclip:"])

        except:

            try:

                # 従来のSnipping Tool

                subprocess.Popen(["SnippingTool.exe", "/clip"])

            except:

                messagebox.showerror("エラー", "Snipping Toolが見つかりません")

                return

        

        # クリップボード監視スレッドが既に実行中なら中止

        self.clipboard_check_running = True

        

        # Snipping Toolでキャプチャ後、クリップボードを監視

        threading.Thread(target=self.check_clipboard_for_image, daemon=True).start()

    

    def check_clipboard_for_image(self):

        """クリップボードの画像を確認（スレッド処理）"""

        max_wait_time = self.region_wait_time.get()

        start_time = time.time()

        

        while self.clipboard_check_running and time.time() - start_time < max_wait_time:

            try:

                # クリップボードから画像を取得

                screenshot = ImageGrab.grabclipboard()

                if isinstance(screenshot, Image.Image):

                    self.clipboard_check_running = False

                    # UIスレッドで処理するために呼び出し

                    self.root.after(0, lambda: self.process_screenshot(screenshot, "region"))

                    return

                time.sleep(0.5)

            except Exception as e:

                print(f"クリップボード確認エラー: {e}")

                time.sleep(0.5)

        

        # タイムアウト

        if self.clipboard_check_running:

            self.clipboard_check_running = False

            self.root.after(0, lambda: self.update_status("範囲選択タイムアウト - 画像が取得できませんでした"))

    

    def process_screenshot(self, screenshot, capture_type):

        """スクリーンショットの処理"""

        if not screenshot:

            self.update_status("スクリーンショットの取得に失敗しました")

            return

            

        # クリップボードにコピー

        if self.copy_to_clipboard.get():

            self.copy_image_to_clipboard(screenshot)

        

        # 自動保存

        if self.auto_save.get():

            filename = self.generate_filename(capture_type)

            filepath = self.save_folder / filename

            

            # フォルダが存在しない場合は作成

            self.save_folder.mkdir(parents=True, exist_ok=True)

            

            # 保存

            screenshot.save(filepath, "PNG")

            self.update_status(f"保存完了: {filename}")

            

            # プレビューウィンドウを表示（任意）

            self.show_notification(f"スクリーンショット保存: {filename}", str(filepath))

        else:

            self.update_status("キャプチャ完了（クリップボードにコピー）")

            

        # 効果音を鳴らす

        if self.play_sound.get():

            self.play_capture_sound()

    

    def play_capture_sound(self):

        """キャプチャ時の効果音を鳴らす"""

        try:

            if self.sound_file == "default":

                winsound.MessageBeep(winsound.MB_OK)

            elif self.sound_file == "ding":

                winsound.MessageBeep(winsound.MB_ICONASTERISK)

            elif self.sound_file == "chime":

                winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)

            elif self.sound_file == "notify":

                winsound.PlaySound("SystemNotification", winsound.SND_ALIAS)

            elif self.sound_file == "recycle":

                winsound.PlaySound("SystemRecycle", winsound.SND_ALIAS)

            elif os.path.exists(self.sound_file):

                # カスタム音声ファイル

                winsound.PlaySound(self.sound_file, winsound.SND_FILENAME)

        except Exception as e:

            print(f"効果音エラー: {e}")

    

    def test_sound(self):

        """効果音をテスト再生"""

        # 現在選択されている効果音を取得

        self.sound_file = self.sound_var.get()

        self.play_capture_sound()

    

    def select_custom_sound(self):

        """カスタム効果音を選択"""

        sound_file = filedialog.askopenfilename(

            title="効果音ファイルを選択",

            filetypes=[("WAVファイル", "*.wav"), ("すべてのファイル", "*.*")]

        )

        if sound_file:

            self.sound_file = sound_file

            self.sound_var.set(os.path.basename(sound_file))

    

    def copy_image_to_clipboard(self, image):

        """画像をクリップボードにコピー"""

        output = BytesIO()

        image.save(output, "BMP")

        data = output.getvalue()[14:]  # BMPヘッダーを除去

        output.close()

        

        win32clipboard.OpenClipboard()

        win32clipboard.EmptyClipboard()

        win32clipboard.SetClipboardData(win32con.CF_DIB, data)

        win32clipboard.CloseClipboard()

    

    def generate_filename(self, capture_type):

        """ファイル名を生成"""

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        return f"screenshot_{capture_type}_{timestamp}.png"

    

    def change_folder(self):

        """保存先フォルダを変更"""

        folder = filedialog.askdirectory(initialdir=self.save_folder)

        if folder:

            self.save_folder = Path(folder)

            self.folder_label.config(text=str(self.save_folder))

            self.update_status(f"保存先を変更: {self.save_folder}")

    

    def update_status(self, message):

        """ステータスバーを更新"""

        self.status_var.set(message)

        print(message)  # デバッグ用

        self.root.update_idletasks()

    

    def show_notification(self, title, message):

        """通知を表示（Windows 10/11のトースト通知）"""

        try:

            from win10toast import ToastNotifier

            toaster = ToastNotifier()

            toaster.show_toast(title, message, duration=3, threaded=True)

        except ImportError:

            # 通知ライブラリがない場合は小さいウィンドウで表示

            self.root.after(0, lambda: self.show_preview_dialog(title, message))

    

    def show_preview_dialog(self, title, message):

        """スクリーンショットのプレビューダイアログを表示"""

        preview = tk.Toplevel(self.root)

        preview.title("スクリーンショット完了")

        preview.geometry("300x100")

        preview.attributes('-topmost', True)

        

        ttk.Label(preview, text=title, font=("", 10, "bold")).pack(pady=(10, 5))

        ttk.Label(preview, text=message).pack()

        ttk.Button(preview, text="OK", command=preview.destroy).pack(pady=10)

        

        # 5秒後に自動で閉じる

        preview.after(5000, preview.destroy)

    

    def setup_tray_icon(self):

        """タスクトレイアイコンをセットアップ"""

        # アイコン画像（実際のアイコンに置き換えてください）

        icon_image = PilImage.open(self.resource_path("screenshot_icon.ico"))

        

        # アイコンメニュー

        menu = (

            item('全画面キャプチャ', self.capture_fullscreen),

            item('範囲選択キャプチャ', self.capture_region),

            item('ウィンドウを表示', self.show_window),

            item('終了', self.exit_app)

        )

        

            # アイコン作成

        self.icon = pystray.Icon("ScreenshotTool", icon_image, "スクリーンショット補助ツール", menu)

    

    def show_window(self, _=None):

        """ウィンドウを表示（タスクトレイから）"""

        self.root.deiconify()

        self.root.lift()

        self.root.focus_force()

    

    def minimize_to_tray(self):

        """ウィンドウをタスクトレイに最小化"""

        self.root.withdraw()

        # アイコンがまだ作成されていない場合は作成

        if not self.icon:

            self.setup_tray_icon()

            # 別スレッドでアイコンを実行

            threading.Thread(target=self.icon.run, daemon=True).start()

    

    def exit_app(self, _=None):

        """アプリケーションを終了"""

        # 設定を保存

        self.save_config()

        

        # キーボードフックを解除

        keyboard.unhook_all()

        

        # アイコンがある場合は停止

        if self.icon:

            self.icon.stop()

        

        # ウィンドウを破棄

        self.root.destroy()

        sys.exit(0)

    

    def on_closing(self):

        """ウィンドウの閉じるボタンが押されたときの処理"""

        # タスクトレイに最小化

        self.minimize_to_tray()

    

    def run(self):

        """アプリケーションを実行"""

        # 終了時の処理

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        

        # タスクトレイアイコンの準備

        self.setup_tray_icon()

        

        # タスクトレイアイコンを別スレッドで実行

        threading.Thread(target=self.icon.run, daemon=True).start()

        

        # 最小化オプションが有効なら最小化

        if self.config.get("minimize_on_startup", False):

            self.root.after(500, self.minimize_to_tray)

        

        # メインループ

        self.root.mainloop()



# EXE化のための準備

def resource_path(relative_path):

    """リソースファイルへのパスを取得（exe化時に必要）"""

    try:

        base_path = sys._MEIPASS

    except Exception:

        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)



# 実行

if __name__ == "__main__":

    try:

        app = SnippingToolWrapper()

        app.run()

    except Exception as e:

        print(f"エラーが発生しました: {e}")

        import traceback

        traceback.print_exc()

        messagebox.showerror("エラー", f"アプリケーションでエラーが発生しました:\n{e}")
