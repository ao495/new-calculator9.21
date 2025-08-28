import os
import json
import threading
import math
import tempfile
from datetime import datetime
from tkinter import Tk, filedialog, Label, Button, Checkbutton, IntVar, Frame, Scrollbar, Canvas, messagebox, Toplevel, StringVar, Entry, OptionMenu, Spinbox
from PIL import Image, ImageTk
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

APP_NAME = "かんたん写真整理"
CONFIG_FILE = "config.json"
FOLDER_FILE = "last_folder.json"
BATCH_SIZE = 4
TARGET_HEIGHT_EXCEL = 300
PREVIEW_HEIGHT = 300

class PhotoApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.all_filepaths = []
        self.selection_vars = {}
        self.current_page = 0
        self.total_pages = 0
        self.preview_cache = {}
        self.settings_win = None

        # 設定
        self.default_ratios = {
            "landscape_top": 3219 / 3864,
            "landscape_left": 4011 / 5152,
            "portrait_top": 4470 / 5152
        }
        self.config = {"ratios": self.default_ratios.copy(), "save_format": "PNG", "jpeg_quality": 95}
        self.load_config()
        self.current_ratios = self.config.get("ratios", self.default_ratios.copy())
        self.save_format = self.config.get("save_format", "PNG")
        self.jpeg_quality = self.config.get("jpeg_quality", 95)

        self._setup_ui()

    # ---------------- 設定読み書き ----------------
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    self.config.update(json.load(f))
            except (json.JSONDecodeError, TypeError) as e:
                messagebox.showerror("エラー", f"設定ファイルの形式が正しくありません: {e}")
            except Exception as e:
                messagebox.showerror("エラー", f"設定ファイルの読み込みに失敗しました: {e}")

    def save_config(self):
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            messagebox.showerror("エラー", f"設定ファイルの保存に失敗しました: {e}")

    # ---------------- UI ----------------
    def _setup_ui(self):
        action_frame = Frame(self.root)
        action_frame.pack(pady=5)
        Button(action_frame, text="写真を選択", command=self.select_files, bg="lightblue").pack(side="left", padx=5)
        Button(action_frame, text="選択済み写真を保存", command=self.save_selected, bg="lightgreen").pack(side="left", padx=5)
        Button(action_frame, text="Excelに貼付", command=self.save_to_excel, bg="lightyellow").pack(side="left", padx=5)
        Button(action_frame, text="すべて選択解除", command=self.deselect_all, bg="lightgray").pack(side="left", padx=5)
        Button(action_frame, text="設定", command=self.open_settings).pack(side="left", padx=5)

        self.canvas = Canvas(self.root)
        self.scrollbar = Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.hscrollbar = Scrollbar(self.root, orient="horizontal", command=self.canvas.xview)
        self.scrollable_frame = Frame(self.canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0,0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set, xscrollcommand=self.hscrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        self.hscrollbar.pack(side="bottom", fill="x")

        self.status_label = Label(self.root, text="準備完了", bd=1, relief="sunken", anchor="w")
        self.status_label.pack(side="bottom", fill="x")
        self.canvas.bind("<MouseWheel>", self._on_mouse_wheel)

    # ---------------- ファイル選択・プレビュー ----------------
    def select_files(self):
        filepaths = filedialog.askopenfilenames(filetypes=[("画像","*.jpg *.jpeg *.png")])
        if not filepaths: return

        self.all_filepaths = list(filepaths)
        self.total_pages = math.ceil(len(self.all_filepaths) / BATCH_SIZE)
        self.selection_vars.clear()
        self.preview_cache.clear() # キャッシュをクリア

        for filepath in self.all_filepaths:
            try:
                with Image.open(filepath) as img:
                    w,h = img.size
                    if w>h:
                        for mode in ["landscape_top","landscape_left"]:
                            self.selection_vars[(filepath,mode)] = IntVar(value=1)
                    else:
                        self.selection_vars[(filepath,"portrait_top")] = IntVar(value=1)
            except Exception as e:
                messagebox.showerror("エラー", f"画像情報の読み取りに失敗: {os.path.basename(filepath)}")

        self.load_page(0)

    def load_page(self, page_number):
        if not (0<=page_number<self.total_pages): return
        self.current_page = page_number
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        start_index = self.current_page*BATCH_SIZE
        end_index = start_index + BATCH_SIZE
        filepaths_to_load = self.all_filepaths[start_index:end_index]

        for filepath in filepaths_to_load:
            try:
                with Image.open(filepath) as img:
                    self._add_previews_for_file(img, filepath)
            except Exception as e:
                messagebox.showerror("エラー", f"""プレビュー生成に失敗: {os.path.basename(filepath)}
{e}""")

        self._create_navigation_buttons()
        self._update_scroll_region()

    def _add_previews_for_file(self, img, filepath):
        w,h = img.size
        if w>h:
            for mode in ["landscape_top","landscape_left"]:
                self.add_preview(self.crop_image(img, mode), filepath, mode)
        else:
            self.add_preview(self.crop_image(img,"portrait_top"), filepath,"portrait_top")

    def add_preview(self, img, filepath, mode):
        var = self.selection_vars.get((filepath,mode))
        if not var: return

        # キャッシュを確認
        cache_key = (filepath, mode)
        if cache_key in self.preview_cache:
            tk_img = self.preview_cache[cache_key]
        else:
            resized_img = self.resize_to_height(img.copy(), PREVIEW_HEIGHT)
            tk_img = ImageTk.PhotoImage(resized_img)
            self.preview_cache[cache_key] = tk_img

        preview_frame = Frame(self.scrollable_frame)
        preview_frame.pack(anchor="nw", pady=5, padx=5)

        chk = Checkbutton(preview_frame, text=f"{os.path.basename(filepath)} ({mode})", variable=var)
        lbl = Label(preview_frame, image=tk_img)
        lbl.image = tk_img # 参照を保持

        chk.pack(anchor="w")
        lbl.pack(anchor="w")

    def _create_navigation_buttons(self):
        nav_frame = Frame(self.scrollable_frame)
        nav_frame.pack(pady=10)

        prev_button = Button(nav_frame, text="< 前へ", command=lambda: self.load_page(self.current_page-1))
        if self.current_page==0: prev_button.config(state="disabled")
        prev_button.pack(side="left", padx=5)

        page_label = Label(nav_frame, text=f"ページ {self.current_page+1} / {self.total_pages}")
        page_label.pack(side="left", padx=5)

        next_button = Button(nav_frame, text="次へ >", command=lambda: self.load_page(self.current_page+1))
        if self.current_page>=self.total_pages-1: next_button.config(state="disabled")
        next_button.pack(side="left", padx=5)

        self.status_label.config(text=f"{len(self.all_filepaths)}枚中 {self.current_page*BATCH_SIZE+1}-{min((self.current_page+1)*BATCH_SIZE,len(self.all_filepaths))}枚目を表示中")

    def _update_scroll_region(self):
        self.root.after(100, lambda: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

    def resize_to_height(self,img,target_height):
        if img.height<=target_height: return img
        try:
            resample_method = Image.Resampling.LANCZOS
        except AttributeError:
            resample_method = Image.LANCZOS # Fallback for older Pillow versions
        return img.resize((int(img.width*(target_height/img.height)),target_height), resample_method)

    # ---------------- 選択解除 ----------------
    def deselect_all(self):
        for var in self.selection_vars.values():
            var.set(0)

    # ---------------- 高画質保存 ----------------
    def save_selected(self):
        selected_items = [(fp, mode) for (fp, mode), var in self.selection_vars.items() if var.get()==1]
        if not selected_items:
            messagebox.showinfo("確認", "選択された写真がありません。")
            return
        output_dir = filedialog.askdirectory(title="保存先フォルダを選択")
        if not output_dir: return
        
        self.status_label.config(text=f"{len(selected_items)}枚の写真を保存中...")
        thread = threading.Thread(target=self._save_selected_thread, args=(selected_items, output_dir), daemon=True)
        thread.start()

    def _save_selected_thread(self, selected_items, output_dir):
        saved_count = 0
        for filepath, mode in selected_items:
            try:
                with Image.open(filepath) as img:
                    img_to_save = self.crop_image(img, mode)
                    base = os.path.splitext(os.path.basename(filepath))[0]
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    out_path = os.path.join(output_dir, f"{APP_NAME}_{base}_{mode}_{timestamp}.{self.save_format.lower()}")
                    
                    if self.save_format == "PNG":
                        img_to_save.save(out_path, format="PNG", compress_level=0)
                    else:
                        img_to_save.save(out_path, format="JPEG", quality=self.jpeg_quality)
                    saved_count += 1
            except Exception as e:
                self.root.after(0, messagebox.showerror, "エラー", f"""ファイルの保存に失敗しました: {os.path.basename(filepath)} ({mode})
{e}""")
        
        self.root.after(0, lambda: self.status_label.config(text=f"{saved_count}枚の写真を保存しました。"))
        self.root.after(0, lambda: messagebox.showinfo("完了", f"{saved_count}枚の写真を正常に保存しました。"))

    # ---------------- Excel貼付 ----------------
    def save_to_excel(self):
        selected_items = [(fp, mode) for (fp, mode), var in self.selection_vars.items() if var.get()==1]
        if not selected_items:
            messagebox.showinfo("確認", "選択された写真がありません。")
            return
        output_path = filedialog.asksaveasfilename(title="Excelファイルを保存", defaultextension=".xlsx", filetypes=[("Excelファイル","*.xlsx")])
        if not output_path: return

        self.status_label.config(text=f"{len(selected_items)}枚の写真をExcelに貼付中...")
        thread = threading.Thread(target=self._save_to_excel_thread, args=(selected_items, output_path), daemon=True)
        thread.start()

    def _save_to_excel_thread(self, selected_items, output_path):
        try:
            wb = Workbook()
            ws = wb.active
            row = 2
            
            for filepath, mode in selected_items:
                with tempfile.NamedTemporaryFile(suffix=".png", delete=True) as tmpfile:
                    tmp_path = tmpfile.name
                    with Image.open(filepath) as img:
                        img_to_save = self.crop_image(img, mode)
                        img_to_save = self.resize_to_height(img_to_save, TARGET_HEIGHT_EXCEL)
                        img_to_save.save(tmp_path, format="PNG")
                        
                        xl_img = XLImage(tmp_path)
                        ws.add_image(xl_img, f"A{row}")
                        ws.row_dimensions[row].height = TARGET_HEIGHT_EXCEL * 0.75 # ポイントに変換
                        row += int(TARGET_HEIGHT_EXCEL / 20) + 2

            wb.save(output_path)
            self.root.after(0, lambda: self.status_label.config(text=f"Excelファイルに保存しました。"))
            self.root.after(0, lambda: messagebox.showinfo("完了", f"""選択した写真をExcelファイルに保存しました。
{output_path}"""))

        except Exception as e:
            self.root.after(0, messagebox.showerror, "エラー", f"""Excelファイルへの保存中にエラーが発生しました。
{e}""")

    # ---------------- 画像トリミング ----------------
    def crop_image(self,img,mode="landscape_top"):
        w,h = img.size
        r = self.current_ratios
        if w>h:
            if mode=="landscape_top": return img.crop((0,0,w,int(h*r["landscape_top"])))
            if mode=="landscape_left": return img.crop((0,0,int(w*r["landscape_left"]),h))
        else:
            return img.crop((0,0,w,int(h*r["portrait_top"])))

    # ---------------- 設定ウィンドウ ----------------
    def on_settings_close(self):
        if self.settings_win:
            self.settings_win.destroy()
            self.settings_win = None

    def open_settings(self):
        if self.settings_win and self.settings_win.winfo_exists():
            self.settings_win.lift()
            return

        self.settings_win = Toplevel(self.root)
        self.settings_win.title("設定")
        self.settings_win.protocol("WM_DELETE_WINDOW", self.on_settings_close)

        frame = Frame(self.settings_win,padx=10,pady=10)
        frame.pack()
        self.ratio_entries = {}
        for i,(key,label) in enumerate({"landscape_top":"横写真（上部）比率","landscape_left":"横写真（左部）比率","portrait_top":"縦写真（上部）比率"}.items()):
            Label(frame,text=label).grid(row=i,column=0,sticky="w",pady=5)
            var = StringVar(value=str(self.current_ratios.get(key,self.default_ratios[key])))
            entry = Entry(frame,textvariable=var)
            entry.grid(row=i,column=1,padx=5,pady=5)
            self.ratio_entries[key]=var

        Label(frame,text="保存形式:").grid(row=3,column=0,sticky="w",pady=5)
        format_var = StringVar(value=self.save_format)
        OptionMenu(frame,format_var,"PNG","JPEG").grid(row=3,column=1,sticky="w")
        self.format_var = format_var

        Label(frame,text="JPEG画質(1-100):").grid(row=4,column=0,sticky="w",pady=5)
        quality_var = StringVar(value=str(self.jpeg_quality))
        Spinbox(frame,from_=1,to=100,textvariable=quality_var).grid(row=4,column=1,sticky="w")
        self.quality_var = quality_var

        btn_frame = Frame(self.settings_win,pady=10)
        btn_frame.pack()
        Button(btn_frame,text="保存",command=self.save_settings).pack(side="left",padx=5)
        Button(btn_frame,text="デフォルトに戻す",command=self.reset_ratios).pack(side="left",padx=5)
        Button(btn_frame,text="キャンセル",command=self.on_settings_close).pack(side="left",padx=5)

    def save_settings(self):
        try:
            new_ratios={key:float(var.get()) for key,var in self.ratio_entries.items()}
            if any(v<=0 for v in new_ratios.values()): raise ValueError("比率は0より大きい値を入力してください")
            self.current_ratios=new_ratios
            self.config["ratios"]=new_ratios

            self.save_format = self.format_var.get()
            self.config["save_format"] = self.save_format
            self.jpeg_quality = int(self.quality_var.get())
            self.config["jpeg_quality"]=self.jpeg_quality

            self.save_config()
            messagebox.showinfo("設定","保存しました")
            self.on_settings_close()
            self.preview_cache.clear() # 比率が変わったのでキャッシュをクリア
            self.load_page(self.current_page) # プレビューを再読み込み
        except Exception as e:
            messagebox.showerror("エラー",f"保存中にエラー:{e}")

    def reset_ratios(self):
        if messagebox.askyesno("確認","比率をデフォルトに戻しますか？"):
            self.current_ratios=self.default_ratios.copy()
            self.config["ratios"]=self.default_ratios.copy()
            self.save_config()
            messagebox.showinfo("設定","デフォルトに戻しました")
            for key,var in self.ratio_entries.items(): var.set(str(self.current_ratios[key]))
            self.preview_cache.clear() # 比率が変わったのでキャッシュをクリア
            self.load_page(self.current_page) # プレビューを再読み込み


    # ---------------- マウスホイール ----------------
    def _on_mouse_wheel(self,event):
        if event.num==5 or event.delta<0: self.canvas.yview_scroll(1,"units")
        elif event.num==4 or event.delta>0: self.canvas.yview_scroll(-1,"units")

# ---------------- 実行 ----------------
if __name__=="__main__":
    root=Tk()
    app=PhotoApp(root)
    root.mainloop()
