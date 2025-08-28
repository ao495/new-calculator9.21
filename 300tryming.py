import shutil
import sys
import os
from pathlib import Path
from tkinter import Tk, Label, Button, filedialog, messagebox, Checkbutton, IntVar, Frame, Canvas, Scrollbar, Scale, HORIZONTAL
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image, ImageTk, UnidentifiedImageError
from datetime import datetime

# --- 定数 ---
TARGET_WIDTH = 300

def get_base_dir():
    """ .exe化した場合と.pyで実行した場合で、基準となるパスを正しく取得する """
    if getattr(sys, 'frozen', False):
        # .exeとして実行されている場合
        return Path(sys.executable).parent
    else:
        # .pyスクリプトとして実行されている場合
        return Path(__file__).parent

SCRIPT_DIR = get_base_dir()

# --- アプリケーションクラス ---
class ImageProcessorApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("画像リサイズ＆選択ツール v3.2 (Final EXE)")
        self.geometry("750x600")

        # --- 状態変数 ---
        self.checkbox_data = []
        self.quality_var = IntVar(value=95)
        self.webp_var = IntVar(value=0)

        # --- GUIの構築 ---
        self._build_gui()

    def _build_gui(self):
        top_frame = Frame(self)
        top_frame.pack(fill="x", padx=10, pady=5)
        drop_label = Label(top_frame, text="ここに画像をドラッグ＆ドロップしてください", bg="lightgray", height=3)
        drop_label.pack(fill="x", expand=True)
        drop_label.drop_target_register(DND_FILES)
        drop_label.dnd_bind('<<Drop>>', self.handle_drop_files)

        options_frame = Frame(self, bd=2, relief="groove", padx=5, pady=5)
        options_frame.pack(fill="x", padx=10, pady=5)

        Label(options_frame, text="保存オプション:").pack(side="left", padx=5)
        Checkbutton(options_frame, text="WebP形式で保存", variable=self.webp_var).pack(side="left")
        
        self.quality_label = Label(options_frame, text=f"JPEG品質: {self.quality_var.get()}")
        self.quality_label.pack(side="left", padx=(15, 0))
        quality_slider = Scale(options_frame, from_=50, to=100, orient=HORIZONTAL, variable=self.quality_var, command=lambda v: self.quality_label.config(text=f"JPEG品質: {v}"))
        quality_slider.pack(side="left", fill="x", expand=True)

        btn_frame = Frame(self)
        btn_frame.pack(fill="x", padx=10, pady=5)
        Button(btn_frame, text="ファイル選択", command=self.handle_select_files).pack(side="left")
        Button(btn_frame, text="リストをクリア", command=self.clear_list).pack(side="left", padx=5)
        Button(btn_frame, text="選択した画像を保存", command=self.save_selected_files, bg="#007bff", fg="white").pack(side="right")

        canvas_container = Frame(self)
        canvas_container.pack(fill="both", expand=True, padx=10, pady=10)
        self.canvas = Canvas(canvas_container)
        scrollbar = Scrollbar(canvas_container, orient="vertical", command=self.canvas.yview)
        self.canvas_content_frame = Frame(self.canvas)

        self.canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.create_window((0, 0), window=self.canvas_content_frame, anchor="nw")

        self.canvas_content_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def handle_drop_files(self, event):
        files = self.tk.splitlist(event.data)
        self.process_images(files)

    def handle_select_files(self):
        files = filedialog.askopenfilenames(title="画像を選択してください", filetypes=[("画像ファイル", "*.jpg *.jpeg *.png")])
        if files:
            self.process_images(list(files))

    def clear_list(self):
        self.cleanup_temp_folders()
        self.checkbox_data.clear()
        self.show_thumbnails()

    def process_images(self, file_list):
        self.clear_list()
        processed_count = 0
        source_folder = Path(file_list[0]).parent
        output_square_dir = source_folder / "output_square_temp"
        output_resize_dir = source_folder / "output_resize_temp"
        output_square_dir.mkdir(exist_ok=True)
        output_resize_dir.mkdir(exist_ok=True)

        for file in file_list:
            try:
                p_file = Path(file)
                img = Image.open(p_file).convert("RGB")
                square_img = crop_center_square(img).resize((TARGET_WIDTH, TARGET_WIDTH))
                square_path = output_square_dir / p_file.name
                square_img.save(square_path, "JPEG")
                w, h = img.size
                new_height = int(h * TARGET_WIDTH / w)
                resized_img = img.resize((TARGET_WIDTH, new_height))
                resize_path = output_resize_dir / p_file.name
                resized_img.save(resize_path, "JPEG")
                self.checkbox_data.append({
                    "original_path": p_file,
                    "square_path": square_path, "square_var": IntVar(),
                    "resize_path": resize_path, "resize_var": IntVar()
                })
                processed_count += 1
            except UnidentifiedImageError:
                print(f"スキップ(非画像): {file}")
            except Exception as e:
                print(f"エラー: {file}, {e}")

        if processed_count > 0:
            self.show_thumbnails()
        else:
            messagebox.showwarning("処理失敗", "処理できる画像がありませんでした。")

    def show_thumbnails(self):
        for widget in self.canvas_content_frame.winfo_children():
            widget.destroy()
        
        self.canvas.yview_moveto(0)

        labels = {"square": "正方形", "resize": "縦横比維持"}
        for idx, data in enumerate(self.checkbox_data):
            row_frame = Frame(self.canvas_content_frame, bd=2, relief="groove")
            row_frame.grid(row=idx, column=0, padx=5, pady=5, sticky="w")

            for i, key in enumerate(["square", "resize"]):
                thumb_frame = Frame(row_frame)
                thumb_frame.grid(row=0, column=i, padx=5, pady=5)
                path_key, var_key = f"{key}_path", f"{key}_var"
                img = Image.open(data[path_key])
                thumb = ImageTk.PhotoImage(img)
                
                Label(thumb_frame, text=labels[key], font=("TkDefaultFont", 9, "bold")).pack()
                Label(thumb_frame, image=thumb).pack()
                Checkbutton(thumb_frame, text="選択", variable=data[var_key]).pack()
                data[f"{key}_thumb"] = thumb

    def save_selected_files(self):
        if not self.checkbox_data:
            messagebox.showwarning("リストなし", "処理する画像がありません。")
            return

        has_double_selection = False
        has_zero_selection = False
        total_selected = 0
        for data in self.checkbox_data:
            s_var = data["square_var"].get() == 1
            r_var = data["resize_var"].get() == 1
            if s_var and r_var: has_double_selection = True
            if not s_var and not r_var: has_zero_selection = True
            if s_var: total_selected += 1
            if r_var: total_selected += 1

        if total_selected == 0:
            messagebox.showwarning("選択なし", "保存する画像が選択されていません。")
            return

        if has_double_selection:
            if not messagebox.askyesno("確認", "両方選択されているペアがあります。\nこのまま両方のバージョンを保存しますか？"):
                return

        if has_zero_selection:
            if not messagebox.askyesno("確認", "選択されていないペアがあります。\nチェックした画像のみを保存してよろしいですか？"):
                return

        date_str = datetime.now().strftime('%Y-%m-%d')
        output_dir = SCRIPT_DIR / "output_final" / date_str
        output_dir.mkdir(parents=True, exist_ok=True)

        for data in self.checkbox_data:
            for key in ["square", "resize"]:
                if data[f"{key}_var"].get() == 1:
                    img_path = data[f"{key}_path"]
                    original_img = Image.open(img_path)
                    save_as_webp = self.webp_var.get() == 1
                    quality = self.quality_var.get()
                    ext = ".webp" if save_as_webp else ".jpg"
                    final_path = output_dir / (data["original_path"].stem + f"_{key}" + ext)
                    original_img.save(final_path, quality=quality, optimize=True)
        
        messagebox.showinfo("保存完了", f"{total_selected}個の画像を\n{output_dir}\nに保存しました。")
        
        try:
            os.startfile(output_dir)
        except Exception as e:
            print(f"フォルダを開けませんでした: {e}")

        self.clear_list()

    def cleanup_temp_folders(self):
        temp_dirs = set()
        for data in self.checkbox_data:
            temp_dirs.add(data["square_path"].parent)
            temp_dirs.add(data["resize_path"].parent)
        for temp_dir in temp_dirs:
            if temp_dir.exists():
                shutil.rmtree(temp_dir)

def crop_center_square(img):
    w, h = img.size
    min_side = min(w, h)
    left = (w - min_side) // 2
    top = (h - min_side) // 2
    return img.crop((left, top, left + min_side, top + min_side))

if __name__ == "__main__":
    app = ImageProcessorApp()
    app.mainloop()