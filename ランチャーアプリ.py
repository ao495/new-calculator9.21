import tkinter as tk
from tkinter import simpledialog, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import threading, subprocess, time, json, os
import pystray
from PIL import Image, ImageDraw

APP_JSON = "apps.json"

# -----------------------------
# データ管理
# -----------------------------
app_groups = {}
running_processes = {}
app_status = {}
tab_timers = {}
tab_running_flags = {}
tab_tray_icons = {}
status_trees = {}
tabs = {}

def load_apps():
    global app_groups
    if os.path.exists(APP_JSON):
        with open(APP_JSON, "r", encoding="utf-8") as f:
            app_groups = json.load(f)
    else:
        app_groups = {}

def save_apps():
    with open(APP_JSON, "w", encoding="utf-8") as f:
        json.dump(app_groups, f, ensure_ascii=False, indent=2)

# -----------------------------
# トレイ・タイマー関連
# -----------------------------
def create_image(color="blue"):
    img = Image.new('RGB', (16,16), color='white')
    d = ImageDraw.Draw(img)
    d.rectangle([0,0,15,15], fill=color)
    return img

def update_tray(tab_name):
    tray_icon = tab_tray_icons[tab_name]
    while tab_running_flags.get(tab_name, False):
        seconds = tab_timers.get(tab_name,0)
        mins, secs = divmod(seconds, 60)
        tray_icon.title = f"{tab_name} 残り時間: {mins:02}:{secs:02}"
        time.sleep(1)

def kill_apps(tab_name):
    for proc in running_processes.get(tab_name, []):
        try: proc.terminate()
        except: pass
    for app in app_status.get(tab_name, []):
        if app['status'] == '起動中': app['status'] = 'タイマー終了'
    running_processes[tab_name] = []

def start_timer(tab_name, seconds):
    tray_icon = pystray.Icon(tab_name, create_image("blue"), f"{tab_name} タイマー")
    tab_tray_icons[tab_name] = tray_icon
    threading.Thread(target=tray_icon.run, daemon=True).start()

    tab_timers[tab_name] = seconds
    tab_running_flags[tab_name] = True
    threading.Thread(target=update_tray, args=(tab_name,), daemon=True).start()

    while tab_timers[tab_name] > 0 and tab_running_flags[tab_name]:
        time.sleep(1)
        tab_timers[tab_name] -= 1
        update_status_table()

    tab_running_flags[tab_name] = False
    tray_icon.visible = False
    kill_apps(tab_name)
    update_status_table()

def set_tab_timer(tab_name):
    seconds = simpledialog.askinteger("タブタイマー", f"{tab_name} タブの全アプリ終了までの時間（秒）:")
    if seconds:
        threading.Thread(target=start_timer, args=(tab_name, seconds), daemon=True).start()

# -----------------------------
# アプリ起動・状態管理
# -----------------------------
def run_apps(tab_name):
    running_processes[tab_name] = []
    app_status[tab_name] = []
    for app in app_groups[tab_name]:
        try:
            proc = subprocess.Popen(app)
            running_processes[tab_name].append(proc)
            app_status[tab_name].append({'name': app, 'proc': proc, 'status': '起動中'})
        except Exception as e:
            app_status[tab_name].append({'name': app, 'proc': None, 'status': f'起動失敗: {e}'})
    update_status_table()

def update_status_table():
    for tree in status_trees.values():
        for item in tree.get_children(): tree.delete(item)
    combined_list = []
    for tab, apps in app_status.items():
        for a in apps:
            if a['proc'] and a['status'] == '起動中':
                if a['proc'].poll() is not None: a['status'] = 'ユーザー終了'
            combined_list.append({'tab': tab, 'name': a['name'], 'status': a['status']})
    combined_list.sort(key=lambda x: x['status'] != '起動中')
    for app in combined_list:
        status_trees['起動中一覧'].insert('', 'end', values=(app['tab'], app['name'], app['status']))

def on_tab_changed(event):
    selected_tab = tab_control.tab(tab_control.select(), "text")
    for tab_name, tray_icon in tab_tray_icons.items():
        if tab_running_flags.get(tab_name, False):
            color = "green" if tab_name == selected_tab else "blue"
            tray_icon.icon = create_image(color)

# -----------------------------
# ドラッグ＆ドロップ登録
# -----------------------------
def drop_app(event, tab_name, scroll_frame):
    files = root.tk.splitlist(event.data)
    for f in files:
        if os.path.isfile(f):
            app_groups[tab_name].append(f)
            save_apps()
            add_app_button(scroll_frame, tab_name, f)

def add_app_button(frame, tab_name, app_path):
    btn = tk.Button(frame, text=os.path.basename(app_path), width=20,
                    command=lambda a=app_path, n=tab_name: run_apps(n))
    btn.bind("<Button-3>", lambda e, n=tab_name, a=app_path, b=btn: on_app_right_click(e,n,a,b))
    row = len([w for w in frame.grid_slaves() if isinstance(w, tk.Button)]) // 3
    col = len([w for w in frame.grid_slaves() if isinstance(w, tk.Button)]) % 3
    btn.grid(row=row, column=col, padx=5, pady=5)

def on_app_right_click(event, tab_name, app_path, btn):
    menu = tk.Menu(root, tearoff=0)
    menu.add_command(label="削除", command=lambda: delete_app(tab_name, app_path, btn))
    index = app_groups[tab_name].index(app_path)
    menu.add_command(label="↑上に移動", command=lambda: move_app_up(tab_name, index))
    menu.add_command(label="↓下に移動", command=lambda: move_app_down(tab_name, index))
    menu.post(event.x_root, event.y_root)

def delete_app(tab_name, app_path, btn):
    if app_path in app_groups[tab_name]:
        app_groups[tab_name].remove(app_path)
        save_apps()
        btn.destroy()
        refresh_tab_buttons(tab_name)

def move_app_up(tab_name, index):
    if index > 0:
        app_groups[tab_name][index-1], app_groups[tab_name][index] = app_groups[tab_name][index], app_groups[tab_name][index-1]
        save_apps()
        refresh_tab_buttons(tab_name)

def move_app_down(tab_name, index):
    if index < len(app_groups[tab_name]) - 1:
        app_groups[tab_name][index+1], app_groups[tab_name][index] = app_groups[tab_name][index], app_groups[tab_name][index+1]
        save_apps()
        refresh_tab_buttons(tab_name)

def refresh_tab_buttons(tab_name):
    frame = tabs[tab_name]
    scroll_frame = frame.winfo_children()[0].winfo_children()[0]
    for w in scroll_frame.winfo_children():
        if isinstance(w, tk.Button): w.destroy()
    for app in app_groups[tab_name]:
        add_app_button(scroll_frame, tab_name, app)
    btn_timer = tk.Button(scroll_frame, text="タイマー付き終了", command=lambda n=tab_name: set_tab_timer(n))
    btn_timer.grid(row=1000, column=0, columnspan=3, pady=5)

# -----------------------------
# タブ管理
# -----------------------------
def add_new_tab():
    name = simpledialog.askstring("新規タブ", "タブ名を入力:")
    if name and name not in app_groups:
        app_groups[name] = []
        save_apps()
        create_tab(name)

def create_tab(name):
    frame = ttk.Frame(tab_control)
    tabs[name] = frame
    tab_control.add(frame, text=name)

    canvas = tk.Canvas(frame)
    scrollbar = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas)
    scroll_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
    canvas.create_window((0,0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    for app in app_groups[name]:
        add_app_button(scroll_frame, name, app)

    btn_timer = tk.Button(scroll_frame, text="タイマー付き終了", command=lambda n=name: set_tab_timer(n))
    btn_timer.grid(row=1000, column=0, columnspan=3, pady=5)

    scroll_frame.drop_target_register(DND_FILES)
    scroll_frame.dnd_bind('<<Drop>>', lambda e, n=name, f=scroll_frame: drop_app(e,n,f))

# -----------------------------
# GUI構築
# -----------------------------
load_apps()
root = TkinterDnD.Tk()
root.title("起動順＆一括タイマーランチャー")
root.geometry("700x500")

tab_control = ttk.Notebook(root)
tab_control.pack(expand=1, fill='both')

for name in app_groups.keys():
    create_tab(name)

btn_add_tab = tk.Button(root, text="+ タブ追加", command=add_new_tab)
btn_add_tab.pack(side="top", anchor="ne", padx=5, pady=5)

status_tab = ttk.Frame(tab_control)
tab_control.add(status_tab, text="起動中一覧")
cols = ('グループ名','アプリ名','状態')
tree = ttk.Treeview(status_tab, columns=cols, show='headings')
for c in cols: tree.heading(c,text=c)
tree.pack(expand=1, fill='both')
status_trees['起動中一覧'] = tree

tab_control.bind("<<NotebookTabChanged>>", lambda e: on_tab_changed(e))
root.after(1000, lambda: periodic_update())
def periodic_update(): update_status_table(); root.after(1000, periodic_update)

root.mainloop()
