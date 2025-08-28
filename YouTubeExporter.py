import PySimpleGUI as sg
import threading
import yt_dlp
import os
import json

CONFIG_FILE = 'config.json'

# --- 設定ファイル操作 ---
def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except json.JSONDecodeError:
            return {}
    return {}

def save_config(config):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

# --- ダウンロード処理 ---
def download_video(url, output_path, format_choice, progress_callback, window):
    ffmpeg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ffmpeg', 'ffmpeg-8.0-full_build', 'bin', 'ffmpeg.exe')

    ydl_opts = {
        'format': format_choice if format_choice != 'mp3' else 'bestaudio/best',
        'outtmpl': os.path.join(output_path, '%(title)s.%(ext)s'),
        'noplaylist': True,
        'ffmpeg_location': ffmpeg_path,
        'windowsfilenames': True,
        'progress_hooks': [lambda d: progress_callback(d, window)],
        'retries': 10,
    }

    if format_choice == 'mp3':
        ydl_opts['postprocessors'] = [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
            'preferredquality': '192',
        }]

    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            ydl.download([url])
    except Exception as e:
        progress_callback({'status': 'error', 'message': str(e)}, window)

# --- GUI 更新 ---
def update_gui_progress(d, window):
    status = d.get('status')

    if status == 'downloading':
        total = d.get('total_bytes') or d.get('total_bytes_estimate') or 1
        downloaded = d.get('downloaded_bytes', 0)
        percent = int(downloaded / total * 100)
        window.write_event_value('-PROGRESS_UPDATE-', {'percent': percent})
    elif status in ['finished', 'completed']:
        window.write_event_value('-PROGRESS_FINISHED-', None)
    elif status == 'error':
        window.write_event_value('-PROGRESS_ERROR-', d.get('message', '不明なエラー'))

# --- メイン ---
def main():
    config = load_config()
    last_folder = config.get('last_folder', os.path.expanduser('~'))

    sg.theme('LightBlue3')

    layout = [
        [sg.Text('YouTube URL:', size=(15, 1)), sg.InputText(key='-URL-')],
        [sg.Text('保存先フォルダ:', size=(15, 1)), sg.InputText(default_text=last_folder, key='-FOLDER-'), sg.FolderBrowse('参照')],
        [sg.Text('フォーマット:', size=(15, 1)), sg.Combo(
            ['bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]', 'bestaudio[ext=m4a]', 'mp3'],
            default_value='bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]',
            key='-FORMAT-')],
        [sg.Button('ダウンロード開始', key='-DOWNLOAD-'), sg.Button('終了')],
        [sg.ProgressBar(100, orientation='h', size=(40, 20), key='-PROGRESSBAR-')],
        [sg.Text('ステータス:', size=(15, 1)), sg.Text('', size=(40, 1), key='-STATUS-')]
    ]

    window = sg.Window('YouTube Downloader', layout, finalize=True)

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == '終了':
            break

        if event == '-DOWNLOAD-':
            url = values['-URL-']
            output_path = values['-FOLDER-']
            format_choice = values['-FORMAT-']

            if not url:
                sg.popup_error('YouTube URLを入力してください。')
                continue
            if not output_path:
                sg.popup_error('保存先フォルダを選択してください。')
                continue

            os.makedirs(output_path, exist_ok=True)

            config['last_folder'] = output_path
            save_config(config)

            window['-STATUS-'].update('ダウンロード開始...')
            window['-PROGRESSBAR-'].update(0)

            threading.Thread(target=download_video, args=(url, output_path, format_choice, update_gui_progress, window), daemon=True).start()

        # --- カスタムイベント処理 ---
        if event == '-PROGRESS_UPDATE-':
            data = values[event]
            window['-PROGRESSBAR-'].update(data['percent'])
            window['-STATUS-'].update(f"ダウンロード中: {data['percent']}%")
        elif event == '-PROGRESS_FINISHED-':
            window['-STATUS-'].update('ダウンロード完了！')
            window['-PROGRESSBAR-'].update(100)
        elif event == '-PROGRESS_ERROR-':
            error_message = values[event]
            window['-STATUS-'].update(f"エラー: {error_message}")
            window['-PROGRESSBAR-'].update(0)

    window.close()

if __name__ == '__main__':
    main()
