import PySimpleGUI as sg
import yt_dlp
import os
import shutil
import json
import re
import sys
import threading
import time
import winsound

CONFIG_FILE = 'config.json'

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_config(config):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', name)

def show_help_window():
    help_text = """
    ## YouTube Downloader 使い方・注意事項

    ### 1. 認証用クッキー
    - このアプリでは **cookies.txt** のみを利用します。
    - 他のブラウザからの直接読み込み（Chrome/Edge/Firefox）はサポートしていません。
    - `cookies.txt` は YouTube にログインした状態でブラウザからエクスポートしてください。
    - アプリと同じフォルダに置くだけでOKです。

    ### 2. cookies.txt の取得手順
    1. お使いのブラウザに「Get cookies.txt LOCALLY」のような拡張機能をインストール。
    2. YouTube にログインした状態で拡張機能を使い cookies.txt をエクスポート。
    3. ファイルをアプリと同じフォルダに配置。
    4. ダウンロード時は自動で読み込みます。

    ### 3. 注意点
    - ブラウザを閉じる必要はありません。
    - セッション切れやパスワード変更時は再度 cookies.txt を取得してください。
    - exe化した場合も同じディレクトリに置くことで利用可能です。

    ### 4. ダウンロードの進行
    - ダウンロード中はプログレスバーと速度・残り時間が表示されます。
    - 個別ダウンロードごとに完了通知あり。
    - エラー時はログに記録され、次の動画のダウンロードに進みます。
    """
    layout = [
        [sg.Multiline(help_text, size=(80, 25), font=('Arial', 10), disabled=True)],
        [sg.Button('閉じる')]
    ]
    window = sg.Window('使い方・ヘルプ', layout, modal=True)
    while True:
        event, _ = window.read()
        if event in (sg.WIN_CLOSED, '閉じる'):
            break
    window.close()

def download_video(url, output_path, format_choice, progress_callback, cancel_event, log_list):
    # exe化対応
    if getattr(sys, 'frozen', False):
        app_dir = os.path.dirname(sys.executable)
    else:
        app_dir = os.path.dirname(os.path.abspath(__file__))

    os.makedirs(output_path, exist_ok=True)

    def hook(d):
        if cancel_event.is_set():
            raise yt_dlp.utils.DownloadError("ユーザーによってキャンセルされました。")
        if d['status'] == 'downloading':
            try:
                percent = float(d.get('_percent_str','0%').replace('%',''))
            except:
                percent = 0
            progress_callback({'status':'downloading','percent':percent,
                               'speed':d.get('_speed_str','不明'), 'eta':d.get('_eta_str','不明')})
        elif d['status'] == 'finished':
            progress_callback({'status':'finished','filename':d['filename']})

    ydl_opts = {
        'outtmpl': os.path.join(output_path, '%(title)s.%(ext)s'),
        'noplaylist': True,
        'progress_hooks':[hook],
        'continuedl':True,
        'ignoreerrors':False,
        'retries':3,
        'cookiefile': os.path.join(app_dir, 'cookies.txt')
    }

    if format_choice == 'MP3':
        ydl_opts['format'] = 'bestaudio/best'
        ydl_opts['postprocessors'] = [{'key':'FFmpegExtractAudio','preferredcodec':'mp3'}]
    else:
        ydl_opts['format'] = format_choice

    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=True)
            title = sanitize_filename(info.get('title','video'))
            ext = 'mp3' if format_choice=='MP3' else info.get('ext','mp4')
            final_file = os.path.join(output_path, f"{title}.{ext}")
            winsound.Beep(1000,200)
    except yt_dlp.utils.DownloadError as e:
        log_list.append(f"Error: {url} -> {str(e)}")
        progress_callback({'status':'error','message':str(e)})
    except Exception as e:
        log_list.append(f"Error: {url} -> {str(e)}")
        progress_callback({'status':'error','message':str(e)})

def main():
    config = load_config()
    last_folder = config.get('last_folder', os.path.expanduser('~'))
    sg.theme('LightBlue3')

    layout = [
        [sg.Text('YouTube URL1:'), sg.InputText(key='-URL1-')],
        [sg.Text('YouTube URL2:'), sg.InputText(key='-URL2-')],
        [sg.Text('YouTube URL3:'), sg.InputText(key='-URL3-')],
        [sg.Text('保存先:'), sg.InputText(default_text=last_folder, key='-FOLDER-'), sg.FolderBrowse()],
        [sg.Text('フォーマット:'), sg.Combo(['高画質','標準画質','MP3'],default_value='高画質', key='-FORMAT-', size=(15,1))],
        [sg.Button('ダウンロード', key='-DOWNLOAD-'), sg.Button('キャンセル', key='-CANCEL-', disabled=True),
         sg.Button('終了'), sg.Button('使い方', key='-HELP-')],
        [sg.ProgressBar(100, orientation='h', size=(40, 20), key='-PROGRESSBAR-')],
        [sg.Text('ステータス:', size=(10,1)), sg.Text('', key='-STATUS-')]
    ]
    window = sg.Window('YouTube Downloader', layout, finalize=True)

    format_map = {
        '高画質':'bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]',
        '標準画質':'best[ext=mp4]',
        'MP3':'MP3'
    }

    cancel_event = threading.Event()
    download_thread = None
    log_list = []

    def progress_update(d):
        if d['status']=='downloading':
            window['-PROGRESSBAR-'].update(d['percent'])
            window['-STATUS-'].update(f"進行中 {d['percent']:.1f}% | 速度: {d['speed']} | 残り: {d['eta']}")
        elif d['status']=='finished':
            window['-STATUS-'].update('完了 🎉')
            window['-PROGRESSBAR-'].update(0)
            window['-DOWNLOAD-'].update(disabled=False)
            window['-CANCEL-'].update(disabled=True)
        elif d['status']=='error':
            window['-STATUS-'].update(f"エラー: {d['message']}")
            window['-PROGRESSBAR-'].update(0)
            window['-DOWNLOAD-'].update(disabled=False)
            window['-CANCEL-'].update(disabled=True)

    while True:
        event, values = window.read(timeout=100)
        if event in (sg.WIN_CLOSED, '終了'):
            break
        if event=='-HELP-':
            show_help_window()
        if event=='-DOWNLOAD-':
            urls = [values['-URL1-'], values['-URL2-'], values['-URL3-']]
            urls = [u for u in urls if u.strip()]
            output_path = values['-FOLDER-']
            if not urls:
                sg.popup_error("URLを入力してください")
                continue
            if not output_path:
                sg.popup_error("保存先を選択してください")
                continue
            config['last_folder']=output_path
            save_config(config)

            window['-STATUS-'].update('ダウンロード開始...')
            window['-PROGRESSBAR-'].update(0)
            window['-DOWNLOAD-'].update(disabled=True)
            window['-CANCEL-'].update(disabled=False)

            cancel_event.clear()
            format_choice = format_map[values['-FORMAT-']]
            def thread_target():
                for url in urls:
                    if cancel_event.is_set():
                        break
                    download_video(url, output_path, format_choice, progress_update, cancel_event, log_list)
                if log_list:
                    sg.popup_scrolled('\n'.join(log_list), title='エラー一覧')
                window['-DOWNLOAD-'].update(disabled=False)
                window['-CANCEL-'].update(disabled=True)
            download_thread = threading.Thread(target=thread_target, daemon=True)
            download_thread.start()
        if event=='-CANCEL-':
            cancel_event.set()
            window['-STATUS-'].update('キャンセル中...')

    window.close()

if __name__=="__main__":
    main()
