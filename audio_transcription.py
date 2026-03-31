import os
import sys
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from faster_whisper import WhisperModel
from moviepy.editor import AudioFileClip
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime
import threading

# アプリケーション情報
APP_NAME = "TND_AudioTranscription"
APP_VERSION = "1.3.0"
APP_TITLE = f"TND audio_transcription v{APP_VERSION}"
APP_ICON_NAME = "TND_AudioTranscription01.ico"

# 単語登録キャッシュファイル名
HOTWORDS_FILE = "hotwords.json"
# 単語登録の上限数（hotwords枠223トークンに安全に収まる範囲）
MAX_HOTWORDS = 50

# ライセンス情報（アプリ内表示用）
LICENSE_TEXT = """\
─────────────────────────────────────
使用ライセンス情報

■ audio_transcription
MIT License
Copyright (c) 2024 HIROKI TANODA(TND)

本ソフトウェアおよび関連文書ファイル(以下「ソフトウェア」)のコピーを取得した
すべての人に対し、ソフトウェアを無制限に扱うことを無償で許可します。これには、
ソフトウェアのコピーを使用、複製、変更、結合、公開、頒布、サブライセンス、
および/または販売する権利、並びにソフトウェアを提供する相手に同じことを
許可する権利も無制限に含まれます。

上記の著作権表示および本許諾表示を、ソフトウェアのすべてのコピーまたは
重要な部分に記載するものとします。

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, \
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF \
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.

■ Whisper Model (faster-whisper-large-v3)
MIT License  https://huggingface.co/Systran/faster-whisper-large-v3

■ faster-whisper
MIT License  Copyright (c) 2023 SYSTRAN
https://github.com/SYSTRAN/faster-whisper

■ OpenAI Whisper (Model)
MIT License  Copyright (c) 2022 OpenAI
https://github.com/openai/whisper

■ MoviePy
MIT License  Copyright (c) 2015 Zulko
https://github.com/Zulko/moviepy

■ pandas
BSD 3-Clause License  Copyright (c) 2008-2011, AQR Capital Management, LLC
https://github.com/pandas-dev/pandas

■ openpyxl
MIT License  Copyright (c) 2010 openpyxl
https://foss.heptapod.net/openpyxl/openpyxl

■ CTranslate2
MIT License  Copyright (c) 2019 OpenNMT
https://github.com/OpenNMT/CTranslate2

■ NumPy
BSD 3-Clause License  Copyright (c) 2005-2024, NumPy Developers
https://github.com/numpy/numpy

■ PyAV (FFmpeg Python bindings)
BSD 3-Clause License  Copyright (c) 2013, Mike Boers
https://github.com/PyAV-Org/PyAV
─────────────────────────────────────\
"""

# 単語登録機能ヘルプテキスト
HOTWORDS_HELP_TEXT = """\
単語登録（カスタム辞書）機能について

■ 機能の概要
一般的な辞書には載っていない専門用語、社内用語、プロジェクト名、\
人名などをあらかじめ登録することで、AIによる文字起こしの誤変換を\
減らすことができます。

■ 登録数の目安と注意点
・登録上限： 最大50単語まで
・推奨登録数： 1回の会議につき 10〜20単語程度

⚠️ 重要：登録のコツ
単語を登録しすぎると、かえってAIが混乱し、関係のない会話まで\
無理やり登録単語に変換してしまう（誤認識が増える）可能性があります。
「どうしても間違えてほしくない重要な単語」に絞って登録するのが、\
最もきれいに文字起こしをするコツです。

■ 効果的な登録例
・固有名詞： 「TNDツール」「〇〇商事」「田野田」
・略語・業界用語： 「DX」「SaaS」「OJT」
・同音異義語の区別： 「あうとるっく（Outlook）」
  「きしょう（起床／気象）」など、文脈で間違えやすいもの\
"""


def get_hotwords_path():
    """単語登録ファイルのパスを取得"""
    app_dir = get_app_dir()
    return os.path.join(app_dir, HOTWORDS_FILE)


def load_hotwords():
    """保存済みの単語リストを読み込む"""
    path = get_hotwords_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                return data
        except (json.JSONDecodeError, IOError):
            pass
    return []


def save_hotwords(words):
    """単語リストをJSONファイルに保存する"""
    path = get_hotwords_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(words, f, ensure_ascii=False, indent=2)


def get_app_dir():
    """アプリケーションのインストールディレクトリを取得"""
    if getattr(sys, 'frozen', False):
        # PyInstallerでビルドされた場合、EXEのあるフォルダ
        return os.path.dirname(sys.executable)
    else:
        # 開発時はスクリプトのあるフォルダ
        return os.path.dirname(os.path.abspath(__file__))


def resource_path(relative_path):
    """リソースファイルのパスを取得（モデル外部化対応）"""
    app_dir = get_app_dir()
    
    # EXEと同じフォルダを優先的に探す
    external_path = os.path.join(app_dir, relative_path)
    if os.path.exists(external_path):
        return external_path
    
    # PyInstallerの一時フォルダ（フォールバック）
    if hasattr(sys, '_MEIPASS'):
        meipass_path = os.path.join(sys._MEIPASS, relative_path)
        if os.path.exists(meipass_path):
            return meipass_path
    
    # 開発時のパス
    return os.path.join(os.path.abspath("."), relative_path)


class AudioTranscriptionApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry("500x560")
        self.root.resizable(True, True)
        
        self.input_file_path = None
        self.output_folder_path = None
        self.is_processing = False
        self.hotwords_list = load_hotwords()

        # ウィンドウアイコンの設定
        self.set_window_icon()

        # モデルの存在確認
        self.check_model()

        self.setup_menu()
        self.setup_ui()
        self.populate_hotwords_listbox()
    
    def set_window_icon(self):
        """ウィンドウアイコンを設定"""
        icon_path = os.path.join(get_app_dir(), APP_ICON_NAME)
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except tk.TclError:
                pass  # アイコン読み込み失敗時は無視

    def check_model(self):
        """モデルフォルダの存在確認"""
        model_dir = resource_path("models")
        if not os.path.exists(model_dir):
            messagebox.showerror(
                "エラー",
                f"モデルフォルダが見つかりません。\n\n"
                f"期待されるパス:\n{model_dir}\n\n"
                f"アプリケーションを再インストールしてください。"
            )
            sys.exit(1)

    def setup_menu(self):
        """メニューバーを作成"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # ヘルプメニュー
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ヘルプ", menu=help_menu)
        help_menu.add_command(label="単語登録機能について", command=self.show_hotwords_help)
        help_menu.add_separator()
        help_menu.add_command(label="ライセンス情報", command=self.show_license_info)

    def show_license_info(self):
        """ライセンス情報ダイアログを表示"""
        dialog = tk.Toplevel(self.root)
        dialog.title("ライセンス情報")
        dialog.geometry("560x450")
        dialog.resizable(True, True)
        dialog.transient(self.root)
        dialog.grab_set()

        # アイコンの設定
        icon_path = os.path.join(get_app_dir(), APP_ICON_NAME)
        if os.path.exists(icon_path):
            try:
                dialog.iconbitmap(icon_path)
            except tk.TclError:
                pass

        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        text_widget = tk.Text(frame, wrap=tk.WORD, font=("", 9))
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.config(yscrollcommand=scrollbar.set)

        text_widget.insert(tk.END, LICENSE_TEXT)
        text_widget.config(state=tk.DISABLED)

        close_btn = ttk.Button(dialog, text="閉じる", command=dialog.destroy)
        close_btn.pack(pady=(0, 10))

    def show_hotwords_help(self):
        """単語登録機能ヘルプダイアログを表示"""
        dialog = tk.Toplevel(self.root)
        dialog.title("単語登録機能について")
        dialog.geometry("500x420")
        dialog.resizable(True, True)
        dialog.transient(self.root)
        dialog.grab_set()

        # アイコンの設定
        icon_path = os.path.join(get_app_dir(), APP_ICON_NAME)
        if os.path.exists(icon_path):
            try:
                dialog.iconbitmap(icon_path)
            except tk.TclError:
                pass

        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        text_widget = tk.Text(frame, wrap=tk.WORD, font=("", 10))
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.config(yscrollcommand=scrollbar.set)

        text_widget.insert(tk.END, HOTWORDS_HELP_TEXT)
        text_widget.config(state=tk.DISABLED)

        close_btn = ttk.Button(dialog, text="閉じる", command=dialog.destroy)
        close_btn.pack(pady=(0, 10))

    def setup_ui(self):
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 入力ファイル選択
        file_frame = ttk.LabelFrame(main_frame, text="入力ファイル", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.file_btn = ttk.Button(file_frame, text="選択", command=self.select_input_file)
        self.file_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        self.file_label = ttk.Label(file_frame, text="ファイルが選択されていません")
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 出力フォルダ選択
        folder_frame = ttk.LabelFrame(main_frame, text="出力フォルダ", padding="10")
        folder_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.folder_btn = ttk.Button(folder_frame, text="選択", command=self.select_output_folder)
        self.folder_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        self.folder_label = ttk.Label(folder_frame, text="フォルダが選択されていません")
        self.folder_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # プログレスバーフレーム
        progress_frame = ttk.LabelFrame(main_frame, text="処理状況", padding="10")
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.status_label = ttk.Label(progress_frame, text="待機中...")
        self.status_label.pack(fill=tk.X, pady=(0, 5))
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', length=400)
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))
        
        self.progress_detail = ttk.Label(progress_frame, text="")
        self.progress_detail.pack(fill=tk.X)
        
        # 単語登録フレーム
        hotwords_frame = ttk.LabelFrame(main_frame, text="単語登録（固有名詞・専門用語）", padding="10")
        hotwords_frame.pack(fill=tk.X, pady=(0, 10))

        # 入力行: テキスト入力 + 追加ボタン
        input_row = ttk.Frame(hotwords_frame)
        input_row.pack(fill=tk.X, pady=(0, 5))

        self.hotword_entry = ttk.Entry(input_row)
        self.hotword_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.hotword_entry.bind("<Return>", lambda e: self.add_hotword())

        self.add_btn = ttk.Button(input_row, text="追加", width=6, command=self.add_hotword)
        self.add_btn.pack(side=tk.LEFT, padx=(0, 5))

        self.remove_btn = ttk.Button(input_row, text="削除", width=6, command=self.remove_hotword)
        self.remove_btn.pack(side=tk.LEFT)

        # 登録数カウンター
        self.hotwords_count_label = ttk.Label(hotwords_frame, text=f"0 / {MAX_HOTWORDS} 件")
        self.hotwords_count_label.pack(anchor=tk.E, pady=(0, 3))

        # 登録済み単語リスト
        list_row = ttk.Frame(hotwords_frame)
        list_row.pack(fill=tk.X)

        self.hotwords_listbox = tk.Listbox(list_row, height=4, selectmode=tk.EXTENDED)
        self.hotwords_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)

        scrollbar = ttk.Scrollbar(list_row, orient=tk.VERTICAL, command=self.hotwords_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.hotwords_listbox.config(yscrollcommand=scrollbar.set)

        # 実行ボタン
        self.run_btn = ttk.Button(main_frame, text="文字起こし開始", command=self.start_processing)
        self.run_btn.pack(pady=10)
    
    def select_input_file(self):
        filetypes = [("音声ファイル", "*.wav;*.mp3;*.m4a;*.mp4")]
        file_path = filedialog.askopenfilename(
            title="音声ファイルを選択", 
            filetypes=filetypes, 
            initialdir='./'
        )
        if file_path:
            self.input_file_path = file_path
            # 長いパスは省略表示
            display_path = file_path if len(file_path) < 50 else "..." + file_path[-47:]
            self.file_label.config(text=display_path)
    
    def select_output_folder(self):
        folder_path = filedialog.askdirectory(title="出力するフォルダを選択", initialdir='./')
        if folder_path:
            self.output_folder_path = folder_path
            display_path = folder_path if len(folder_path) < 50 else "..." + folder_path[-47:]
            self.folder_label.config(text=display_path)
    
    def populate_hotwords_listbox(self):
        """リストボックスに登録済み単語を表示"""
        self.hotwords_listbox.delete(0, tk.END)
        for word in self.hotwords_list:
            self.hotwords_listbox.insert(tk.END, word)
        self.update_hotwords_count()

    def update_hotwords_count(self):
        """登録数カウンターを更新"""
        count = len(self.hotwords_list)
        self.hotwords_count_label.config(text=f"{count} / {MAX_HOTWORDS} 件")

    def add_hotword(self):
        """単語を追加"""
        word = self.hotword_entry.get().strip()
        if not word:
            return
        if len(self.hotwords_list) >= MAX_HOTWORDS:
            messagebox.showwarning(
                "上限",
                f"登録できる単語は最大{MAX_HOTWORDS}件です。\n"
                f"不要な単語を削除してから追加してください。"
            )
            return
        if word in self.hotwords_list:
            messagebox.showinfo("情報", f"「{word}」は既に登録されています。")
            return
        self.hotwords_list.append(word)
        save_hotwords(self.hotwords_list)
        self.hotwords_listbox.insert(tk.END, word)
        self.hotword_entry.delete(0, tk.END)
        self.update_hotwords_count()

    def remove_hotword(self):
        """選択した単語を削除"""
        selected = self.hotwords_listbox.curselection()
        if not selected:
            messagebox.showinfo("情報", "削除する単語を選択してください。")
            return
        # 逆順で削除（インデックスずれ防止）
        for idx in reversed(selected):
            word = self.hotwords_listbox.get(idx)
            self.hotwords_list.remove(word)
            self.hotwords_listbox.delete(idx)
        save_hotwords(self.hotwords_list)
        self.update_hotwords_count()

    def get_hotwords_string(self):
        """登録済み単語をhotwordsパラメータ用の文字列に変換"""
        if not self.hotwords_list:
            return None
        return " ".join(self.hotwords_list)

    def update_progress(self, current, total, status_text, detail_text=""):
        """プログレスバーと状態表示を更新"""
        progress_value = (current / total) * 100 if total > 0 else 0
        self.progress_bar['value'] = progress_value
        self.status_label.config(text=status_text)
        self.progress_detail.config(text=detail_text)
        self.root.update_idletasks()
    
    def set_ui_state(self, enabled):
        """UIの有効/無効を切り替え"""
        state = tk.NORMAL if enabled else tk.DISABLED
        self.file_btn.config(state=state)
        self.folder_btn.config(state=state)
        self.run_btn.config(state=state)
        self.add_btn.config(state=state)
        self.remove_btn.config(state=state)
        self.hotword_entry.config(state=state)
    
    def start_processing(self):
        if not self.input_file_path:
            messagebox.showwarning("警告", "入力ファイルを選択してください。")
            return
        if not self.output_folder_path:
            messagebox.showwarning("警告", "出力フォルダを選択してください。")
            return
        
        # 処理を別スレッドで実行（UIフリーズ防止）
        self.is_processing = True
        self.set_ui_state(False)
        
        thread = threading.Thread(target=self.process_audio)
        thread.daemon = True
        thread.start()
    
    def process_audio(self):
        try:
            self.split_audio_file(self.input_file_path, self.output_folder_path)
            self.root.after(0, lambda: self.on_process_complete())
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("エラー", str(e)))
        finally:
            self.is_processing = False
            self.root.after(0, lambda: self.set_ui_state(True))
            self.root.after(0, lambda: self.update_progress(0, 1, "待機中...", ""))

    def on_process_complete(self):
        """処理完了時の処理"""
        messagebox.showinfo("完了", "処理が完了しました。")
        # 出力フォルダを開く
        if self.output_folder_path:
            os.startfile(self.output_folder_path)

    def split_audio_file(self, file_path, output_folder):
        # 入力ファイルの拡張子を取得
        file_extension = os.path.splitext(file_path)[1].lower()
        file_name = os.path.splitext(os.path.basename(file_path))[0]

        # 進捗更新: モデル読み込み開始
        self.root.after(0, lambda: self.update_progress(0, 100, "モデルを読み込み中...", "初回は時間がかかる場合があります"))

        # ★★★ モデルをループ外で1回だけ読み込む ★★★
        model_dir = resource_path("models")
        model = WhisperModel("large-v3", device="cpu", compute_type="int8", download_root=model_dir)
        
        # 進捗更新: 音声ファイル読み込み
        self.root.after(0, lambda: self.update_progress(5, 100, "音声ファイルを読み込み中..."))

        # 音声ファイルを読み込む
        audio = AudioFileClip(file_path)

        # 分割する時間間隔（１分）を取得
        split_interval = 1 * 60  # 秒単位

        # 分割した音声ファイルを保存するフォルダを作成
        os.makedirs(output_folder, exist_ok=True)

        list1 = ["", "", ""]
        df = pd.DataFrame([list1])
        df.columns = ['No', '音声ファイル', '変換結果']

        # 総分割数を計算
        total_segments = len(range(0, int(audio.duration), split_interval))
        
        # 音声ファイルを分割・文字起こし
        start_time_all = datetime.datetime.now()
        print("処理開始:", start_time_all.strftime("%Y-%m-%d %H:%M:%S"))

        for i, start_time in enumerate(range(0, int(audio.duration), split_interval)):
            # 進捗計算（モデル読み込み5% + 分割処理95%）
            segment_progress = 5 + ((i + 1) / total_segments) * 95
            
            # 進捗更新: 分割処理
            status_text = f"処理中... ({i + 1}/{total_segments})"
            detail_text = f"セグメント {i + 1} を文字起こし中"
            self.root.after(0, lambda st=status_text, dt=detail_text, sp=segment_progress: 
                          self.update_progress(sp, 100, st, dt))

            # 分割開始位置と終了位置を計算
            end_time = min(start_time + split_interval, audio.duration)
            # 音声を分割
            split_audio = audio.subclip(start_time, end_time)
            # 出力ファイル名を作成
            output_file = os.path.join(output_folder, f"{file_name}_{i}{file_extension}")
            # 分割した音声ファイルを保存
            if file_extension == ".wav":
                split_audio.write_audiofile(output_file, codec='pcm_s16le', verbose=False, logger=None)
            elif file_extension == ".mp3":
                split_audio.write_audiofile(output_file, codec='libmp3lame', verbose=False, logger=None)
            elif file_extension == ".m4a":
                split_audio.write_audiofile(output_file, codec='aac', verbose=False, logger=None)
            elif file_extension == ".mp4":
                split_audio.write_audiofile(output_file, codec='aac', verbose=False, logger=None)
            else:
                raise ValueError("サポートされていない音声形式です。")
            print(f"分割ファイル {output_file} を保存しました。")

            # 音声ファイルを文字変換
            print(f"セグメント {i + 1}/{total_segments} 文字起こし開始:", 
                  datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

            # ★★★ モデルは既に読み込み済みなのでここでは使うだけ ★★★
            transcribe_params = dict(
                beam_size=5,
                language='ja',
                temperature=0,
                vad_filter=True,
                initial_prompt="高島宗一郎です。こんにちは、今日はよろしくお願いします。",
            )
            hotwords_str = self.get_hotwords_string()
            if hotwords_str:
                transcribe_params["hotwords"] = hotwords_str
            segments, _ = model.transcribe(output_file, **transcribe_params)

            transcription = ''
            for segment in segments:
                transcription = transcription + str(segment.text) + '\n'
            print(transcription)
            # 結果をdfにセット
            df.loc[i] = [str(i), output_file, transcription]

        # 音声ファイルを閉じる
        audio.close()

        # 進捗更新: Excel出力
        self.root.after(0, lambda: self.update_progress(98, 100, "Excelファイルを出力中..."))

        # excelへ書き出し
        end_time_all = datetime.datetime.now()
        print("処理終了:", end_time_all.strftime("%Y-%m-%d %H:%M:%S"))
        print(f"総処理時間: {end_time_all - start_time_all}")
        
        output_file = os.path.join(output_folder, f"{file_name}_output.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        # DataFrameの値をシートに書き込む
        for r in dataframe_to_rows(df, index=False, header=True):
            sheet.append(r)
        # ファイルへのリンクをセット
        for row in sheet.iter_rows(min_row=2, min_col=2, max_col=2):  # B列の値を処理
            cell = row[0]
            file_path = cell.value

            if file_path:
                cell.hyperlink = file_path
                cell.value = f'{file_path}'
        # Excelファイルを保存
        try:
            workbook.save(output_file)
            print(f"Excelファイル {output_file} を保存しました。")
        except Exception as e:
            print(f"Excelファイルの保存中にエラーが発生しました: {str(e)}")

        # 進捗更新: 完了
        self.root.after(0, lambda: self.update_progress(100, 100, "完了！", 
                       f"総処理時間: {end_time_all - start_time_all}"))

    def run(self):
        self.root.mainloop()


def main():
    app = AudioTranscriptionApp()
    app.run()


if __name__ == "__main__":
    main()
