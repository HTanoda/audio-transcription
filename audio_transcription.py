import os
import re
import sys
import json
import logging

# オフライン専用アプリのため、モデル読込時に一切ネットワークへ出ないよう強制する。
# faster_whisper / huggingface_hub のインポート前に設定する必要がある。
os.environ.setdefault("HF_HUB_OFFLINE", "1")
os.environ.setdefault("TRANSFORMERS_OFFLINE", "1")

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from faster_whisper import WhisperModel
from openpyxl import Workbook
from openpyxl.styles import Alignment
import datetime
import threading
import av

# アプリケーション情報
APP_NAME = "TND_AudioTranscription"
APP_VERSION = "1.4.1"
APP_TITLE = f"TND audio_transcription v{APP_VERSION}"
APP_ICON_NAME = "TND_AudioTranscription01.ico"

# 単語登録キャッシュファイル名
HOTWORDS_FILE = "hotwords.json"
# 単語登録の上限数（hotwords枠223トークンに安全に収まる範囲）
MAX_HOTWORDS = 50

# アプリ設定ファイル名
SETTINGS_FILE = "settings.json"
DEFAULT_MODEL_NAME = "large-v3"

# 句読点を打たせるための呼び水プロンプト（句読点付きの文を与えることで
# モデルが句読点を出力しやすくなる。固有名詞は誤混入を避けるため含めない）
INITIAL_PROMPT = "こんにちは。本日は、よろしくお願いします。それでは、会議を始めます。"

logger = logging.getLogger("audio_transcription")


class ProcessingCancelled(Exception):
    """ユーザー操作により処理がキャンセルされたことを示す例外"""
    pass


def setup_logging():
    """ログ出力の初期設定（logs/app-YYYYMMDD.log に出力）"""
    app_dir = get_app_dir()
    log_dir = os.path.join(app_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, f"app-{datetime.datetime.now().strftime('%Y%m%d')}.log")

    handler = logging.FileHandler(log_file, encoding="utf-8")
    formatter = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    handler.setFormatter(formatter)

    logger.setLevel(logging.INFO)
    logger.addHandler(handler)

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


def get_settings_path():
    """アプリ設定ファイルのパスを取得"""
    app_dir = get_app_dir()
    return os.path.join(app_dir, SETTINGS_FILE)


def load_settings():
    """保存済みのアプリ設定を読み込む"""
    path = get_settings_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
        except (json.JSONDecodeError, IOError):
            pass
    return {}


def save_settings(settings):
    """アプリ設定をJSONファイルに保存する"""
    path = get_settings_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


def export_audio_segment(src_path, dst_path, start_sec, end_sec):
    """元音声から指定区間をストリームコピー（無劣化・再エンコードなし）で切り出す"""
    with av.open(src_path) as in_container:
        in_stream = in_container.streams.audio[0]
        with av.open(dst_path, mode="w") as out_container:
            out_stream = out_container.add_stream_from_template(in_stream)
            try:
                in_container.seek(int(start_sec / in_stream.time_base), stream=in_stream)
            except av.FFmpegError:
                pass
            offset = None
            for packet in in_container.demux(in_stream):
                if packet.pts is None:
                    continue
                t = float(packet.pts * in_stream.time_base)
                if t >= end_sec:
                    break
                if t < start_sec:
                    continue
                if offset is None:
                    offset = packet.pts
                packet.stream = out_stream
                packet.pts -= offset
                if packet.dts is not None:
                    packet.dts -= offset
                out_container.mux(packet)


def detect_models():
    """modelsフォルダ内のHuggingFaceキャッシュ形式フォルダからモデル名を検出する"""
    model_dir = resource_path("models")
    if not os.path.isdir(model_dir):
        return []
    names = []
    for entry in sorted(os.listdir(model_dir)):
        full_path = os.path.join(model_dir, entry)
        if not os.path.isdir(full_path):
            continue
        m = re.match(r"^models--[^-]+(?:-[^-]+)*--faster-whisper-(.+)$", entry)
        if m:
            names.append(m.group(1))
    return names


def resolve_local_model_path(model_name):
    """指定モデルのローカルスナップショットフォルダのパスを返す（無効／不在なら None）。

    HuggingFace の名前解決やネットワークアクセスを一切介さず、同梱済みモデルの
    実ファイルがあるフォルダを直接特定する。model.bin の存在まで確認するため、
    不完全なインストールも None として検出できる。
    """
    model_dir = resource_path("models")
    if not os.path.isdir(model_dir):
        return None
    for entry in sorted(os.listdir(model_dir)):
        full_path = os.path.join(model_dir, entry)
        if not os.path.isdir(full_path):
            continue
        m = re.match(r"^models--[^-]+(?:-[^-]+)*--faster-whisper-(.+)$", entry)
        if not m or m.group(1) != model_name:
            continue
        snapshots = os.path.join(full_path, "snapshots")
        if not os.path.isdir(snapshots):
            return None
        candidates = []
        ref = os.path.join(full_path, "refs", "main")
        if os.path.isfile(ref):
            try:
                with open(ref, "r", encoding="utf-8") as f:
                    candidates.append(os.path.join(snapshots, f.read().strip()))
            except OSError:
                pass
        candidates.extend(
            os.path.join(snapshots, d)
            for d in sorted(os.listdir(snapshots))
            if os.path.isdir(os.path.join(snapshots, d))
        )
        for cand in candidates:
            if os.path.isdir(cand) and os.path.isfile(os.path.join(cand, "model.bin")):
                return cand
        return None
    return None


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
        self.root.geometry("500x680")
        self.root.resizable(True, True)

        self.input_file_paths = []
        self.output_folder_path = None
        self.is_processing = False
        self.cancel_event = threading.Event()
        self.hotwords_list = load_hotwords()
        self.settings = load_settings()

        self.available_models = detect_models()
        self.selected_model_name = self.settings.get("model_name") or (
            self.available_models[0] if self.available_models else DEFAULT_MODEL_NAME
        )
        if self.available_models and self.selected_model_name not in self.available_models:
            self.selected_model_name = self.available_models[0]

        self.output_split_var = tk.BooleanVar(value=bool(self.settings.get("output_split", True)))
        self.output_txt_var = tk.BooleanVar(value=bool(self.settings.get("output_txt", False)))
        self.output_srt_var = tk.BooleanVar(value=bool(self.settings.get("output_srt", False)))

        # ウィンドウアイコンの設定
        self.set_window_icon()

        # モデルの存在確認
        self.check_model()

        self.setup_menu()
        self.setup_ui()
        self.populate_hotwords_listbox()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_close(self):
        """ウィンドウを閉じる際の処理（処理中は確認する）"""
        if self.is_processing:
            if messagebox.askyesno("確認", "処理が実行中です。中断して終了しますか？"):
                self.cancel_event.set()
                self.root.destroy()
        else:
            self.root.destroy()
    
    def set_window_icon(self):
        """ウィンドウアイコンを設定"""
        icon_path = os.path.join(get_app_dir(), APP_ICON_NAME)
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except tk.TclError:
                pass  # アイコン読み込み失敗時は無視

    def check_model(self):
        """モデルの存在確認（フォルダの有無だけでなく、有効なモデル実体まで検証する）"""
        model_dir = resource_path("models")
        if not os.path.exists(model_dir):
            messagebox.showerror(
                "エラー",
                f"モデルフォルダが見つかりません。\n\n"
                f"期待されるパス:\n{model_dir}\n\n"
                f"アプリケーションを再インストールしてください。"
            )
            sys.exit(1)
        # 有効なモデル（model.bin まで揃ったスナップショット）が存在するか確認。
        # 不完全なインストールをここで検知し、実行時のネットワークダウンロード試行を防ぐ。
        if resolve_local_model_path(self.selected_model_name) is None:
            messagebox.showerror(
                "エラー",
                f"利用可能なモデルが見つかりません。\n\n"
                f"モデルフォルダ:\n{model_dir}\n\n"
                f"ZIPを展開せずにセットアップした場合や、モデルファイル "
                f"(約1.5〜3GB) のコピーが完了していない場合に発生します。\n"
                f"ZIPを完全に展開したうえで、アプリケーションを再インストールしてください。"
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

        # モデル選択（検出されたモデルが2つ以上の場合のみ表示）
        if len(self.available_models) >= 2:
            model_frame = ttk.LabelFrame(main_frame, text="モデル", padding="10")
            model_frame.pack(fill=tk.X, pady=(0, 10))

            ttk.Label(model_frame, text="モデル: ").pack(side=tk.LEFT)

            self.model_combo = ttk.Combobox(
                model_frame, state="readonly", values=self.available_models
            )
            self.model_combo.set(self.selected_model_name)
            self.model_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
            self.model_combo.bind("<<ComboboxSelected>>", self.on_model_selected)

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

        # 出力オプション
        output_options_frame = ttk.LabelFrame(main_frame, text="出力オプション", padding="10")
        output_options_frame.pack(fill=tk.X, pady=(0, 10))

        self.output_split_check = ttk.Checkbutton(
            output_options_frame, text="分割音声 (1分ごと・再生用) も出力",
            variable=self.output_split_var, command=self.on_output_option_changed
        )
        self.output_split_check.pack(anchor=tk.W)

        self.output_txt_check = ttk.Checkbutton(
            output_options_frame, text="テキスト (.txt) も出力",
            variable=self.output_txt_var, command=self.on_output_option_changed
        )
        self.output_txt_check.pack(anchor=tk.W)

        self.output_srt_check = ttk.Checkbutton(
            output_options_frame, text="字幕 (.srt) も出力",
            variable=self.output_srt_var, command=self.on_output_option_changed
        )
        self.output_srt_check.pack(anchor=tk.W)

        # 実行ボタン・キャンセルボタン
        button_row = ttk.Frame(main_frame)
        button_row.pack(pady=10)

        self.run_btn = ttk.Button(button_row, text="文字起こし開始", command=self.start_processing)
        self.run_btn.pack(side=tk.LEFT, padx=(0, 5))

        self.cancel_btn = ttk.Button(button_row, text="キャンセル", command=self.cancel_processing, state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT)

    def cancel_processing(self):
        """処理のキャンセルを要求"""
        self.cancel_event.set()
        self.cancel_btn.config(state=tk.DISABLED)
        self.status_label.config(text="キャンセル中...")
    
    def select_input_file(self):
        filetypes = [("音声ファイル", "*.wav;*.mp3;*.m4a;*.mp4")]
        file_paths = filedialog.askopenfilenames(
            title="音声ファイルを選択（複数選択可）",
            filetypes=filetypes,
            initialdir='./'
        )
        if file_paths:
            self.input_file_paths = list(file_paths)
            if len(self.input_file_paths) == 1:
                file_path = self.input_file_paths[0]
                # 長いパスは省略表示
                display_path = file_path if len(file_path) < 50 else "..." + file_path[-47:]
                self.file_label.config(text=display_path)
            else:
                first_name = os.path.basename(self.input_file_paths[0])
                self.file_label.config(
                    text=f"{len(self.input_file_paths)} 件選択: {first_name} ほか"
                )
    
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

    def on_model_selected(self, event=None):
        """モデル選択変更時の処理"""
        self.selected_model_name = self.model_combo.get()
        self.settings["model_name"] = self.selected_model_name
        save_settings(self.settings)

    def on_output_option_changed(self):
        """出力オプション変更時の処理"""
        self.settings["output_split"] = self.output_split_var.get()
        self.settings["output_txt"] = self.output_txt_var.get()
        self.settings["output_srt"] = self.output_srt_var.get()
        save_settings(self.settings)

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
        self.output_split_check.config(state=state)
        self.output_txt_check.config(state=state)
        self.output_srt_check.config(state=state)
        if hasattr(self, "model_combo"):
            self.model_combo.config(state="readonly" if enabled else tk.DISABLED)
        self.cancel_btn.config(state=tk.DISABLED if enabled else tk.NORMAL)
    
    def start_processing(self):
        if not self.input_file_paths:
            messagebox.showwarning("警告", "入力ファイルを選択してください。")
            return
        if not self.output_folder_path:
            messagebox.showwarning("警告", "出力フォルダを選択してください。")
            return

        existing_names = []
        for file_path in self.input_file_paths:
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            output_file = os.path.join(self.output_folder_path, f"{file_name}_output.xlsx")
            if os.path.exists(output_file):
                existing_names.append(f"{file_name}_output.xlsx")
        if existing_names:
            names_text = "\n".join(existing_names)
            if not messagebox.askyesno(
                "確認",
                f"以下の出力ファイルが既に存在します。上書きしますか？\n\n{names_text}"
            ):
                return

        # 処理を別スレッドで実行（UIフリーズ防止）
        self.cancel_event.clear()
        self.is_processing = True
        self.set_ui_state(False)

        thread = threading.Thread(target=self.process_audio)
        thread.daemon = True
        thread.start()

    def process_audio(self):
        file_count = len(self.input_file_paths)
        succeeded = 0
        failed_names = []
        model = None
        try:
            for file_index, file_path in enumerate(self.input_file_paths, start=1):
                if self.cancel_event.is_set():
                    raise ProcessingCancelled()
                file_name = os.path.basename(file_path)
                try:
                    if model is None:
                        self.root.after(0, lambda: self.update_progress(
                            0, 100, "モデルを読み込み中...", "初回は時間がかかる場合があります"))
                        # 同梱モデルのローカルパスを直接指定して読み込む（ネットワークに出ない）。
                        model_path = resolve_local_model_path(self.selected_model_name)
                        if model_path is None:
                            raise RuntimeError(
                                "モデルが見つかりません。アプリケーションを再インストールしてください。"
                            )
                        model = WhisperModel(
                            model_path, device="cpu",
                            compute_type="int8", local_files_only=True
                        )
                    self.transcribe_file(
                        model, file_path, self.output_folder_path,
                        file_index=file_index, file_count=file_count
                    )
                    succeeded += 1
                except ProcessingCancelled:
                    raise
                except Exception:
                    logger.exception(f"文字起こし処理に失敗しました: {file_path}")
                    failed_names.append(file_name)

            self.root.after(0, lambda: self.on_process_complete(file_count, succeeded, failed_names))
        except ProcessingCancelled:
            logger.info("ユーザー操作により処理がキャンセルされました。")
            self.root.after(0, lambda: messagebox.showinfo("キャンセル", "処理をキャンセルしました。"))
        finally:
            self.is_processing = False
            self.root.after(0, lambda: self.set_ui_state(True))
            self.root.after(0, lambda: self.update_progress(0, 1, "待機中...", ""))

    def on_process_complete(self, file_count, succeeded, failed_names):
        """処理完了時の処理"""
        if not failed_names:
            messagebox.showinfo("完了", "処理が完了しました。")
        else:
            names_text = "\n".join(failed_names)
            messagebox.showwarning(
                "完了",
                f"{file_count}件中{succeeded}件成功。\n失敗: {names_text}\n\n"
                f"詳細はログファイル (logs フォルダ) を確認してください。"
            )
        # 出力フォルダを開く
        if self.output_folder_path:
            os.startfile(self.output_folder_path)

    @staticmethod
    def format_time(seconds):
        """秒数を HH:MM:SS 形式の文字列に変換"""
        sec = int(seconds)
        return f"{sec // 3600:02d}:{(sec % 3600) // 60:02d}:{sec % 60:02d}"

    @staticmethod
    def format_srt_time(seconds):
        """秒数を SRT 用の HH:MM:SS,mmm 形式の文字列に変換"""
        total_ms = int(round(seconds * 1000))
        ms = total_ms % 1000
        total_sec = total_ms // 1000
        h = total_sec // 3600
        m = (total_sec % 3600) // 60
        s = total_sec % 60
        return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"

    def transcribe_file(self, model, file_path, output_folder, file_index=1, file_count=1):
        file_extension = os.path.splitext(file_path)[1].lower()
        if file_extension not in (".wav", ".mp3", ".m4a", ".mp4"):
            raise ValueError("サポートされていない音声形式です。")
        file_name = os.path.splitext(os.path.basename(file_path))[0]

        prefix = f"[{file_index}/{file_count}] {os.path.basename(file_path)}: " if file_count > 1 else ""

        os.makedirs(output_folder, exist_ok=True)

        start_time_all = datetime.datetime.now()
        logger.info(f"処理開始: {file_path}")

        transcribe_params = dict(
            beam_size=5,
            language='ja',
            temperature=0,
            vad_filter=True,
            initial_prompt=INITIAL_PROMPT,
        )
        hotwords_str = self.get_hotwords_string()
        if hotwords_str:
            transcribe_params["hotwords"] = hotwords_str

        # ファイル全体を1回で文字起こしし、タイムスタンプで1分単位にまとめる
        # （物理分割しないため文の途中で切れず、再エンコードによる劣化もない）
        self.root.after(0, lambda: self.update_progress(5, 100, f"{prefix}文字起こし中...", ""))
        segments, info = model.transcribe(file_path, **transcribe_params)
        duration = max(info.duration or 0, 1.0)
        logger.info(f"音声長: {duration:.1f}秒")

        split_interval = 60
        buckets = {}
        raw_segments = []
        total_chars = 0
        for segment in segments:
            if self.cancel_event.is_set():
                raise ProcessingCancelled()
            idx = int(segment.start // split_interval)
            text = str(segment.text).strip()
            buckets.setdefault(idx, []).append(text)
            raw_segments.append((segment.start, segment.end, text))
            total_chars += len(text)
            progress = 5 + min(segment.end / duration, 1.0) * 93
            status_text = f"{prefix}文字起こし中... ({self.format_time(segment.end)} / {self.format_time(duration)})"
            self.root.after(0, lambda p=progress, st=status_text:
                            self.update_progress(p, 100, st, ""))

        num_rows = max(buckets.keys()) + 1 if buckets else 1
        end_time_all = datetime.datetime.now()
        logger.info(f"文字起こし完了 ({total_chars}文字) 処理時間: {end_time_all - start_time_all}")

        # 再生用の分割音声を専用フォルダに出力（Excelの各行から該当区間へ頭出しできるようにする）
        row_links = {}
        if self.output_split_var.get():
            self.root.after(0, lambda: self.update_progress(
                95, 100, f"{prefix}再生用の分割音声を出力中..."))
            split_dir = os.path.join(output_folder, f"{file_name}_分割音声")
            os.makedirs(split_dir, exist_ok=True)
            for idx in range(num_rows):
                if self.cancel_event.is_set():
                    raise ProcessingCancelled()
                start_sec = idx * split_interval
                end_sec = min((idx + 1) * split_interval, duration)
                split_path = os.path.join(split_dir, f"{file_name}_{idx}{file_extension}")
                try:
                    export_audio_segment(file_path, split_path, start_sec, end_sec)
                    row_links[idx] = split_path
                except Exception:
                    logger.exception(f"分割音声の出力に失敗しました: {split_path}")
            logger.info(f"分割音声 {len(row_links)}/{num_rows} 件を {split_dir} に出力しました。")

        # 進捗更新: Excel出力
        self.root.after(0, lambda: self.update_progress(98, 100, f"{prefix}Excelファイルを出力中..."))

        output_file = os.path.join(output_folder, f"{file_name}_output.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(['No', '時間帯', '音声ファイル', '変換結果'])
        for idx in range(num_rows):
            start_sec = idx * split_interval
            end_sec = min((idx + 1) * split_interval, duration)
            time_label = f"{self.format_time(start_sec)} - {self.format_time(end_sec)}"
            link_path = row_links.get(idx)
            link_name = os.path.basename(link_path) if link_path else ""
            sheet.append([str(idx), time_label, link_name, '\n'.join(buckets.get(idx, []))])
            # 音声ファイルセルから該当区間の分割音声を開けるようにリンクを付与
            if link_path:
                sheet.cell(row=idx + 2, column=3).hyperlink = link_path
        sheet.column_dimensions['A'].width = 6
        sheet.column_dimensions['B'].width = 22
        sheet.column_dimensions['C'].width = 28
        sheet.column_dimensions['D'].width = 100
        for row in sheet.iter_rows(min_row=2, min_col=4, max_col=4):
            row[0].alignment = Alignment(wrap_text=True, vertical='top')

        try:
            workbook.save(output_file)
        except PermissionError:
            raise RuntimeError(
                f"Excelファイルを保存できません。\n{output_file}\n"
                f"このファイルを開いている場合は閉じてから再実行してください。"
            )
        logger.info(f"Excelファイル {output_file} を保存しました。")

        if self.output_txt_var.get():
            self.write_txt_output(output_folder, file_name, buckets, num_rows, split_interval, duration)
        if self.output_srt_var.get():
            self.write_srt_output(output_folder, file_name, raw_segments)

        # 進捗更新: 完了
        self.root.after(0, lambda: self.update_progress(100, 100, f"{prefix}完了！",
                       f"総処理時間: {end_time_all - start_time_all}"))

    def write_txt_output(self, output_folder, file_name, buckets, num_rows, split_interval, duration):
        """1分ブロック単位のテキストファイルを出力"""
        output_file = os.path.join(output_folder, f"{file_name}_output.txt")
        lines = []
        for idx in range(num_rows):
            start_sec = idx * split_interval
            end_sec = min((idx + 1) * split_interval, duration)
            lines.append(f"[{self.format_time(start_sec)} - {self.format_time(end_sec)}]")
            lines.append('\n'.join(buckets.get(idx, [])))
            lines.append("")
        with open(output_file, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        logger.info(f"テキストファイル {output_file} を保存しました。")

    def write_srt_output(self, output_folder, file_name, raw_segments):
        """セグメント単位のSRT字幕ファイルを出力"""
        output_file = os.path.join(output_folder, f"{file_name}_output.srt")
        lines = []
        for i, (start, end, text) in enumerate(raw_segments, start=1):
            lines.append(str(i))
            lines.append(f"{self.format_srt_time(start)} --> {self.format_srt_time(end)}")
            lines.append(text)
            lines.append("")
        with open(output_file, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        logger.info(f"字幕ファイル {output_file} を保存しました。")

    def run(self):
        self.root.mainloop()


def main():
    try:
        setup_logging()
    except Exception:
        pass
    app = AudioTranscriptionApp()
    app.run()


if __name__ == "__main__":
    main()
