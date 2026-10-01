import os
import re
import sys
import math
import json
import logging

# オフライン専用アプリのため、モデル読込時に一切ネットワークへ出ないよう強制する。
# faster_whisper / huggingface_hub のインポート前に設定する必要がある。
os.environ.setdefault("HF_HUB_OFFLINE", "1")
os.environ.setdefault("TRANSFORMERS_OFFLINE", "1")
# pyannote.audio は既定でテレメトリ（OTLP経由で otel.pyannote.ai へ使用状況を送信）が
# 有効になっている。オフライン専用アプリのため、pyannote.audio のインポート前に
# 明示的に無効化する（PYANNOTE_METRICS_ENABLED は pyannote.audio.telemetry.metrics が
# 参照する環境変数）。
os.environ.setdefault("PYANNOTE_METRICS_ENABLED", "false")

import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES, COPY
from faster_whisper import WhisperModel
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill
import datetime
import threading
import queue
import uuid
import wave
import av
import docx
import winsound
import ctypes
import platform
import numpy as np
import sounddevice as sd
from faster_whisper.audio import decode_audio

# アプリケーション情報
APP_NAME = "TND_AudioTranscription"
APP_VERSION = "1.7.2"
APP_TITLE = f"TND audio_transcription v{APP_VERSION}"
APP_ICON_NAME = "TND_AudioTranscription01.ico"

# 低信頼区間とみなす avg_logprob のしきい値（これ未満はハイライト対象）
LOW_CONFIDENCE_LOGPROB = -0.8

# 単語登録キャッシュファイル名
HOTWORDS_FILE = "hotwords.json"
# 単語登録の上限数（hotwords枠223トークンに安全に収まる範囲）
MAX_HOTWORDS = 50

# アプリ設定ファイル名
SETTINGS_FILE = "settings.json"
# 同梱ライブラリのライセンス全文（インストーラーが EXE と同じフォルダに置く）
THIRD_PARTY_LICENSES_FILE = "THIRD_PARTY_LICENSES.txt"
DEFAULT_MODEL_NAME = "large-v3"

# 句読点を打たせるための呼び水プロンプト（句読点付きの文を与えることで
# モデルが句読点を出力しやすくなる。固有名詞は誤混入を避けるため含めない）
INITIAL_PROMPT = "こんにちは。本日は、よろしくお願いします。それでは、会議を始めます。"

# マイク録音の形式（16kHz / 16bit / モノラル固定）
RECORD_SAMPLE_RATE = 16000
RECORD_CHANNELS = 1
RECORD_DTYPE = "int16"

# ドラッグ＆ドロップで受け付ける音声ファイルの拡張子（ファイル選択ダイアログの絞り込みと同じ）
AUDIO_EXTENSIONS = (".wav", ".mp3", ".m4a", ".mp4")

# 話者分離（同梱モデルがある場合のみ有効化）
DIARIZATION_MODEL_DIR = "models_diarization"
DIARIZATION_MODEL_REPO = "pyannote/speaker-diarization-community-1"
DIARIZATION_SAMPLE_RATE = 16000
DIARIZATION_UNKNOWN_SPEAKER = "不明"
DIARIZATION_NEAREST_THRESHOLD_SEC = 2.0

logger = logging.getLogger("audio_transcription")


class ProcessingCancelled(Exception):
    """ユーザー操作により処理がキャンセルされたことを示す例外"""
    pass


def cleanup_old_logs(log_dir, days=30):
    """30日より古いログファイルを削除する"""
    cutoff = datetime.datetime.now() - datetime.timedelta(days=days)
    for entry in os.listdir(log_dir):
        m = re.match(r"^app-(\d{8})\.log$", entry)
        if not m:
            continue
        try:
            file_date = datetime.datetime.strptime(m.group(1), "%Y%m%d")
        except ValueError:
            continue
        if file_date < cutoff:
            try:
                os.remove(os.path.join(log_dir, entry))
            except OSError:
                pass


def setup_logging():
    """ログ出力の初期設定（logs/app-YYYYMMDD.log に出力）"""
    app_dir = get_app_dir()
    log_dir = os.path.join(app_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    try:
        cleanup_old_logs(log_dir)
    except OSError:
        pass
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

■ python-docx
MIT License  https://github.com/python-openxml/python-docx

■ lxml
BSD 3-Clause License  https://github.com/lxml/lxml

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

■ pyannote.audio（話者分離）
MIT License  Copyright (c) 2020 CNRS
https://github.com/pyannote/pyannote-audio

■ pyannote/speaker-diarization-community-1（話者分離モデル）
CC BY 4.0（帰属: pyannote）
https://huggingface.co/pyannote/speaker-diarization-community-1

■ PyTorch（話者分離の計算基盤・CPU版）
BSD 3-Clause License  Copyright (c) 2016- Facebook, Inc他
https://github.com/pytorch/pytorch

■ torchaudio
BSD 2-Clause License  Copyright (c) 2017 Facebook Inc.(Soumith Chintala)
https://github.com/pytorch/audio

■ python-sounddevice（マイク録音）
MIT License  Copyright (c) 2015-2025 Matthias Geier
https://github.com/spatialaudio/python-sounddevice

■ PortAudio（マイク録音の音声入出力ライブラリ）
MIT License  Copyright (c) 1999-2011 Ross Bencina and Phil Burk
https://www.portaudio.com/

■ CustomTkinter（画面部品）
MIT License  Copyright (c) 2023 Tom Schimansky
https://github.com/TomSchimansky/CustomTkinter

■ darkdetect（ライト/ダーク表示の判定）
BSD 3-Clause License  Copyright (c) 2019, Alberto Sottile
https://github.com/albertosottile/darkdetect

■ packaging
Apache License 2.0 / BSD 2-Clause License（デュアルライセンス、本アプリは BSD 2-Clause で利用）
Copyright (c) Donald Stufft and individual contributors.
https://github.com/pypa/packaging

■ TkinterDnD2（ドラッグ＆ドロップ）
MIT License  Copyright (c) 2020 Philippe Gagné
https://github.com/Eliav2/tkinterdnd2

■ tkDnD（ドラッグ＆ドロップの Tcl/Tk 拡張）
Tcl/Tk 系 BSD スタイルライセンス  Copyright (c) Georgios Petasis
https://github.com/petasis/tkdnd
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


def _spectral_gate(x, n_fft=512, hop=128):
    """1チャンクに対するスペクトルゲーティング（ノイズ抑制）"""
    win = np.hanning(n_fft)
    n_frames = 1 + (len(x) - n_fft) // hop
    idx = np.arange(n_fft)[None, :] + hop * np.arange(n_frames)[:, None]
    frames = x[idx] * win[None, :]
    spec = np.fft.rfft(frames, axis=1)
    mag = np.abs(spec).astype(np.float32)

    # 周波数ビンごとの雑音床を時間方向の下位パーセンタイルで推定（チャンクごとに適応）
    noise = np.percentile(mag, 15, axis=0)

    # ゲイン計算（過減算 1.5 倍、下限 0.1 で自然さを維持）
    gain = 1.0 - 1.5 * noise[None, :] / np.maximum(mag, 1e-10)
    gain = np.clip(gain, 0.1, 1.0)

    # 時間方向・周波数方向に軽く平滑化してミュージカルノイズを抑える
    g = gain
    g = (np.vstack([g[:1], g[:-1]]) + g + np.vstack([g[1:], g[-1:]])) / 3.0
    g = (np.hstack([g[:, :1], g[:, :-1]]) + g + np.hstack([g[:, 1:], g[:, -1:]])) / 3.0

    spec *= g

    rec = np.fft.irfft(spec, n=n_fft, axis=1) * win[None, :]
    y = np.zeros(len(x))
    norm = np.zeros(len(x))
    win_sq = win ** 2
    for i in range(n_frames):
        s = i * hop
        y[s:s + n_fft] += rec[i]
        norm[s:s + n_fft] += win_sq
    y[:n_frames * hop + n_fft] /= np.maximum(norm[:n_frames * hop + n_fft], 1e-8)
    # フレーム化で端数となった末尾はそのまま残す
    tail = n_frames * hop + n_fft - hop
    if tail < len(x):
        y[tail:] = x[tail:]
    return y


def preprocess_low_quality(audio, cancel_event=None, sampling_rate=16000):
    """雑音の多い音声を認識向けに前処理する（ノイズ抑制+音量正規化）。

    16kHz float32 モノラル波形を受け取り、60秒チャンクごとに
    スペクトルゲーティングを適用（メモリ使用を一定に保つ）し、
    最後に全体の音量を正規化した波形を返す。
    """
    x = np.asarray(audio, dtype=np.float64)
    if len(x) < 2048:
        return np.asarray(audio, dtype=np.float32)

    chunk = sampling_rate * 60
    y = np.empty(len(x))
    for s in range(0, len(x), chunk):
        if cancel_event is not None and cancel_event.is_set():
            raise ProcessingCancelled()
        seg = x[s:s + chunk]
        if len(seg) < 2048:
            y[s:s + len(seg)] = seg
        else:
            y[s:s + len(seg)] = _spectral_gate(seg)

    # RMS 正規化（目標 -20dBFS 相当、増幅は最大20倍まで）
    rms = np.sqrt(np.mean(y ** 2))
    if rms > 1e-8:
        y *= min(0.1 / rms, 20.0)
    y = np.clip(y, -1.0, 1.0)
    return y.astype(np.float32)


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


def resolve_diarization_model_path():
    """話者分離モデルの同梱キャッシュフォルダのパスを返す（無効／不在なら None）。

    Pipeline.from_pretrained(..., cache_dir=<この戻り値>) にそのまま渡せる、
    HuggingFace hub キャッシュのルート（models_diarization フォルダ自体）を返す。
    """
    cache_dir = resource_path(DIARIZATION_MODEL_DIR)
    if not os.path.isdir(cache_dir):
        return None
    model_dir = os.path.join(
        cache_dir, "models--" + DIARIZATION_MODEL_REPO.replace("/", "--")
    )
    snapshots = os.path.join(model_dir, "snapshots")
    if not os.path.isdir(snapshots):
        return None
    for entry in sorted(os.listdir(snapshots)):
        snap_path = os.path.join(snapshots, entry)
        if os.path.isdir(snap_path) and os.path.isfile(os.path.join(snap_path, "config.yaml")):
            return cache_dir
    return None


def run_selftest():
    """--selftest: GUIを起動せず、凍結EXEの同梱漏れを検出するための疎通確認を行う。

    1. faster-whisper: detect_models() + resolve_local_model_path() でモデル実体を解決できるか
    2. sounddevice: import + query_devices()（入力デバイス0件でも成功。import/DLLエラーのみ失敗）
    3. pyannote.audio: resolve_diarization_model_path() 解決 + Pipeline.from_pretrained() ロード
       + 5秒のダミー波形で pipeline を実行完走できるか（話者0人でも成功）
    4. tkdnd: 非表示の tk ルートに TkinterDnD.require() で拡張を読み込めるか

    結果は logs/selftest_YYYYMMDD_HHMMSS.log に書き出す。
    戻り値は終了コード（0=全成功 / 1=失敗あり）。
    """
    app_dir = get_app_dir()
    log_dir = os.path.join(app_dir, "logs")
    try:
        os.makedirs(log_dir, exist_ok=True)
    except OSError:
        pass
    log_path = os.path.join(
        log_dir, "selftest_{}.log".format(datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
    )

    results = []

    def record(name, ok, detail=""):
        results.append((name, ok, detail))

    # 1. faster-whisper: ローカルモデル実体の解決
    try:
        model_names = detect_models()
        if not model_names:
            raise RuntimeError("モデルが検出できません（detect_models() が空）")
        target_model = model_names[0]
        model_path = resolve_local_model_path(target_model)
        if not model_path:
            raise RuntimeError(f"モデル実体を解決できません: {target_model}")
        record("faster_whisper", True, f"model={target_model} path={model_path}")
    except Exception as e:
        record("faster_whisper", False, f"{type(e).__name__}: {e}")

    # 2. sounddevice: import + デバイス列挙（デバイス0件は許容、import/DLLエラーのみ失敗）
    try:
        devices = sd.query_devices()
        record("sounddevice", True, f"device_count={len(devices)}")
    except Exception as e:
        record("sounddevice", False, f"{type(e).__name__}: {e}")

    # 3. pyannote.audio: モデル解決 + ロード + ダミー波形での実行完走
    try:
        diarization_path = resolve_diarization_model_path()
        if not diarization_path:
            raise RuntimeError("話者分離モデルが検出できません")
        import torch
        from pyannote.audio import Pipeline
        pipeline = Pipeline.from_pretrained(DIARIZATION_MODEL_REPO, cache_dir=diarization_path)
        dummy_audio = (np.random.randn(DIARIZATION_SAMPLE_RATE * 5) * 0.01).astype(np.float32)
        waveform = torch.from_numpy(dummy_audio).unsqueeze(0)
        audio_input = {"waveform": waveform, "sample_rate": DIARIZATION_SAMPLE_RATE}
        diarization = pipeline(audio_input)
        annotation = getattr(diarization, "speaker_diarization", diarization)
        turn_count = sum(1 for _ in annotation.itertracks(yield_label=True))
        record("pyannote", True, f"path={diarization_path} turns={turn_count}")
    except Exception as e:
        record("pyannote", False, f"{type(e).__name__}: {e}")

    # 4. tkdnd: ドラッグ＆ドロップ拡張の読み込み（ルートは withdraw して表示しない）
    tk_root = None
    try:
        tk_root = tk.Tk()
        tk_root.withdraw()
        tkdnd_version = TkinterDnD.require(tk_root)
        record("tkdnd", True, f"version={tkdnd_version}")
    except Exception as e:
        record("tkdnd", False, f"{type(e).__name__}: {e}")
    finally:
        if tk_root is not None:
            try:
                tk_root.destroy()
            except Exception:
                pass

    all_ok = all(ok for _, ok, _ in results)
    lines = [
        f"selftest {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"APP_VERSION={APP_VERSION}",
        "",
    ]
    for name, ok, detail in results:
        lines.append(f"[{'OK' if ok else 'NG'}] {name}: {detail}")
    lines.append("")
    lines.append("RESULT: {}".format("ALL_OK" if all_ok else "FAILED"))

    try:
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
    except OSError:
        pass

    # --noconsole ビルドではコンソール未接続時に標準出力が使えないことがある
    # （ダブルクリック起動等）。ログファイルへの記録は上で完了済みのため、
    # コンソール出力の失敗で終了コードの返却を妨げないようにする。
    try:
        for line in lines:
            print(line)
    except Exception:
        pass

    return 0 if all_ok else 1


def decode_audio_pyav_16k_mono(path):
    """PyAVで音声を16kHzモノラルfloat32のnumpy配列にデコードする（話者分離の入力用）"""
    resampler = av.AudioResampler(format="fltp", layout="mono", rate=DIARIZATION_SAMPLE_RATE)
    chunks = []
    with av.open(path) as container:
        stream = container.streams.audio[0]
        for frame in container.decode(stream):
            for rf in resampler.resample(frame):
                chunks.append(rf.to_ndarray().reshape(-1))
        for rf in resampler.resample(None):
            chunks.append(rf.to_ndarray().reshape(-1))
    if not chunks:
        raise RuntimeError(f"音声フレームをデコードできませんでした: {path}")
    return np.concatenate(chunks).astype(np.float32)


def assign_speaker_to_segment(seg_start, seg_end, turns):
    """Whisperセグメントに話者を割り当てる（重なり合算最大→近傍2秒以内→不明）"""
    overlaps = {}
    for turn in turns:
        ov = min(seg_end, turn["end"]) - max(seg_start, turn["start"])
        if ov > 0:
            overlaps[turn["speaker"]] = overlaps.get(turn["speaker"], 0.0) + ov
    if overlaps:
        return max(overlaps.items(), key=lambda kv: kv[1])[0]

    center = (seg_start + seg_end) / 2.0
    best_speaker = None
    best_dist = None
    for turn in turns:
        dist = min(abs(center - turn["start"]), abs(center - turn["end"]))
        if best_dist is None or dist < best_dist:
            best_dist = dist
            best_speaker = turn["speaker"]
    if best_dist is not None and best_dist <= DIARIZATION_NEAREST_THRESHOLD_SEC:
        return best_speaker
    return DIARIZATION_UNKNOWN_SPEAKER


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


# ウィンドウ寸法はすべて論理ピクセル（customtkinter が geometry() に DPI 倍率を掛ける）
WINDOW_WIDTH = 920
WINDOW_MIN_WIDTH = 640
WINDOW_MIN_HEIGHT = 420
# タイトルバー + 余白
WINDOW_CHROME_MARGIN = 60

# 色は (ライト, ダーク) の組
PLACEHOLDER_FOREGROUND = ("gray45", "gray60")
NOTE_TEXT_COLOR = ("gray40", "gray60")
CARD_FG_COLOR = ("gray94", "gray16")
LIST_FG_COLOR = ("gray99", "gray20")
SEPARATOR_COLOR = ("gray80", "gray30")
OUTLINE_BUTTON_STYLE = dict(
    fg_color="transparent",
    border_width=1,
    border_color=("gray70", "gray35"),
    text_color=("gray10", "gray90"),
)
UI_FONT_FAMILY = "Yu Gothic UI"
# モデル数がこれ以下ならセグメント、超えたらドロップダウンで選ばせる
MODEL_SEGMENTED_MAX = 3
# CTkSwitch の文字の開始位置（スイッチ本体 36 + 間隔 6）。補足文をここに揃える
SWITCH_TEXT_OFFSET = 42
INPUT_PLACEHOLDER_TEXT = "ファイルをここにドロップ、または選択"
OUTPUT_FOLDER_PLACEHOLDER_TEXT = "未指定（音声ファイルと同じフォルダに出力）"
# 一覧 1 行の高さ（行 24 + 上下の間隔 1 + 1）と、選択ファイル一覧をスクロールなしで見せる最大行数
LIST_ROW_HEIGHT = 26
INPUT_LIST_MAX_ROWS = 4
HOTWORDS_EXPORT_FILE_NAME = "単語登録.txt"


def compute_window_width(work_area_width: int) -> int:
    """作業領域の幅に収まるウィンドウ幅（初期幅と最小幅の両方に使う）"""
    return max(WINDOW_MIN_WIDTH, min(WINDOW_WIDTH, work_area_width - WINDOW_CHROME_MARGIN))


def compute_initial_geometry(work_area_width: int, work_area_height: int, content_height: int) -> str:
    """内容の必要高さと作業領域の大きさから初期ウィンドウサイズ (WxH) を決める"""
    width = compute_window_width(work_area_width)
    height = max(WINDOW_MIN_HEIGHT, min(content_height, work_area_height - WINDOW_CHROME_MARGIN))
    return f"{width}x{height}"


def get_work_area():
    """Windows の作業領域 (タスクバーを除く) の (幅, 高さ) を物理ピクセルで返す。取得できなければ None"""
    if sys.platform != "win32":
        return None
    try:
        class RECT(ctypes.Structure):
            _fields_ = [
                ("left", ctypes.c_long),
                ("top", ctypes.c_long),
                ("right", ctypes.c_long),
                ("bottom", ctypes.c_long),
            ]

        rect = RECT()
        SPI_GETWORKAREA = 0x0030
        if not ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(rect), 0):
            return None
        width = rect.right - rect.left
        height = rect.bottom - rect.top
        return (width, height) if width > 0 and height > 0 else None
    except Exception:
        logger.warning("作業領域の取得に失敗しました", exc_info=True)
        return None


def get_window_and_monitor_work_rect(hwnd):
    """窓の外枠と、窓のあるモニターの作業領域を ((左, 上, 右, 下), (左, 上, 右, 下)) の物理ピクセルで返す。取得できなければ None"""
    if sys.platform != "win32":
        return None
    try:
        from ctypes import wintypes

        class MONITORINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize", wintypes.DWORD),
                ("rcMonitor", wintypes.RECT),
                ("rcWork", wintypes.RECT),
                ("dwFlags", wintypes.DWORD),
            ]

        user32 = ctypes.WinDLL("user32")
        user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
        user32.GetWindowRect.restype = wintypes.BOOL
        user32.MonitorFromWindow.argtypes = [wintypes.HWND, wintypes.DWORD]
        user32.MonitorFromWindow.restype = wintypes.HMONITOR
        user32.GetMonitorInfoW.argtypes = [wintypes.HMONITOR, ctypes.POINTER(MONITORINFO)]
        user32.GetMonitorInfoW.restype = wintypes.BOOL

        MONITOR_DEFAULTTONEAREST = 2
        window_rect = wintypes.RECT()
        if not user32.GetWindowRect(hwnd, ctypes.byref(window_rect)):
            return None
        monitor = user32.MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
        info = MONITORINFO()
        info.cbSize = ctypes.sizeof(MONITORINFO)
        if not monitor or not user32.GetMonitorInfoW(monitor, ctypes.byref(info)):
            return None
        work = info.rcWork
        return ((window_rect.left, window_rect.top, window_rect.right, window_rect.bottom),
                (work.left, work.top, work.right, work.bottom))
    except Exception:
        logger.warning("窓・モニターの位置の取得に失敗しました", exc_info=True)
        return None


FOLDERID_DESKTOP = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
KF_FLAG_DEFAULT = 0


def get_known_folder(folder_id_guid):
    """Windows の既知フォルダの実パスを返す（OneDrive へのリダイレクトも反映される）。取得できなければ None"""
    if sys.platform != "win32":
        return None
    try:
        from ctypes import wintypes

        class GUID(ctypes.Structure):
            _fields_ = [
                ("Data1", wintypes.DWORD),
                ("Data2", wintypes.WORD),
                ("Data3", wintypes.WORD),
                ("Data4", ctypes.c_ubyte * 8),
            ]

        ole32 = ctypes.WinDLL("ole32")
        shell32 = ctypes.WinDLL("shell32")
        ole32.CLSIDFromString.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(GUID)]
        ole32.CLSIDFromString.restype = ctypes.c_long
        ole32.CoTaskMemFree.argtypes = [ctypes.c_void_p]
        ole32.CoTaskMemFree.restype = None
        shell32.SHGetKnownFolderPath.argtypes = [
            ctypes.POINTER(GUID), wintypes.DWORD, wintypes.HANDLE, ctypes.POINTER(ctypes.c_void_p)]
        shell32.SHGetKnownFolderPath.restype = ctypes.c_long

        guid = GUID()
        if ole32.CLSIDFromString(folder_id_guid, ctypes.byref(guid)) != 0:
            return None
        path_ptr = ctypes.c_void_p()
        hr = shell32.SHGetKnownFolderPath(ctypes.byref(guid), KF_FLAG_DEFAULT, None, ctypes.byref(path_ptr))
        try:
            if hr != 0 or not path_ptr.value:
                return None
            return ctypes.wstring_at(path_ptr.value)
        finally:
            if path_ptr.value:
                ole32.CoTaskMemFree(path_ptr.value)
    except Exception:
        logger.warning("既知フォルダ %s の取得に失敗しました", folder_id_guid, exc_info=True)
        return None


class TaskbarProgress:
    """Windows タスクバーのボタンに進捗を表示する (ITaskbarList3)。

    UI スレッドからだけ呼ぶこと。失敗したら一度だけ警告を記録し、以後は何もしない（本体の処理に影響させない）。
    各メソッドは呼び出した API の HRESULT を返す（無効化後・未実行時は None）。
    """

    CLSID_TASKBAR_LIST = "{56FDF344-FD6D-11d0-958A-006097C9A090}"
    IID_ITASKBAR_LIST3 = "{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"
    CLSCTX_INPROC_SERVER = 0x1
    COINIT_APARTMENTTHREADED = 0x2
    RPC_E_CHANGED_MODE = -2147417850  # 0x80010106
    # vtable の位置: IUnknown 3 + ITaskbarList 5 + ITaskbarList2 1 の後に ITaskbarList3
    VTBL_HR_INIT = 3
    VTBL_SET_PROGRESS_VALUE = 9
    VTBL_SET_PROGRESS_STATE = 10
    TBPF_NOPROGRESS = 0
    TBPF_NORMAL = 2
    TBPF_ERROR = 4
    PROGRESS_SCALE = 1000

    def __init__(self, hwnd_getter):
        self._hwnd_getter = hwnd_getter
        self._ptr = None
        self._hwnd = None
        self._state = None
        self._disabled = sys.platform != "win32"
        self._warned = False

    def _fail(self, what, hr=None):
        self._disabled = True
        if self._warned:
            return
        self._warned = True
        if hr is None:
            logger.warning("タスクバーの進捗表示に失敗しました (%s)。以後は表示しません", what, exc_info=True)
        else:
            logger.warning("タスクバーの進捗表示に失敗しました (%s, HRESULT=0x%08X)。以後は表示しません",
                           what, hr & 0xFFFFFFFF)

    def _ensure(self):
        if self._disabled:
            return False
        if self._ptr is not None:
            return True
        try:
            from ctypes import wintypes

            class GUID(ctypes.Structure):
                _fields_ = [
                    ("Data1", wintypes.DWORD),
                    ("Data2", wintypes.WORD),
                    ("Data3", wintypes.WORD),
                    ("Data4", ctypes.c_ubyte * 8),
                ]

            ole32 = ctypes.WinDLL("ole32")
            ole32.CoInitializeEx.argtypes = [ctypes.c_void_p, wintypes.DWORD]
            ole32.CoInitializeEx.restype = ctypes.c_long
            ole32.CLSIDFromString.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(GUID)]
            ole32.CLSIDFromString.restype = ctypes.c_long
            ole32.CoCreateInstance.argtypes = [
                ctypes.POINTER(GUID), ctypes.c_void_p, wintypes.DWORD,
                ctypes.POINTER(GUID), ctypes.POINTER(ctypes.c_void_p)]
            ole32.CoCreateInstance.restype = ctypes.c_long

            # Tk のスレッドは STA。既に別の方式で初期化済み (RPC_E_CHANGED_MODE) でもそのまま使える
            hr = ole32.CoInitializeEx(None, self.COINIT_APARTMENTTHREADED)
            if hr < 0 and hr != self.RPC_E_CHANGED_MODE:
                self._fail("CoInitializeEx", hr)
                return False

            clsid = GUID()
            iid = GUID()
            hr = ole32.CLSIDFromString(self.CLSID_TASKBAR_LIST, ctypes.byref(clsid))
            if hr < 0:
                self._fail("CLSIDFromString", hr)
                return False
            hr = ole32.CLSIDFromString(self.IID_ITASKBAR_LIST3, ctypes.byref(iid))
            if hr < 0:
                self._fail("CLSIDFromString", hr)
                return False
            ptr = ctypes.c_void_p()
            hr = ole32.CoCreateInstance(
                ctypes.byref(clsid), None, self.CLSCTX_INPROC_SERVER, ctypes.byref(iid), ctypes.byref(ptr))
            if hr < 0 or not ptr.value:
                self._fail("CoCreateInstance", hr)
                return False
            self._ptr = ptr

            hr = self._call(self.VTBL_HR_INIT, ())
            if hr < 0:
                self._ptr = None
                self._fail("HrInit", hr)
                return False

            user32 = ctypes.WinDLL("user32")
            user32.GetParent.argtypes = [wintypes.HWND]
            user32.GetParent.restype = wintypes.HWND
            hwnd = user32.GetParent(self._hwnd_getter())
            if not hwnd:
                self._ptr = None
                self._fail("GetParent", 0)
                return False
            self._hwnd = hwnd
            return True
        except Exception:
            self._ptr = None
            self._fail("初期化")
            return False

    def _call(self, index, argtypes, *args):
        vtbl = ctypes.cast(self._ptr, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
        prototype = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, *argtypes)
        return prototype(vtbl[index])(self._ptr, *args)

    def _set_state(self, state):
        if self._state == state:
            return 0
        hr = self._call(self.VTBL_SET_PROGRESS_STATE, (ctypes.c_void_p, ctypes.c_int), self._hwnd, state)
        if hr < 0:
            self._fail("SetProgressState", hr)
            return hr
        self._state = state
        return hr

    def set_progress(self, current, total):
        if total <= 0 or not self._ensure():
            return None
        try:
            hr = self._set_state(self.TBPF_NORMAL)
            if hr < 0:
                return hr
            completed = int(round(min(max(current / total, 0.0), 1.0) * self.PROGRESS_SCALE))
            hr = self._call(
                self.VTBL_SET_PROGRESS_VALUE, (ctypes.c_void_p, ctypes.c_ulonglong, ctypes.c_ulonglong),
                self._hwnd, completed, self.PROGRESS_SCALE)
            if hr < 0:
                self._fail("SetProgressValue", hr)
            return hr
        except Exception:
            self._fail("SetProgressValue")
            return None

    def set_error(self):
        if not self._ensure():
            return None
        try:
            return self._set_state(self.TBPF_ERROR)
        except Exception:
            self._fail("SetProgressState")
            return None

    def clear(self):
        if self._disabled or self._ptr is None:
            return None
        try:
            return self._set_state(self.TBPF_NOPROGRESS)
        except Exception:
            self._fail("SetProgressState")
            return None


class DnDRoot(ctk.CTk, TkinterDnD.DnDWrapper):
    """ドラッグ＆ドロップを受け付けられる CTk ルート（tkdnd を読めない環境では D&D なしで動く）"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        try:
            self.TkdndVersion = TkinterDnD.require(self)
        except Exception:
            self.TkdndVersion = None
            logger.warning("ドラッグ＆ドロップ機能 (tkdnd) を読み込めませんでした。ファイル選択ボタンのみ使用できます", exc_info=True)


class AudioTranscriptionApp:
    def __init__(self):
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
        self.root = DnDRoot()
        self.root.title(APP_TITLE)
        self.root.resizable(True, True)

        self.input_file_paths = []
        self.output_folder_path = None
        self.is_processing = False
        self.cancel_event = threading.Event()
        self.hotwords_list = load_hotwords()
        self.settings = load_settings()

        # マイク録音用の状態
        self.is_recording = False
        self.record_stream = None
        self.record_queue = None
        self.record_writer_thread = None
        self.record_wave_file = None
        self.record_output_path = None
        self.record_start_dt = None
        self.record_level = 0
        self.record_error = None
        self.record_poll_id = None

        # 開いているライセンス情報・ヘルプのダイアログ（二重に開かないため）
        self._text_dialog = None

        self.available_models = detect_models()
        self.selected_model_name = self.settings.get("model_name") or (
            self.available_models[0] if self.available_models else DEFAULT_MODEL_NAME
        )
        if self.available_models and self.selected_model_name not in self.available_models:
            self.selected_model_name = self.available_models[0]

        self.output_split_var = tk.BooleanVar(value=bool(self.settings.get("output_split", True)))
        self.output_txt_var = tk.BooleanVar(value=bool(self.settings.get("output_txt", False)))
        self.output_srt_var = tk.BooleanVar(value=bool(self.settings.get("output_srt", False)))
        self.output_docx_var = tk.BooleanVar(value=bool(self.settings.get("output_docx", False)))
        self.low_quality_var = tk.BooleanVar(value=bool(self.settings.get("low_quality_mode", False)))

        self.diarization_model_path = resolve_diarization_model_path()
        self.diarization_var = tk.BooleanVar(value=bool(self.settings.get("diarization", False)))

        # ウィンドウアイコンの設定
        self.set_window_icon()

        # モデルの存在確認
        self.check_model()

        self.setup_ui()
        self.setup_drop_target()
        self.populate_hotwords_listbox()
        self.apply_initial_geometry()
        self.taskbar = TaskbarProgress(self.root.winfo_id)
        # 処理中のファイル番号と総数（タスクバーの進捗をバッチ全体で出すため。UI スレッドだけが読み書きする）
        self._batch_pos = (1, 1)

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def apply_initial_geometry(self):
        """内容が収まる大きさ（作業領域を超えない範囲）でウィンドウを開き、最小サイズも作業領域に合わせる。

        winfo_reqheight() と作業領域は物理ピクセルなので、DPI 倍率で割って論理ピクセルにしてから渡す。
        """
        self.root.update_idletasks()
        scaling = ctk.ScalingTracker.get_window_scaling(self.root)
        content_px = (
            self.scroll_content.winfo_reqheight()
            + self.bottom_frame.winfo_reqheight()
            + self.bottom_separator.winfo_reqheight()
        )
        # 切り捨てると 1px はみ出してスクロールバーが出るため切り上げる
        content_height = math.ceil(content_px / scaling)
        work_area_px = get_work_area()
        if work_area_px is not None:
            work_area_width = int(work_area_px[0] / scaling)
            work_area_height = int(work_area_px[1] / scaling)
        else:
            # Tk は customtkinter がプロセスを DPI 対応にする前に画面サイズを取得するため、
            # winfo_screen* は既に論理ピクセル（125% の実測: 1728 = 2160 / 1.25）。倍率で割らない
            work_area_width = self.root.winfo_screenwidth()
            work_area_height = self.root.winfo_screenheight()
        # CTk.geometry() は最小サイズで幅・高さを丸めるため、最小サイズを先に設定する
        self.root.minsize(compute_window_width(work_area_width), WINDOW_MIN_HEIGHT)
        self.root.geometry(compute_initial_geometry(work_area_width, work_area_height, content_height))

    def on_close(self):
        """ウィンドウを閉じる際の処理（処理中・録音中は確認する）"""
        if self.is_recording:
            if not messagebox.askyesno("確認", "録音中です。停止して終了しますか？"):
                return
            self.stop_recording(on_close=True)
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

    def show_license_info(self):
        """ライセンス情報ダイアログを表示"""
        full_text_exists = os.path.exists(os.path.join(get_app_dir(), THIRD_PARTY_LICENSES_FILE))
        self._open_text_dialog(
            "ライセンス情報", "560x450", LICENSE_TEXT, font_size=12,
            extra_button=("全文を開く", self.open_third_party_licenses, full_text_exists))

    def show_hotwords_help(self):
        """単語登録機能ヘルプダイアログを表示"""
        self._open_text_dialog("単語登録機能について", "500x420", HOTWORDS_HELP_TEXT, font_size=13)

    def open_third_party_licenses(self):
        """同梱ライブラリのライセンス全文 (THIRD_PARTY_LICENSES.txt) を既定のアプリで開く"""
        path = os.path.join(get_app_dir(), THIRD_PARTY_LICENSES_FILE)
        try:
            os.startfile(path)
        except OSError:
            logger.exception(f"ライセンス全文を開けませんでした: {path}")
            messagebox.showerror(
                "エラー",
                f"ファイルを開けませんでした。\n{path}\n\n"
                f"詳細はログ (logs/app-YYYYMMDD.log) を確認してください。",
                parent=self._text_dialog or self.root,
            )

    def _open_text_dialog(self, title, geometry, text, font_size, extra_button=None):
        """読み取り専用の文章と「閉じる」ボタンのモーダルダイアログを開く（開いていれば前面に出すだけ）。

        extra_button に (文言, コマンド, 有効か) を渡すと「閉じる」の左に副ボタンを置く。
        """
        if self._text_dialog is not None and self._text_dialog.winfo_exists():
            self._text_dialog.lift()
            self._text_dialog.focus_set()
            return

        dialog = ctk.CTkToplevel(self.root)
        self._text_dialog = dialog

        def close():
            self._text_dialog = None
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", close)
        dialog.title(title)
        dialog.geometry(geometry)
        dialog.resizable(True, True)
        dialog.transient(self.root)

        # アイコンの設定（設定しないと customtkinter 既定のアイコンに差し替えられる）
        icon_path = os.path.join(get_app_dir(), APP_ICON_NAME)
        if os.path.exists(icon_path):
            try:
                dialog.iconbitmap(icon_path)
            except tk.TclError:
                pass

        button_row = ctk.CTkFrame(dialog, fg_color="transparent", corner_radius=0)
        button_row.pack(side=tk.BOTTOM, pady=(0, 16))
        if extra_button is not None:
            extra_text, extra_command, extra_enabled = extra_button
            ctk.CTkButton(
                button_row, text=extra_text, font=self.font_body, command=extra_command,
                state="normal" if extra_enabled else "disabled", **OUTLINE_BUTTON_STYLE,
            ).pack(side=tk.LEFT, padx=(0, 8))
        close_btn = ctk.CTkButton(button_row, text="閉じる", font=self.font_body, command=close)
        close_btn.pack(side=tk.LEFT)

        text_widget = ctk.CTkTextbox(
            dialog, wrap="word", font=ctk.CTkFont(family=UI_FONT_FAMILY, size=font_size))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=16, pady=(16, 12))
        text_widget.insert("1.0", text)
        text_widget.configure(state="disabled")

        # CTkToplevel は生成直後にいったん非表示になるため、表示されてからモーダルにする
        def make_modal():
            if not dialog.winfo_exists():
                return
            try:
                dialog.grab_set()
                dialog.focus_set()
            except tk.TclError:
                dialog.after(50, make_modal)

        dialog.after(50, make_modal)

    def setup_ui(self):
        self.font_heading = ctk.CTkFont(family=UI_FONT_FAMILY, size=15, weight="bold")
        self.font_body = ctk.CTkFont(family=UI_FONT_FAMILY, size=13)
        self.font_note = ctk.CTkFont(family=UI_FONT_FAMILY, size=11)
        self.font_small = ctk.CTkFont(family=UI_FONT_FAMILY, size=12)
        self.font_primary = ctk.CTkFont(family=UI_FONT_FAMILY, size=14, weight="bold")
        root_fg_color = ctk.ThemeManager.theme["CTk"]["fg_color"]

        # 処理状況 + 操作（画面が低くても見えるよう、スクロール領域の外の最下部に固定）
        self.bottom_frame = ctk.CTkFrame(self.root, fg_color="transparent", corner_radius=0)
        self.bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)

        status_row = ctk.CTkFrame(self.bottom_frame, fg_color="transparent", corner_radius=0)
        status_row.pack(fill=tk.X, padx=20, pady=(8, 2))

        self.progress_detail = ctk.CTkLabel(
            status_row, text="", font=self.font_note, text_color=NOTE_TEXT_COLOR)
        self.progress_detail.pack(side=tk.RIGHT)

        self.status_label = ctk.CTkLabel(status_row, text="待機中...", font=self.font_body, anchor="w")
        self.status_label.pack(side=tk.LEFT)

        action_row = ctk.CTkFrame(self.bottom_frame, fg_color="transparent", corner_radius=0)
        action_row.pack(fill=tk.X, padx=20, pady=(0, 8))

        # ヘルプ（メニューバーの代わりに文字リンクを並べる）
        link_row = ctk.CTkFrame(self.bottom_frame, fg_color="transparent", corner_radius=0)
        # ボタン内側の余白 (約 6) の分だけ左を詰め、文字の左端を上の行に揃える
        link_row.pack(fill=tk.X, padx=(14, 20), pady=(0, 10))
        for i, (text, command) in enumerate((
            ("単語登録について", self.show_hotwords_help),
            ("サポート情報をコピー", self.copy_support_info),
            ("ライセンス情報", self.show_license_info),
        )):
            ctk.CTkButton(
                link_row, text=text, command=command, font=self.font_note,
                fg_color="transparent", hover_color=CARD_FG_COLOR, text_color=NOTE_TEXT_COLOR,
                height=22, width=0,
            ).pack(side=tk.LEFT, padx=(0 if i == 0 else 12, 0))

        self.progress_bar = ctk.CTkProgressBar(action_row)
        self._set_bar(self.progress_bar, 0)

        self.run_btn = ctk.CTkButton(
            action_row, text="文字起こし開始", height=38, font=self.font_primary,
            command=self.start_processing)

        self.cancel_btn = ctk.CTkButton(
            action_row, text="キャンセル", height=38, width=110, font=self.font_body,
            command=self.cancel_processing, state="disabled", **OUTLINE_BUTTON_STYLE)

        self.cancel_btn.pack(side=tk.RIGHT)
        self.run_btn.pack(side=tk.RIGHT, padx=(0, 8))
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 16))

        # 高さ 1 では角丸矩形が描かれないため、背景色 (bg_color) 側で線を塗る
        self.bottom_separator = ctk.CTkFrame(
            self.root, height=1, corner_radius=0, fg_color=SEPARATOR_COLOR, bg_color=SEPARATOR_COLOR)
        self.bottom_separator.pack(side=tk.BOTTOM, fill=tk.X)

        # スクロール領域
        self.scroll_canvas = tk.Canvas(
            self.root, highlightthickness=0, bd=0,
            bg=self.root._apply_appearance_mode(root_fg_color))
        self.scroll_bar = ctk.CTkScrollbar(self.root, command=self.scroll_canvas.yview)
        self.scroll_canvas.configure(yscrollcommand=self.scroll_bar.set)
        self.scroll_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        # Tab キーの移動順を「上の内容 → 最下段の操作」にする（移動順は兄弟ウィジェットの重なり順に従う）
        self.bottom_frame.lift()
        # tk.Canvas はライト/ダークの切り替えに追従しないため、背景色を自前で塗り直す
        ctk.AppearanceModeTracker.add(self._on_appearance_mode_changed, self.root)

        # メインフレーム
        # 親が tk.Canvas だと背景色が起動時の色で固定されるため、bg_color も明示してライト/ダークに追従させる
        main_frame = ctk.CTkFrame(
            self.scroll_canvas, corner_radius=0, fg_color=root_fg_color, bg_color=root_fg_color)
        self.scroll_window_id = self.scroll_canvas.create_window((0, 0), window=main_frame, anchor='nw')
        self.scroll_content = main_frame

        # CTkFrame.bind は内部の描画用 Canvas に付くため、フレーム自身の <Configure> に追加で付ける
        tk.Frame.bind(main_frame, "<Configure>", self.on_scroll_content_configure, "+")
        self.scroll_canvas.bind("<Configure>", self.on_scroll_canvas_configure)
        self.root.bind_all("<MouseWheel>", self.on_mouse_wheel, "+")

        # 2 列: 左 = 入力と出力、右 = 調整
        main_frame.grid_columnconfigure(0, weight=1, uniform="col")
        main_frame.grid_columnconfigure(1, weight=1, uniform="col")
        # 両列を同じ高さに引き伸ばし、余りは各列の伸縮するカードが吸収する
        main_frame.grid_rowconfigure(0, weight=1)
        left_col = ctk.CTkFrame(main_frame, fg_color="transparent", corner_radius=0)
        left_col.grid(row=0, column=0, sticky="nsew", padx=(20, 6), pady=20)
        right_col = ctk.CTkFrame(main_frame, fg_color="transparent", corner_radius=0)
        right_col.grid(row=0, column=1, sticky="nsew", padx=(6, 20), pady=20)

        # 音声の入力（ファイル選択とマイク録音）
        input_card, _ = self._make_card(left_col, "音声の入力")
        self.input_card = input_card

        file_row = ctk.CTkFrame(input_card, fg_color="transparent", corner_radius=0)
        file_row.pack(fill=tk.X, padx=16)

        self.file_btn = ctk.CTkButton(
            file_row, text="ファイルを選択...", width=124, font=self.font_body,
            command=self.select_input_file, **OUTLINE_BUTTON_STYLE)
        self.file_btn.pack(side=tk.RIGHT, padx=(8, 0))

        self.clear_input_btn = ctk.CTkButton(
            file_row, text="クリア", width=0, font=self.font_body, state="disabled",
            command=lambda: self._set_input_files([]), **OUTLINE_BUTTON_STYLE)
        self.clear_input_btn.pack(side=tk.RIGHT, padx=(10, 0))

        self.file_label = ctk.CTkLabel(file_row, font=self.font_body, anchor="w")
        self._set_path_label(self.file_label, INPUT_PLACEHOLDER_TEXT, placeholder=True)
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.file_row = file_row

        # 2 件以上選んだときだけ表示する選択ファイルの一覧（1 件以下では pack_forget して高さを戻す）
        self.input_list_frame = ctk.CTkScrollableFrame(
            input_card, height=LIST_ROW_HEIGHT * 2, corner_radius=8, fg_color=LIST_FG_COLOR)
        input_list_scrollbar = getattr(self.input_list_frame, "_scrollbar", None)
        if input_list_scrollbar is not None:
            input_list_scrollbar.configure(height=20)
        else:
            logger.warning("選択ファイル一覧のスクロールバーを取得できませんでした（customtkinter の内部構造が変わった可能性）")
        self.input_list_visible = False
        self.input_file_rows = []

        record_btn_row = ctk.CTkFrame(input_card, fg_color="transparent", corner_radius=0)
        record_btn_row.pack(fill=tk.X, padx=16, pady=(10, 0))

        self.record_start_btn = ctk.CTkButton(
            record_btn_row, text="録音開始", width=88, font=self.font_body,
            command=self.start_recording, **OUTLINE_BUTTON_STYLE)
        self.record_start_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.record_stop_btn = ctk.CTkButton(
            record_btn_row, text="停止", width=64, font=self.font_body,
            command=self.stop_recording, state="disabled", **OUTLINE_BUTTON_STYLE)
        self.record_stop_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.record_time_label = ctk.CTkLabel(record_btn_row, text="00:00", font=self.font_body)
        self.record_time_label.pack(side=tk.LEFT, padx=(8, 0))

        self.record_level_bar = ctk.CTkProgressBar(input_card, height=6)
        self._set_bar(self.record_level_bar, 0)
        self.record_level_bar.pack(fill=tk.X, padx=16, pady=(12, 16))

        # 出力フォルダ選択
        folder_card, _ = self._make_card(left_col, "出力フォルダ")

        folder_row = ctk.CTkFrame(folder_card, fg_color="transparent", corner_radius=0)
        folder_row.pack(fill=tk.X, padx=16, pady=(0, 16))

        self.folder_btn = ctk.CTkButton(
            folder_row, text="フォルダを選択...", width=124, font=self.font_body,
            command=self.select_output_folder, **OUTLINE_BUTTON_STYLE)
        self.folder_btn.pack(side=tk.RIGHT, padx=(10, 0))

        self.folder_label = ctk.CTkLabel(folder_row, font=self.font_body, anchor="w")
        self._set_path_label(self.folder_label, OUTPUT_FOLDER_PLACEHOLDER_TEXT, placeholder=True)
        self.folder_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # モデル選択（検出されたモデルが2つ以上の場合のみ表示）
        if len(self.available_models) >= 2:
            model_card, _ = self._make_card(left_col, "モデル")
            if len(self.available_models) <= MODEL_SEGMENTED_MAX:
                self.model_combo = ctk.CTkSegmentedButton(
                    model_card, values=self.available_models, font=self.font_body,
                    command=self.on_model_selected)
            else:
                self.model_combo = ctk.CTkOptionMenu(
                    model_card, values=self.available_models, font=self.font_body,
                    dropdown_font=self.font_body, command=self.on_model_selected)
            self.model_combo.set(self.selected_model_name)
            self.model_combo.pack(fill=tk.X, padx=16, pady=(0, 16))

        # 認識オプション
        # 窓を内容より高くしたとき、左列の下端を右列（単語登録が伸びる）とそろえるため最後のカードも伸ばす
        recog_card, _ = self._make_card(left_col, "認識オプション", expand=True)

        self.low_quality_check = ctk.CTkSwitch(
            recog_card, text="低品質音源モード（ノイズ抑制）", font=self.font_body,
            variable=self.low_quality_var, onvalue=True, offvalue=False,
            command=self.on_output_option_changed
        )
        self.low_quality_check.pack(anchor=tk.W, padx=16)
        ctk.CTkLabel(
            recog_card, text="雑音がひどい音源のみ ON を推奨", font=self.font_note,
            text_color=NOTE_TEXT_COLOR, height=16, anchor="w"
        ).pack(anchor=tk.W, padx=(16 + SWITCH_TEXT_OFFSET, 16), pady=(0, 10))

        diarization_text = "話者分離を行う"
        if not self.diarization_model_path:
            diarization_text += "（モデル未導入）"
        self.diarization_check = ctk.CTkSwitch(
            recog_card, text=diarization_text, font=self.font_body,
            variable=self.diarization_var, onvalue=True, offvalue=False,
            command=self.on_output_option_changed
        )
        self.diarization_check.pack(anchor=tk.W, padx=16)
        if not self.diarization_model_path:
            self.diarization_check.configure(state="disabled")
        ctk.CTkLabel(
            recog_card, text="Excel・Word に話者を記載。処理時間が延びます", font=self.font_note,
            text_color=NOTE_TEXT_COLOR, height=16, anchor="w"
        ).pack(anchor=tk.W, padx=(16 + SWITCH_TEXT_OFFSET, 16), pady=(0, 16))

        # 単語登録
        hotwords_card, hotwords_header = self._make_card(
            right_col, "単語登録（固有名詞・専門用語）", expand=True)

        self.hotwords_count_label = ctk.CTkLabel(
            hotwords_header, text=f"0 / {MAX_HOTWORDS} 件", font=self.font_note,
            text_color=NOTE_TEXT_COLOR)
        self.hotwords_count_label.pack(side=tk.RIGHT)

        input_row = ctk.CTkFrame(hotwords_card, fg_color="transparent", corner_radius=0)
        input_row.pack(fill=tk.X, padx=16)

        self.hotword_entry = ctk.CTkEntry(
            input_row, placeholder_text="単語を入力して Enter または 追加", font=self.font_body)
        self.hotword_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.hotword_entry.bind("<Return>", lambda e: self.add_hotword())

        self.add_btn = ctk.CTkButton(
            input_row, text="追加", width=64, font=self.font_body,
            command=self.add_hotword, **OUTLINE_BUTTON_STYLE)
        self.add_btn.pack(side=tk.LEFT, padx=(8, 0))

        self.hotwords_list_frame = ctk.CTkScrollableFrame(
            hotwords_card, height=64, corner_radius=8, fg_color=LIST_FG_COLOR)
        self.hotwords_list_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=(10, 8))
        # 内蔵スクロールバーは既定で高さ 200 を要求し、一覧の height 指定より優先されてしまう。
        # 公開 API が無いため内部属性で要求高さだけを下げる（表示時は一覧の高さまで伸びる）
        list_scrollbar = getattr(self.hotwords_list_frame, "_scrollbar", None)
        if list_scrollbar is not None:
            list_scrollbar.configure(height=40)
        else:
            logger.warning("単語一覧のスクロールバーを取得できませんでした（customtkinter の内部構造が変わった可能性）。"
                           "一覧の初期高さが大きくなります")
        self.hotword_rows = []

        hotwords_tools_row = ctk.CTkFrame(hotwords_card, fg_color="transparent", corner_radius=0)
        hotwords_tools_row.pack(fill=tk.X, padx=16, pady=(0, 16))
        self.hotwords_tools_row = hotwords_tools_row
        self.hotwords_export_btn = ctk.CTkButton(
            hotwords_tools_row, text="書き出し", command=self.export_hotwords, font=self.font_small,
            height=26, width=0, **OUTLINE_BUTTON_STYLE)
        self.hotwords_export_btn.pack(side=tk.LEFT)
        self.hotwords_import_btn = ctk.CTkButton(
            hotwords_tools_row, text="取り込み", command=self.import_hotwords, font=self.font_small,
            height=26, width=0, **OUTLINE_BUTTON_STYLE)
        self.hotwords_import_btn.pack(side=tk.LEFT, padx=(8, 0))

        # 出力オプション
        output_card, _ = self._make_card(right_col, "出力オプション")

        output_options = (
            ("分割音声 (1分ごと・再生用) も出力", self.output_split_var),
            ("テキスト (.txt) も出力", self.output_txt_var),
            ("字幕 (.srt) も出力", self.output_srt_var),
            ("Word (.docx) も出力", self.output_docx_var),
        )
        output_checks = []
        for i, (text, var) in enumerate(output_options):
            check = ctk.CTkCheckBox(
                output_card, text=text, font=self.font_body,
                variable=var, onvalue=True, offvalue=False,
                command=self.on_output_option_changed
            )
            is_last = i == len(output_options) - 1
            check.pack(anchor=tk.W, padx=16, pady=(0, 16 if is_last else 8))
            output_checks.append(check)
        (self.output_split_check, self.output_txt_check,
         self.output_srt_check, self.output_docx_check) = output_checks

    def _make_card(self, parent, title, expand=False):
        """角丸のカード（見出し付き）を parent に縦積みし、(カード, 見出し行) を返す。

        expand=True のカードは列の余った高さを吸収する。
        """
        top_pad = 12 if parent.pack_slaves() else 0
        card = ctk.CTkFrame(parent, corner_radius=12, fg_color=CARD_FG_COLOR)
        if expand:
            card.pack(fill=tk.BOTH, expand=True, pady=(top_pad, 0))
        else:
            card.pack(fill=tk.X, pady=(top_pad, 0))
        header = ctk.CTkFrame(card, fg_color="transparent", corner_radius=0)
        header.pack(fill=tk.X, padx=16, pady=(12, 4))
        ctk.CTkLabel(header, text=title, font=self.font_heading, anchor="w").pack(side=tk.LEFT)
        return card, header

    @staticmethod
    def _set_bar(bar, ratio):
        """プログレスバーを ratio (0..1) にする。0 のときは端の丸い点も見えないよう進捗色を地の色にする"""
        if ratio <= 0:
            bar.configure(progress_color=bar.cget("fg_color"))
        else:
            bar.configure(progress_color=ctk.ThemeManager.theme["CTkProgressBar"]["progress_color"])
        bar.set(ratio)

    def _on_appearance_mode_changed(self, mode_string):
        try:
            self.scroll_canvas.configure(
                bg=self.root._apply_appearance_mode(ctk.ThemeManager.theme["CTk"]["fg_color"]))
        except tk.TclError:
            pass

    def on_scroll_content_configure(self, event):
        self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox('all'))
        self.update_scroll_bar_visibility()

    def on_scroll_canvas_configure(self, event):
        # 窓が内容より高いときは内容を窓の高さまで伸ばし、余りを単語登録の一覧に吸収させる
        self.scroll_canvas.itemconfigure(
            self.scroll_window_id, width=event.width,
            height=max(event.height, self.scroll_content.winfo_reqheight()))
        self.update_scroll_bar_visibility()

    def refit_scroll_content(self):
        """内容の必要高さが変わったとき（選択ファイル一覧の表示・非表示）に、内容の高さとスクロール範囲を合わせ直す。

        内容の高さは窓の <Configure> でしか更新されないため、窓の大きさが変わらない変化はここで反映する。
        """
        self.root.update_idletasks()
        self.scroll_canvas.itemconfigure(
            self.scroll_window_id,
            height=max(self.scroll_canvas.winfo_height(), self.scroll_content.winfo_reqheight()))
        self.root.update_idletasks()
        self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox('all'))
        self.update_scroll_bar_visibility()

    def _fit_window_to_content(self):
        """内容が窓より高くなったとき、作業領域に収まる範囲で窓を高くする（縮めない・最大化中は何もしない）。

        寸法は apply_initial_geometry と同じく、物理ピクセルを DPI 倍率で割った論理ピクセルで扱う。
        geometry() の位置 (+x+y) は customtkinter が倍率を掛けないため物理ピクセルのまま。
        """
        try:
            if self.root.state() == "zoomed":
                return
            self.root.update_idletasks()
            scaling = ctk.ScalingTracker.get_window_scaling(self.root)
            content_px = (
                self.scroll_content.winfo_reqheight()
                + self.bottom_frame.winfo_reqheight()
                + self.bottom_separator.winfo_reqheight()
            )
            needed_height = math.ceil(content_px / scaling)
            m = re.match(r"^(\d+)x(\d+)\+(-?\d+)\+(-?\d+)$", self.root.geometry())
            if not m:
                return
            width, current_height, x, y = (int(v) for v in m.groups())
            if needed_height <= current_height:
                return
            work_area_px = get_work_area()
            if work_area_px is not None:
                work_area_height = int(work_area_px[1] / scaling)
            else:
                work_area_height = self.root.winfo_screenheight()
            new_height = max(WINDOW_MIN_HEIGHT, min(needed_height, work_area_height - WINDOW_CHROME_MARGIN))
            if new_height <= current_height:
                return
            # 下端が窓のあるモニターの作業領域からはみ出すなら、はみ出す分だけ上に寄せる
            rects = get_window_and_monitor_work_rect(ctypes.windll.user32.GetParent(self.root.winfo_id()))
            if rects is not None:
                window_rect, work_rect = rects
                new_bottom = window_rect[3] + round((new_height - current_height) * scaling)
                if new_bottom > work_rect[3]:
                    y = max(work_rect[1], y - (new_bottom - work_rect[3]))
            self.root.geometry(f"{width}x{new_height}+{x}+{y}")
            logger.info(f"選択ファイル一覧の表示に合わせて窓の高さを {current_height} → {new_height} にしました")
        except Exception:
            logger.warning("窓の高さの調整に失敗しました", exc_info=True)

    def is_scroll_content_overflowing(self):
        return self.scroll_content.winfo_reqheight() > self.scroll_canvas.winfo_height()

    def update_scroll_bar_visibility(self):
        """内容がはみ出すときだけスクロールバーを表示する"""
        if self.is_scroll_content_overflowing():
            if not self.scroll_bar.winfo_manager():
                self.scroll_bar.pack(side=tk.RIGHT, fill=tk.Y, before=self.scroll_canvas)
        else:
            if self.scroll_bar.winfo_manager():
                self.scroll_bar.pack_forget()
            self.scroll_canvas.yview_moveto(0)

    def on_mouse_wheel(self, event):
        if not self.is_scroll_content_overflowing():
            return
        widget = event.widget
        if isinstance(widget, str):
            try:
                widget = self.root.nametowidget(widget)
            except (KeyError, tk.TclError):
                return
        widget_path = str(widget)

        def is_within(ancestor):
            ancestor_path = str(ancestor)
            return widget_path == ancestor_path or widget_path.startswith(ancestor_path + ".")

        # スクロール領域外（ダイアログ等）と、単語リスト・選択ファイル一覧が自前でスクロールできるときは対象外
        if not is_within(self.scroll_canvas):
            return
        for list_frame in (self.hotwords_list_frame, self.input_list_frame):
            list_canvas = list_frame.master
            if is_within(list_canvas.master) and list_canvas.yview() != (0.0, 1.0):
                return
        self.scroll_canvas.yview_scroll(-int(event.delta / 120), 'units')

    def cancel_processing(self):
        """処理のキャンセルを要求"""
        self.cancel_event.set()
        self.cancel_btn.configure(state="disabled")
        self.status_label.configure(text="キャンセル中...")

    @staticmethod
    def format_mmss(seconds):
        """秒数を MM:SS 形式の文字列に変換（録音経過時間表示用）"""
        sec = int(seconds)
        return f"{sec // 60:02d}:{sec % 60:02d}"

    def _recording_save_dir(self):
        """録音ファイルの保存先（出力フォルダ → デスクトップ → ホーム）"""
        if self.output_folder_path:
            return self.output_folder_path
        desktop = get_known_folder(FOLDERID_DESKTOP)
        if desktop and os.path.isdir(desktop):
            return desktop
        return os.path.expanduser("~")

    def _resolve_output_dir(self, file_path):
        """出力先フォルダ（出力フォルダが指定されていればそこ、未指定なら音声ファイルのあるフォルダ）"""
        if self.output_folder_path:
            return self.output_folder_path
        return os.path.dirname(os.path.abspath(file_path))

    def start_recording(self):
        """マイク録音を開始する"""
        if self.is_processing or self.is_recording:
            return
        save_dir = self._recording_save_dir()

        file_name = f"録音_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.wav"
        output_path = os.path.join(save_dir, file_name)

        try:
            os.makedirs(save_dir, exist_ok=True)
            wave_file = wave.open(output_path, 'wb')
            wave_file.setnchannels(RECORD_CHANNELS)
            wave_file.setsampwidth(2)
            wave_file.setframerate(RECORD_SAMPLE_RATE)
        except OSError:
            logger.exception(f"録音ファイルを作成できませんでした: {output_path}")
            messagebox.showerror(
                "エラー",
                f"録音ファイルを作成できませんでした。\n保存先フォルダの権限を確認してください。\n{save_dir}"
            )
            return

        self.record_queue = queue.Queue()
        self.record_error = None
        self.record_level = 0

        def on_audio_block(indata, frames, time_info, status):
            if status:
                logger.warning(f"録音ステータス: {status}")
            if len(indata):
                peak = int(np.abs(indata.astype(np.int32)).max())
                self.record_level = min(int(peak / 32767 * 100), 100)
            self.record_queue.put(bytes(indata))

        try:
            stream = sd.InputStream(
                samplerate=RECORD_SAMPLE_RATE, channels=RECORD_CHANNELS,
                dtype=RECORD_DTYPE, callback=on_audio_block
            )
            stream.start()
        except Exception:
            logger.exception("マイクのオープンに失敗しました。")
            wave_file.close()
            try:
                os.remove(output_path)
            except OSError:
                pass
            messagebox.showerror("エラー", "マイクを開けませんでした。マイクが接続されているか確認してください。")
            return

        def writer():
            while True:
                data = self.record_queue.get()
                if data is None:
                    break
                try:
                    wave_file.writeframes(data)
                except OSError:
                    logger.exception("録音データの書き込みに失敗しました。")
                    self.record_error = "録音データの書き込みに失敗しました。"
                    break

        self.record_stream = stream
        self.record_wave_file = wave_file
        self.record_output_path = output_path
        self.record_writer_thread = threading.Thread(target=writer, daemon=True)
        self.record_writer_thread.start()

        self.is_recording = True
        self.record_start_dt = datetime.datetime.now()
        self._set_recording_ui(True)
        logger.info(f"録音を開始しました: {output_path} (保存先フォルダ: {save_dir})")
        self._poll_recording()

    def _poll_recording(self):
        """録音中の経過時間・レベルメーターを定期更新する"""
        if not self.is_recording:
            return
        elapsed = (datetime.datetime.now() - self.record_start_dt).total_seconds()
        self.record_time_label.configure(text=self.format_mmss(elapsed))
        self._set_bar(self.record_level_bar, self.record_level / 100)
        if self.record_error:
            message = self.record_error
            self.record_error = None
            self._finalize_recording()
            messagebox.showerror("エラー", f"{message}\n録音を停止しました。")
            return
        self.record_poll_id = self.root.after(150, self._poll_recording)

    def _finalize_recording(self):
        """録音ストリーム・ライタースレッド・WAVファイルを後始末する"""
        if self.record_poll_id is not None:
            self.root.after_cancel(self.record_poll_id)
            self.record_poll_id = None
        try:
            if self.record_stream is not None:
                self.record_stream.stop()
                self.record_stream.close()
        except Exception:
            logger.exception("録音ストリームの停止に失敗しました。")
        self.record_stream = None

        if self.record_queue is not None:
            self.record_queue.put(None)
        if self.record_writer_thread is not None:
            self.record_writer_thread.join(timeout=5)
        self.record_writer_thread = None

        try:
            if self.record_wave_file is not None:
                self.record_wave_file.close()
        except Exception:
            logger.exception("録音ファイルのクローズに失敗しました。")
        self.record_wave_file = None

        self.is_recording = False
        self.record_time_label.configure(text="00:00")
        self._set_bar(self.record_level_bar, 0)
        self._set_recording_ui(False)

    def stop_recording(self, on_close=False):
        """マイク録音を停止する（on_close=True の場合はダイアログを出さず終了処理のみ行う）"""
        if not self.is_recording:
            return
        output_path = self.record_output_path
        self._finalize_recording()
        logger.info(f"録音を保存しました: {output_path}")
        if on_close:
            return
        self._add_input_files([output_path], remember=False)
        count = len(self.input_file_paths)
        question = (f"選択中の {count} 件をこのまま文字起こししますか？" if count >= 2
                    else "このまま文字起こしを開始しますか？")
        if messagebox.askyesno(
                "録音完了",
                f"録音を保存しました。\n保存先: {os.path.dirname(output_path)}\n\n{question}"):
            self.start_processing()

    def select_input_file(self):
        filetypes = [("音声ファイル", "*.wav;*.mp3;*.m4a;*.mp4")]
        file_paths = filedialog.askopenfilenames(
            title="音声ファイルを選択（複数選択可）",
            filetypes=filetypes,
            initialdir=self._default_browse_dir("input")
        )
        if file_paths:
            self._add_input_files(file_paths)

    def _default_browse_dir(self, kind):
        """選択ダイアログの初期フォルダ（前回の場所 → [出力のみ] 入力ファイルの場所 → デスクトップ → ホーム）"""
        if kind == "output":
            candidates = [self.settings.get("last_output_dir")]
            if self.input_file_paths:
                candidates.append(os.path.dirname(os.path.normpath(self.input_file_paths[0])))
        else:
            candidates = [self.settings.get("last_input_dir")]
        candidates.append(get_known_folder(FOLDERID_DESKTOP))
        for path in candidates:
            if isinstance(path, str) and path and os.path.isdir(path):
                return path
        return os.path.expanduser("~")

    def _remember_browse_dir(self, key, folder):
        """選んだフォルダを次回のダイアログの初期フォルダとして設定に保存する"""
        if not folder or self.settings.get(key) == folder:
            return
        self.settings[key] = folder
        try:
            save_settings(self.settings)
        except OSError:
            logger.warning("設定ファイルへの %s の保存に失敗しました", key, exc_info=True)

    def _set_input_files(self, paths):
        """入力ファイルを設定し、ファイル表示（1 件ならパス、2 件以上なら件数と一覧）を更新する"""
        self.input_file_paths = list(paths)
        if self.input_file_paths:
            self._remember_browse_dir(
                "last_input_dir", os.path.dirname(os.path.normpath(self.input_file_paths[0])))
        self._refresh_input_files_view()

    def _add_input_files(self, paths, remember=True):
        """選択中のファイルに paths を追加する（同じファイルは大文字小文字・区切り文字の違いも含めて 1 つにし、順序は保つ）。

        追加した件数を返す。0 件（すべて選択済み）のときは表示を変えず、状態欄で知らせるだけにする。
        """
        def key(p):
            return os.path.normcase(os.path.normpath(p))

        known = {key(p) for p in self.input_file_paths}
        new_paths = []
        for p in paths:
            k = key(p)
            if k not in known:
                known.add(k)
                new_paths.append(p)
        if not new_paths:
            self._show_transient_status("既に選択済みです")
            return 0
        self.input_file_paths = self.input_file_paths + new_paths
        if remember:
            self._remember_browse_dir("last_input_dir", os.path.dirname(os.path.normpath(new_paths[0])))
        self._refresh_input_files_view()
        return len(new_paths)

    def _show_transient_status(self, text, ms=2000):
        """処理中でなければ状態欄に text を少しの間だけ出す"""
        if self.is_processing:
            return
        previous = self.status_label.cget("text")
        self.status_label.configure(text=text)

        def restore():
            if not self.is_processing and self.status_label.cget("text") == text:
                self.status_label.configure(text=previous)

        self.root.after(ms, restore)

    def _refresh_input_files_view(self):
        """self.input_file_paths に合わせてファイル表示ラベルと選択ファイル一覧を作り直す"""
        count = len(self.input_file_paths)
        if count == 0:
            self._set_path_label(self.file_label, INPUT_PLACEHOLDER_TEXT, placeholder=True)
        elif count == 1:
            file_path = self.input_file_paths[0]
            # 長いパスは省略表示
            display_path = file_path if len(file_path) < 50 else "..." + file_path[-47:]
            self._set_path_label(self.file_label, display_path, placeholder=False)
        else:
            self._set_path_label(self.file_label, f"{count} 件選択", placeholder=False)
        self._update_clear_button_state()

        for row, _remove_btn in self.input_file_rows:
            row.destroy()
        self.input_file_rows = []
        if count >= 2:
            enabled = not (self.is_processing or self.is_recording)
            for index, file_path in enumerate(self.input_file_paths):
                self.input_file_rows.append(self._make_list_row(
                    self.input_list_frame, os.path.basename(file_path),
                    lambda i=index: self._remove_input_file(i), enabled))
            self.input_list_frame.configure(height=LIST_ROW_HEIGHT * min(count, INPUT_LIST_MAX_ROWS))
            if not self.input_list_visible:
                self.input_list_frame.pack(fill=tk.X, padx=16, pady=(8, 0), after=self.file_row)
                self.input_list_visible = True
            self.input_list_frame.master.yview_moveto(0)
        elif self.input_list_visible:
            self.input_list_frame.pack_forget()
            self.input_list_visible = False
        self.refit_scroll_content()
        self._fit_window_to_content()

    def _update_clear_button_state(self):
        """「クリア」は選択が 1 件以上あり、処理中・録音中でないときだけ押せる"""
        enabled = bool(self.input_file_paths) and not (self.is_processing or self.is_recording)
        self.clear_input_btn.configure(state=tk.NORMAL if enabled else tk.DISABLED)

    def _remove_input_file(self, index):
        """選択ファイル一覧の ✕ で、該当ファイルを入力から外す"""
        if self.is_processing or self.is_recording:
            return
        if 0 <= index < len(self.input_file_paths):
            del self.input_file_paths[index]
            self._refresh_input_files_view()

    def setup_drop_target(self):
        """ウィンドウ全体を音声ファイルのドロップ先にする（tkdnd を読めなかった場合は何もしない）"""
        if self.root.TkdndVersion is None:
            return
        try:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<DropEnter>>", self.on_drop_enter)
            self.root.dnd_bind("<<DropLeave>>", self.on_drop_leave)
            self.root.dnd_bind("<<Drop>>", self.on_drop_files)
        except tk.TclError:
            logger.warning("ドロップ先の登録に失敗しました。ファイル選択ボタンのみ使用できます", exc_info=True)

    def _set_drop_highlight(self, active):
        if active:
            self.input_card.configure(
                border_width=2, border_color=ctk.ThemeManager.theme["CTkButton"]["fg_color"])
        else:
            self.input_card.configure(border_width=0)

    def on_drop_enter(self, event):
        if not (self.is_processing or self.is_recording):
            self._set_drop_highlight(True)
        return COPY

    def on_drop_leave(self, event):
        self._set_drop_highlight(False)

    def on_drop_files(self, event):
        """ドロップされたファイルのうち対応する音声ファイルだけを入力に設定する"""
        self._set_drop_highlight(False)
        if self.is_processing or self.is_recording:
            return COPY
        paths = []
        other_files = 0
        for path in self.root.tk.splitlist(event.data):
            if os.path.isdir(path) or os.path.splitext(path)[1].lower() == ".lnk":
                continue
            if os.path.splitext(path)[1].lower() in AUDIO_EXTENSIONS:
                paths.append(os.path.normpath(path))
            else:
                other_files += 1
        if not paths:
            if other_files:
                messagebox.showwarning("警告", "対応していないファイルです。wav / mp3 / m4a / mp4 を選択してください。")
            else:
                messagebox.showwarning(
                    "警告",
                    "フォルダやショートカットはドロップできません。音声ファイル (wav / mp3 / m4a / mp4) を選んでください。")
            return COPY
        added = self._add_input_files(paths)
        logger.info(f"ドロップで入力ファイルを {added} 件追加しました（選択中 {len(self.input_file_paths)} 件）")
        return COPY

    def select_output_folder(self):
        folder_path = filedialog.askdirectory(
            title="出力するフォルダを選択", initialdir=self._default_browse_dir("output"))
        if folder_path:
            self._remember_browse_dir("last_output_dir", os.path.normpath(folder_path))
            self.output_folder_path = folder_path
            display_path = folder_path if len(folder_path) < 50 else "..." + folder_path[-47:]
            self._set_path_label(self.folder_label, display_path, placeholder=False)

    @staticmethod
    def _set_path_label(label, text, placeholder: bool):
        """パス表示ラベルを更新する（未選択の案内文はプレースホルダー色）"""
        text_color = PLACEHOLDER_FOREGROUND if placeholder else ctk.ThemeManager.theme["CTkLabel"]["text_color"]
        label.configure(text=text, text_color=text_color)

    def populate_hotwords_listbox(self):
        """登録済み単語の一覧を 1 単語 1 行で作り直す"""
        for _word, row, _remove_btn in self.hotword_rows:
            row.destroy()
        self.hotword_rows = []
        for word in self.hotwords_list:
            self._add_hotword_row(word)
        self.update_hotwords_count()

    def _add_hotword_row(self, word):
        """単語 1 つ分の行（単語 + 右端の削除ボタン）を一覧の末尾に足す"""
        row, remove_btn = self._make_list_row(
            self.hotwords_list_frame, word, lambda w=word: self.remove_hotword(w), not self.is_processing)
        self.hotword_rows.append((word, row, remove_btn))

    def _make_list_row(self, list_frame, text, on_remove, enabled):
        """一覧の末尾に 1 行（文字 + 右端の ✕ ボタン）を足し、(行, ✕ ボタン) を返す"""
        row = ctk.CTkFrame(list_frame, fg_color="transparent", corner_radius=0)
        row.pack(fill=tk.X, padx=(6, 0), pady=1)
        remove_btn = ctk.CTkButton(
            row, text="✕", width=28, height=24, font=self.font_body,
            fg_color="transparent", text_color=NOTE_TEXT_COLOR,
            command=on_remove, state="normal" if enabled else "disabled")
        remove_btn.pack(side=tk.RIGHT)
        ctk.CTkLabel(row, text=text, font=self.font_body, height=24, anchor="w").pack(
            side=tk.LEFT, fill=tk.X, expand=True)
        return row, remove_btn

    def update_hotwords_count(self):
        """登録数カウンターを更新"""
        count = len(self.hotwords_list)
        self.hotwords_count_label.configure(text=f"{count} / {MAX_HOTWORDS} 件")

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
        self._add_hotword_row(word)
        self.hotword_entry.delete(0, tk.END)
        self.update_hotwords_count()
        # 追加した行が見えるよう一覧の末尾までスクロールする
        self.root.update_idletasks()
        self.hotwords_list_frame.master.yview_moveto(1.0)

    def remove_hotword(self, word):
        """指定した単語を削除"""
        if word not in self.hotwords_list:
            return
        self.hotwords_list.remove(word)
        save_hotwords(self.hotwords_list)
        self.populate_hotwords_listbox()

    def export_hotwords(self):
        """登録済み単語をテキストファイル（1 行 1 語、UTF-8 BOM 付き・CRLF）に書き出す"""
        if not self.hotwords_list:
            messagebox.showinfo("情報", "登録された単語がありません")
            return
        path = filedialog.asksaveasfilename(
            title="単語登録を書き出し",
            initialdir=self._default_browse_dir("output"),
            initialfile=HOTWORDS_EXPORT_FILE_NAME,
            defaultextension=".txt",
            filetypes=[("テキストファイル", "*.txt")],
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8-sig", newline="") as f:
                f.write("".join(f"{word}\r\n" for word in self.hotwords_list))
        except OSError:
            logger.exception(f"単語登録の書き出しに失敗しました: {path}")
            messagebox.showerror(
                "エラー",
                f"単語登録の書き出しに失敗しました。\n{path}\n\n"
                f"詳細はログ (logs/app-YYYYMMDD.log) を確認してください。"
            )
            return
        logger.info(f"単語登録 {len(self.hotwords_list)} 件を書き出しました: {path}")
        messagebox.showinfo("書き出し", f"{len(self.hotwords_list)} 件を書き出しました。")

    def import_hotwords(self):
        """テキストファイル（1 行 1 語）の単語を登録に追加する（空行・重複・上限超過分は除く）"""
        path = filedialog.askopenfilename(
            title="単語登録を取り込み",
            initialdir=self._default_browse_dir("output"),
            filetypes=[("テキストファイル", "*.txt")],
        )
        if not path:
            return
        try:
            with open(path, "rb") as f:
                data = f.read()
        except OSError:
            logger.exception(f"単語登録ファイルを読み込めませんでした: {path}")
            messagebox.showerror(
                "エラー",
                f"ファイルを読み込めませんでした。\n{path}\n\n"
                f"詳細はログ (logs/app-YYYYMMDD.log) を確認してください。"
            )
            return
        try:
            if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
                text = data.decode("utf-16")
            else:
                try:
                    text = data.decode("utf-8-sig")
                except UnicodeDecodeError:
                    text = data.decode("cp932")
        except UnicodeDecodeError:
            logger.warning(f"単語登録ファイルの文字コードを判別できませんでした: {path}")
            messagebox.showerror(
                "エラー",
                "ファイルの文字コードを読み取れませんでした。\nUTF-8 または Shift_JIS のテキストファイルを選択してください。"
            )
            return

        added = 0
        duplicates = 0
        over_limit = 0
        invalid = 0
        seen = set()
        for line in text.splitlines():
            word = line.strip()
            if not word:
                continue
            # BOM なし UTF-16 を UTF-8 として読んだ場合などは NUL 等の制御文字が混じる
            if any(ord(c) < 32 for c in word):
                invalid += 1
                continue
            if word in seen or word in self.hotwords_list:
                duplicates += 1
                continue
            seen.add(word)
            if len(self.hotwords_list) >= MAX_HOTWORDS:
                over_limit += 1
            else:
                self.hotwords_list.append(word)
                added += 1

        if added:
            try:
                save_hotwords(self.hotwords_list)
            except OSError:
                logger.exception("単語登録の保存に失敗しました")
                messagebox.showerror(
                    "エラー",
                    "取り込んだ単語を保存できませんでした。今回の起動中のみ有効です。\n"
                    "詳細はログ (logs/app-YYYYMMDD.log) を確認してください。"
                )
            self.populate_hotwords_listbox()
        logger.info(f"単語登録を取り込みました: 追加 {added} 件, 重複 {duplicates} 件, 上限超過 {over_limit} 件, "
                    f"不正な行 {invalid} 件 ({path})")
        messagebox.showinfo(
            "取り込み",
            f"{added} 件を取り込みました（重複 {duplicates} 件、上限超過 {over_limit} 件、不正な行 {invalid} 件はスキップ）"
        )

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
        self.settings["output_docx"] = self.output_docx_var.get()
        self.settings["low_quality_mode"] = self.low_quality_var.get()
        self.settings["diarization"] = self.diarization_var.get()
        save_settings(self.settings)

    def update_progress(self, current, total, status_text, detail_text="", taskbar=True):
        """プログレスバーと状態表示を更新（taskbar=False のときはタスクバーの進捗に触れない）"""
        progress_value = (current / total) * 100 if total > 0 else 0
        self._set_bar(self.progress_bar, progress_value / 100)
        self.status_label.configure(text=status_text)
        self.progress_detail.configure(text=detail_text)
        if taskbar and total > 0:
            # タスクバーはバッチ全体の進み具合（画面のバーはファイルごと）
            file_index, file_count = self._batch_pos
            self.taskbar.set_progress((file_index - 1) * total + current, file_count * total)
        self.root.update_idletasks()

    def set_ui_state(self, enabled):
        """UIの有効/無効を切り替え"""
        state = tk.NORMAL if enabled else tk.DISABLED
        self.file_btn.configure(state=state)
        self.folder_btn.configure(state=state)
        self.run_btn.configure(state=state)
        self.add_btn.configure(state=state)
        for _word, _row, remove_btn in self.hotword_rows:
            remove_btn.configure(state=state)
        self.hotwords_export_btn.configure(state=state)
        self.hotwords_import_btn.configure(state=state)
        input_remove_state = tk.NORMAL if (enabled and not self.is_recording) else tk.DISABLED
        for _row, remove_btn in self.input_file_rows:
            remove_btn.configure(state=input_remove_state)
        self.clear_input_btn.configure(
            state=tk.NORMAL if (enabled and self.input_file_paths and not self.is_recording) else tk.DISABLED)
        self.hotword_entry.configure(state=state)
        self.output_split_check.configure(state=state)
        self.output_txt_check.configure(state=state)
        self.output_srt_check.configure(state=state)
        self.output_docx_check.configure(state=state)
        self.low_quality_check.configure(state=state)
        if self.diarization_model_path:
            self.diarization_check.configure(state=state)
        if hasattr(self, "model_combo"):
            self.model_combo.configure(state=state)
        self.cancel_btn.configure(state=tk.DISABLED if enabled else tk.NORMAL)
        self.record_start_btn.configure(state=tk.DISABLED if (not enabled or self.is_recording) else tk.NORMAL)

    def _set_recording_ui(self, recording):
        """録音中は入力ファイル選択・出力フォルダ選択・文字起こし開始を無効化する"""
        lock_state = tk.DISABLED if recording else tk.NORMAL
        self.file_btn.configure(state=lock_state)
        self.folder_btn.configure(state=lock_state)
        self.run_btn.configure(state=lock_state)
        for _row, remove_btn in self.input_file_rows:
            remove_btn.configure(state=lock_state)
        self.clear_input_btn.configure(
            state=tk.NORMAL if (not recording and self.input_file_paths) else tk.DISABLED)
        self.record_start_btn.configure(state=tk.DISABLED if recording else tk.NORMAL)
        self.record_stop_btn.configure(state=tk.NORMAL if recording else tk.DISABLED)

    def start_processing(self):
        if self.is_recording:
            messagebox.showwarning("警告", "録音中は文字起こしを開始できません。録音を停止してください。")
            return
        if not self.input_file_paths:
            messagebox.showwarning("警告", "入力ファイルを選択してください。")
            return

        for output_dir in dict.fromkeys(self._resolve_output_dir(p) for p in self.input_file_paths):
            # tempfile.NamedTemporaryFile は書き込み拒否 (ACL) のフォルダで名前を変えて再試行を続け、
            # 1 分以上戻らない（Windows の os.access が ACL を見ないため）。1 回だけ作って消す
            probe_path = os.path.join(output_dir, f".write_test_{os.getpid()}_{uuid.uuid4().hex}.tmp")
            try:
                os.makedirs(output_dir, exist_ok=True)
                fd = os.open(probe_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY | os.O_TEMPORARY)
                os.close(fd)
            except OSError as e:
                logger.warning(f"出力先に書き込めません: {output_dir} ({e})")
                messagebox.showerror("エラー", f"出力先 {output_dir} に書き込めません。出力フォルダを指定してください。")
                return

        existing_names = []
        for file_path in self.input_file_paths:
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            output_file = os.path.join(self._resolve_output_dir(file_path), f"{file_name}_output.xlsx")
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
        diarization_pipeline = None
        want_diarization = bool(self.diarization_var.get() and self.diarization_model_path)
        last_output_dir = None
        output_dirs = []
        # 完了・キャンセルの通知を予約したか（予約していれば、タスクバーの進捗はその通知の側で消す）
        outcome_scheduled = False
        try:
            for file_index, file_path in enumerate(self.input_file_paths, start=1):
                if self.cancel_event.is_set():
                    raise ProcessingCancelled()
                file_name = os.path.basename(file_path)
                output_dir = self._resolve_output_dir(file_path)
                last_output_dir = output_dir
                # 書き換えは UI スレッドの after 経由にして、前のファイルの進捗更新より後に反映させる
                self.root.after(0, lambda pos=(file_index, file_count): setattr(self, "_batch_pos", pos))
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
                    if want_diarization and diarization_pipeline is None:
                        # 話者分離パイプラインはバッチ全体で1回だけロードする。
                        # torch / pyannote は重いため、実際に必要になった時点で初めて import する。
                        try:
                            self.root.after(0, lambda: self.update_progress(
                                0, 100, "話者分離モデルを読み込み中...", ""))
                            from pyannote.audio import Pipeline
                            diarization_pipeline = Pipeline.from_pretrained(
                                DIARIZATION_MODEL_REPO, cache_dir=self.diarization_model_path
                            )
                        except Exception:
                            logger.exception("話者分離モデルの読み込みに失敗しました。話者分離なしで続行します。")
                            want_diarization = False
                            diarization_pipeline = None
                    self.transcribe_file(
                        model, file_path, output_dir,
                        file_index=file_index, file_count=file_count,
                        diarization_pipeline=diarization_pipeline if want_diarization else None
                    )
                    succeeded += 1
                    if output_dir not in output_dirs:
                        output_dirs.append(output_dir)
                except ProcessingCancelled:
                    raise
                except Exception:
                    logger.exception(f"文字起こし処理に失敗しました: {file_path}")
                    failed_names.append(file_name)

            self.root.after(0, lambda: self.on_process_complete(
                file_count, succeeded, failed_names, last_output_dir, output_dirs))
            outcome_scheduled = True
        except ProcessingCancelled:
            logger.info("ユーザー操作により処理がキャンセルされました。")
            self.root.after(0, self.on_process_cancelled)
            outcome_scheduled = True
        finally:
            self.is_processing = False
            self.root.after(0, lambda: self.set_ui_state(True))
            self.root.after(0, lambda: self.update_progress(0, 1, "待機中...", "", taskbar=False))
            if not outcome_scheduled:
                self.root.after(0, self.taskbar.clear)

    def on_process_cancelled(self):
        """キャンセル完了時の処理"""
        self.taskbar.clear()
        messagebox.showinfo("キャンセル", "処理をキャンセルしました。")

    @staticmethod
    def format_output_dirs(output_dirs, limit=5):
        """完了ダイアログに載せる出力先フォルダの一覧（limit 件を超えたら「ほか N 件」）"""
        if not output_dirs:
            return ""
        lines = list(output_dirs[:limit])
        if len(output_dirs) > limit:
            lines.append(f"ほか {len(output_dirs) - limit} 件")
        return "\n\n出力先:\n" + "\n".join(lines)

    def on_process_complete(self, file_count, succeeded, failed_names, last_output_dir=None, output_dirs=None):
        """処理完了時の処理"""
        self.notify_completion()
        dirs_text = self.format_output_dirs(output_dirs or [])
        if not failed_names:
            self.taskbar.clear()
            messagebox.showinfo("完了", f"処理が完了しました。{dirs_text}")
        else:
            self.taskbar.set_error()
            names_text = "\n".join(failed_names)
            messagebox.showwarning(
                "完了",
                f"{file_count}件中{succeeded}件成功。\n失敗: {names_text}\n\n"
                f"詳細はログファイル (logs フォルダ) を確認してください。{dirs_text}"
            )
            self.taskbar.clear()
        # 出力先フォルダを開く（未指定時は最後に処理したファイルの出力先）
        open_dir = self.output_folder_path or last_output_dir
        if open_dir:
            try:
                os.startfile(open_dir)
            except OSError:
                logger.warning(f"出力先フォルダを開けませんでした: {open_dir}", exc_info=True)

    def notify_completion(self):
        """処理完了を音とタスクバー点滅で通知する"""
        try:
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
            FLASHW_ALL = 0x00000003
            FLASHW_TIMERNOFG = 0x0000000C

            class FLASHWINFO(ctypes.Structure):
                _fields_ = [
                    ("cbSize", ctypes.c_uint),
                    ("hwnd", ctypes.c_void_p),
                    ("dwFlags", ctypes.c_uint),
                    ("uCount", ctypes.c_uint),
                    ("dwTimeout", ctypes.c_uint),
                ]

            hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
            info = FLASHWINFO(
                ctypes.sizeof(FLASHWINFO), hwnd,
                FLASHW_ALL | FLASHW_TIMERNOFG, 5, 0
            )
            ctypes.windll.user32.FlashWindowEx(ctypes.byref(info))
        except Exception:
            pass

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

    def diarize_and_assign(self, diarization_pipeline, file_path, raw_segments, prefix):
        """話者分離を実行し、raw_segments と同じ順序で話者ラベル（話者A等）のリストを返す"""
        import torch

        audio = decode_audio_pyav_16k_mono(file_path)
        waveform = torch.from_numpy(audio).unsqueeze(0)
        audio_input = {"waveform": waveform, "sample_rate": DIARIZATION_SAMPLE_RATE}

        def hook(step_name, step_artifact, file=None, total=None, completed=None):
            if self.cancel_event.is_set():
                raise ProcessingCancelled()
            if total and completed is not None:
                detail_text = f"話者分離中: {step_name} {completed}/{total}"
            else:
                detail_text = f"話者分離中: {step_name}"
            self.root.after(0, lambda dt=detail_text: self.update_progress(
                96, 100, f"{prefix}話者分離中...", dt))

        diarization = diarization_pipeline(audio_input, hook=hook)
        annotation = getattr(diarization, "speaker_diarization", diarization)
        turns = [
            {"start": turn.start, "end": turn.end, "speaker": speaker}
            for turn, _, speaker in annotation.itertracks(yield_label=True)
        ]

        # 話者ラベルは登場順（区間の開始時刻順）に話者A、話者B…へ変換する
        speaker_order = []
        for turn in sorted(turns, key=lambda t: t["start"]):
            if turn["speaker"] not in speaker_order:
                speaker_order.append(turn["speaker"])
        label_map = {}
        for i, raw_speaker in enumerate(speaker_order):
            label_map[raw_speaker] = f"話者{chr(ord('A') + i)}" if i < 26 else f"話者{i + 1}"

        labels = []
        for seg_start, seg_end, _text in raw_segments:
            raw_speaker = assign_speaker_to_segment(seg_start, seg_end, turns)
            labels.append(label_map.get(raw_speaker, DIARIZATION_UNKNOWN_SPEAKER))
        return labels

    def transcribe_file(self, model, file_path, output_folder, file_index=1, file_count=1,
                         diarization_pipeline=None):
        file_extension = os.path.splitext(file_path)[1].lower()
        if file_extension not in (".wav", ".mp3", ".m4a", ".mp4"):
            raise ValueError("サポートされていない音声形式です。")
        file_name = os.path.splitext(os.path.basename(file_path))[0]

        prefix = f"[{file_index}/{file_count}] {os.path.basename(file_path)}: " if file_count > 1 else ""

        os.makedirs(output_folder, exist_ok=True)

        start_time_all = datetime.datetime.now()
        logger.info(f"処理開始: {file_path} (出力先: {output_folder})")

        transcribe_params = dict(
            beam_size=5,
            language='ja',
            vad_filter=True,
            initial_prompt=INITIAL_PROMPT,
            # 音質の悪い音源で一度誤認識（幻覚）が起きると、直前の出力を文脈として
            # 引き継ぐ仕組みにより残り全体へ連鎖するため無効化する。
            # v1.3の「1分ごとの独立認識」が持っていたリセット効果に相当。
            # temperature は既定のフォールバック（失敗時に温度を上げて再試行）を使う。
            condition_on_previous_text=False,
        )
        # hotwords は30秒ごとの認識窓すべてのプロンプトに注入されるため、
        # 句読点誘導文を載せることで全区間に句読点が付くようにする
        # （initial_prompt は先頭の窓にしか効かない）。ユーザー登録単語も併記する。
        hotwords_str = self.get_hotwords_string()
        transcribe_params["hotwords"] = (
            f"{INITIAL_PROMPT} {hotwords_str}" if hotwords_str else INITIAL_PROMPT
        )

        # 低品質音源モード: ノイズ抑制+音量正規化した波形を直接認識にかける。
        # 通常音源ではわずかに悪化しうるため既定OFF（雑音がひどい音源の救済用）。
        audio_input = file_path
        if self.low_quality_var.get():
            self.root.after(0, lambda: self.update_progress(
                3, 100, f"{prefix}音声を前処理中（ノイズ抑制）..."))
            logger.info("低品質音源モード: ノイズ抑制前処理を実行")
            audio_input = preprocess_low_quality(
                decode_audio(file_path, sampling_rate=16000), self.cancel_event)
            transcribe_params["vad_parameters"] = dict(threshold=0.25)

        # ファイル全体を1回で文字起こしし、タイムスタンプで1分単位にまとめる
        # （物理分割しないため文の途中で切れず、再エンコードによる劣化もない）
        self.root.after(0, lambda: self.update_progress(5, 100, f"{prefix}文字起こし中...", ""))
        segments, info = model.transcribe(audio_input, **transcribe_params)
        duration = max(info.duration or 0, 1.0)
        logger.info(f"音声長: {duration:.1f}秒")

        split_interval = 60
        buckets = {}
        bucket_logprobs = {}
        raw_segments = []
        total_chars = 0
        segment_loop_start = datetime.datetime.now()
        for segment in segments:
            if self.cancel_event.is_set():
                raise ProcessingCancelled()
            idx = int(segment.start // split_interval)
            text = str(segment.text).strip()
            buckets.setdefault(idx, []).append(text)
            raw_segments.append((segment.start, segment.end, text))
            total_chars += len(text)
            avg_logprob = segment.avg_logprob
            if idx not in bucket_logprobs or avg_logprob < bucket_logprobs[idx]:
                bucket_logprobs[idx] = avg_logprob

            elapsed = (datetime.datetime.now() - segment_loop_start).total_seconds()
            detail_text = ""
            if elapsed >= 1.0 and segment.end > 0:
                speed = segment.end / elapsed
                if speed > 0:
                    remaining = max(duration - segment.end, 0) / speed
                    if remaining >= 60:
                        detail_text = f"残り約{int(remaining // 60)}分"
                    else:
                        detail_text = f"残り約{int(remaining)}秒"

            progress = 5 + min(segment.end / duration, 1.0) * 93
            status_text = f"{prefix}文字起こし中... ({self.format_time(segment.end)} / {self.format_time(duration)})"
            self.root.after(0, lambda p=progress, st=status_text, dt=detail_text:
                            self.update_progress(p, 100, st, dt))

        low_conf_rows = {idx for idx, v in bucket_logprobs.items() if v < LOW_CONFIDENCE_LOGPROB}
        num_rows = max(buckets.keys()) + 1 if buckets else 1
        end_time_all = datetime.datetime.now()
        logger.info(f"文字起こし完了 ({total_chars}文字) 処理時間: {end_time_all - start_time_all}")

        # 話者分離: Whisperの認識結果には一切手を加えず、独立に実行した結果を後合成する。
        row_speakers = {}
        segment_speaker_labels = None
        diarization_succeeded = False
        if diarization_pipeline is not None:
            self.root.after(0, lambda: self.update_progress(96, 100, f"{prefix}話者分離中..."))
            try:
                speaker_labels = self.diarize_and_assign(
                    diarization_pipeline, file_path, raw_segments, prefix)
                for (seg_start, _seg_end, _text), label in zip(raw_segments, speaker_labels):
                    idx = int(seg_start // split_interval)
                    bucket_speakers = row_speakers.setdefault(idx, [])
                    if label not in bucket_speakers:
                        bucket_speakers.append(label)
                # 正規の話者が1人でもいるバケットからは「不明」を除く
                for idx in list(row_speakers.keys()):
                    real = [sp for sp in row_speakers[idx] if sp != DIARIZATION_UNKNOWN_SPEAKER]
                    if real:
                        row_speakers[idx] = real
                diarization_succeeded = True
                segment_speaker_labels = speaker_labels
                logger.info(f"話者分離完了: {file_path}")
            except ProcessingCancelled:
                raise
            except Exception:
                logger.exception(f"話者分離に失敗しました: {file_path}")
                row_speakers = {}
                self.root.after(0, lambda: self.update_progress(
                    96, 100, f"{prefix}話者分離に失敗しました。話者なしで続行します。"))
        show_speaker_column = diarization_succeeded

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
        if show_speaker_column:
            sheet.append(['No', '時間帯', '話者', '音声ファイル', '変換結果'])
            link_col, text_col = 4, 5
        else:
            sheet.append(['No', '時間帯', '音声ファイル', '変換結果'])
            link_col, text_col = 3, 4
        low_conf_fill = PatternFill(fill_type="solid", start_color="FFF9C4")
        for idx in range(num_rows):
            start_sec = idx * split_interval
            end_sec = min((idx + 1) * split_interval, duration)
            time_label = f"{self.format_time(start_sec)} - {self.format_time(end_sec)}"
            link_path = row_links.get(idx)
            link_name = os.path.basename(link_path) if link_path else ""
            text_value = '\n'.join(buckets.get(idx, []))
            if show_speaker_column:
                speaker_label = '、'.join(row_speakers.get(idx, []))
                sheet.append([str(idx), time_label, speaker_label, link_name, text_value])
            else:
                sheet.append([str(idx), time_label, link_name, text_value])
            # 音声ファイルセルから該当区間の分割音声を開けるようにリンクを付与
            if link_path:
                sheet.cell(row=idx + 2, column=link_col).hyperlink = link_path
            if idx in low_conf_rows:
                sheet.cell(row=idx + 2, column=text_col).fill = low_conf_fill
        sheet.column_dimensions['A'].width = 6
        sheet.column_dimensions['B'].width = 22
        if show_speaker_column:
            sheet.column_dimensions['C'].width = 16
            sheet.column_dimensions['D'].width = 28
            sheet.column_dimensions['E'].width = 100
        else:
            sheet.column_dimensions['C'].width = 28
            sheet.column_dimensions['D'].width = 100
        for row in sheet.iter_rows(min_row=2, min_col=text_col, max_col=text_col):
            row[0].alignment = Alignment(wrap_text=True, vertical='top')

        try:
            workbook.save(output_file)
        except PermissionError:
            raise RuntimeError(
                f"Excelファイルを保存できません。\n{output_file}\n"
                f"このファイルを開いている場合は閉じてから再実行してください。"
            )
        logger.info(f"Excelファイル {output_file} を保存しました。")

        if diarization_succeeded and segment_speaker_labels is not None:
            self.write_speakers_output(output_folder, file_name, raw_segments, segment_speaker_labels)

        if self.output_txt_var.get():
            self.write_txt_output(output_folder, file_name, buckets, num_rows, split_interval, duration)
        if self.output_srt_var.get():
            self.write_srt_output(output_folder, file_name, raw_segments)
        if self.output_docx_var.get():
            self.write_docx_output(output_folder, file_name, buckets, num_rows, split_interval, duration, low_conf_rows,
                                    row_speakers=row_speakers, show_speaker_column=show_speaker_column)

        # 進捗更新: 完了
        self.root.after(0, lambda: self.update_progress(100, 100, f"{prefix}完了！",
                       f"総処理時間: {end_time_all - start_time_all}"))

    def write_speakers_output(self, output_folder, file_name, raw_segments, speaker_labels):
        """話者分離結果を発言単位（連続する同一話者を1行に結合）でExcel出力する"""
        rows = []
        for (seg_start, seg_end, text), speaker in zip(raw_segments, speaker_labels):
            if rows and rows[-1]['speaker'] == speaker:
                rows[-1]['end'] = seg_end
                rows[-1]['texts'].append(text)
            else:
                rows.append({'start': seg_start, 'end': seg_end, 'speaker': speaker, 'texts': [text]})

        output_file = os.path.join(output_folder, f"{file_name}_speakers.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(['No', '時間帯', '話者', '発言内容'])
        for i, row in enumerate(rows, start=1):
            time_label = f"{self.format_time(row['start'])} - {self.format_time(row['end'])}"
            text_value = '\n'.join(row['texts'])
            sheet.append([i, time_label, row['speaker'], text_value])
        sheet.column_dimensions['A'].width = 6
        sheet.column_dimensions['B'].width = 22
        sheet.column_dimensions['C'].width = 16
        sheet.column_dimensions['D'].width = 100
        for row_cells in sheet.iter_rows(min_row=2, min_col=4, max_col=4):
            row_cells[0].alignment = Alignment(wrap_text=True, vertical='top')

        try:
            workbook.save(output_file)
        except PermissionError:
            raise RuntimeError(
                f"Excelファイルを保存できません。\n{output_file}\n"
                f"このファイルを開いている場合は閉じてから再実行してください。"
            )
        logger.info(f"話者別Excelファイル {output_file} を保存しました。")

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

    def write_docx_output(self, output_folder, file_name, buckets, num_rows, split_interval, duration, low_conf_rows,
                           row_speakers=None, show_speaker_column=False):
        """1分ブロック単位のWord文書を出力"""
        row_speakers = row_speakers or {}
        output_file = os.path.join(output_folder, f"{file_name}_output.docx")
        document = docx.Document()
        document.add_heading(f"文字起こし結果: {file_name}", level=1)
        document.add_paragraph(
            f"元ファイル: {file_name} / 総時間: {self.format_time(duration)} / "
            f"作成日時: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}"
        )
        for idx in range(num_rows):
            start_sec = idx * split_interval
            end_sec = min((idx + 1) * split_interval, duration)
            heading_text = f"[{self.format_time(start_sec)} - {self.format_time(end_sec)}]"
            if show_speaker_column and row_speakers.get(idx):
                heading_text += f" 話者: {'、'.join(row_speakers[idx])}"
            if idx in low_conf_rows:
                heading_text += " ※要確認"
            heading_para = document.add_paragraph()
            heading_run = heading_para.add_run(heading_text)
            heading_run.bold = True
            document.add_paragraph('\n'.join(buckets.get(idx, [])))
            document.add_paragraph("")

        try:
            document.save(output_file)
        except PermissionError:
            raise RuntimeError(
                f"Wordファイルを保存できません。\n{output_file}\n"
                f"このファイルを開いている場合は閉じてから再実行してください。"
            )
        logger.info(f"Wordファイル {output_file} を保存しました。")

    def copy_support_info(self):
        """サポート情報をまとめてクリップボードにコピーする"""
        is_frozen = getattr(sys, 'frozen', False)
        lines = [
            APP_TITLE,
            f"実行形態: {'PyInstaller' if is_frozen else '開発'}",
            f"OS: {platform.platform()}",
            f"使用モデル: {self.selected_model_name}",
            f"検出モデル一覧: {', '.join(self.available_models) if self.available_models else 'なし'}",
            f"単語登録数: {len(self.hotwords_list)} / {MAX_HOTWORDS}",
            f"出力オプション: 分割音声={self.output_split_var.get()}, "
            f"txt={self.output_txt_var.get()}, srt={self.output_srt_var.get()}, "
            f"docx={self.output_docx_var.get()}",
            f"話者分離: {self.diarization_var.get()} (モデル導入={bool(self.diarization_model_path)})",
        ]

        log_dir = os.path.join(get_app_dir(), "logs")
        log_file = os.path.join(log_dir, f"app-{datetime.datetime.now().strftime('%Y%m%d')}.log")
        if os.path.exists(log_file):
            try:
                with open(log_file, "r", encoding="utf-8") as f:
                    log_lines = f.readlines()
                lines.append("")
                lines.append("--- ログ末尾40行 ---")
                lines.extend(line.rstrip("\n") for line in log_lines[-40:])
            except OSError:
                lines.append("(ログ読込不可)")

        info_text = "\n".join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(info_text)
        messagebox.showinfo("サポート情報", "サポート情報をコピーしました。メール等に貼り付けてご利用ください。")

    def run(self):
        self.root.mainloop()


def main():
    if "--selftest" in sys.argv:
        try:
            setup_logging()
        except Exception:
            pass
        sys.exit(run_selftest())
    try:
        setup_logging()
    except Exception:
        pass
    app = AudioTranscriptionApp()
    app.run()


if __name__ == "__main__":
    main()
