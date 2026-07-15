# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

datas = [
    ('build_env/Lib/site-packages/onnxruntime', 'onnxruntime'),
    ('build_env/Lib/site-packages/faster_whisper/vad.py', 'faster_whisper'),
    ('build_env/Lib/site-packages/faster_whisper/assets/silero_vad_v6.onnx', 'faster_whisper/assets'),
]
# pyannote.audio.telemetry.config.yaml 等の非.pyデータ（collect_submodulesでは拾えない）
datas += collect_data_files('pyannote.audio')

# pyannote.audio は config.yaml に書かれたクラス名（文字列）を importlib で動的に
# 解決するため、静的解析だけでは辿れないモジュールがある。関連パッケージは
# サブモジュールを丸ごと hiddenimports に加える。
hiddenimports = []
for pkg in (
    'pyannote.audio',
    'pyannote.pipeline',
    'pyannote.core',
    'pyannote.database',
    'pytorch_lightning',
    'lightning',
    'lightning_fabric',
    'torch_audiomentations',
    'torch_pitch_shift',
    'asteroid_filterbanks',
    'julius',
    # scipy 同梱の array API 互換レイヤーは配列の型名から importlib で
    # バックエンド（numpy/torch/...）を動的解決するため静的解析で漏れる。
    'scipy._external.array_api_compat',
):
    hiddenimports += collect_submodules(pkg)

# torchcodec は build_env にインストールされているが FFmpeg DLL 不整合で
# import 不能（build_env で検証済み）。本アプリは pyannote.audio に waveform を
# torch.Tensor で直接渡しており torchcodec 経由の音声デコードを使わないため、
# 同梱から除外してサイズを削減する。
#
# matplotlib は pyannote.audio.tasks（学習用コード。推論では未使用）と
# pyannote.metrics.plot からのみ参照される。本アプリは推論（Pipeline.apply）
# のみを使い学習コードは呼ばないため除外する（標準版ビルドで動作確認済みの方針を踏襲）。
excludes = [
    'torchcodec',
    'matplotlib',
]

a = Analysis(
    ['audio_transcription_turbo.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)

# ASIO 版 PortAudio DLL（libportaudio64bit-asio.dll）は配布物から除外する。
# 本アプリは既定の WASAPI/MME 系のみを使用し、sounddevice.py は環境変数
# SD_ENABLE_ASIO が設定された場合にのみ ASIO 版 DLL をロードするため、
# 除去しても動作に影響しない（--selftest の sounddevice 項目で確認）。
# Steinberg ASIO SDK 由来のバイナリを配布物に含めないためのライセンス対応。
a.binaries = [b for b in a.binaries if 'libportaudio64bit-asio' not in b[0].lower()]
a.datas = [d for d in a.datas if 'libportaudio64bit-asio' not in d[0].lower()]

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='TND_audio_transcription_turbo',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['TND_AudioTranscription01.ico'],
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='TND_audio_transcription_turbo',
)
