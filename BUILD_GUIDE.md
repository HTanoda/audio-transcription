# ビルド手順書 (v1.7.0)

このドキュメントでは、音声文字起こしアプリ v1.7.0 の配布用パッケージをビルドする手順を説明します。

v1.6.0 から話者分離（pyannote.audio + PyTorch CPU）とマイク録音（sounddevice）が追加され、
ビルド構成が大きく変わりました。v1.5.0 以前との主な違い:

| 項目 | v1.5.0 以前 | v1.6.0 |
|------|------------|--------|
| PyInstaller 構成 | onefile（単一EXE） | **onedir**（EXE + `_internal\` フォルダ） |
| ビルド用 venv | `new_env` | **`build_env`**（torch CPU版。new_env は GPU版のため使用しない） |
| UPX | 有効 | **全面無効**（torch の DLL が破損するため） |
| 同梱モデル | models のみ | models + **models_diarization**（話者分離、約32MB） |
| ビルド検証 | 手動起動確認 | **`--selftest` ゲート**（同梱漏れを自動検出） |

## 前提条件

- Python 3.12（`build_env` に導入済み）
- `build_env` に GUI ライブラリ **customtkinter==6.0.0** と、その依存 **darkdetect==0.8.0** を導入済みであること
  （開発機で一度だけ実施。本番機はオフラインのため、これらは EXE に同梱して持ち込む）:
  ```powershell
  build_env\Scripts\python.exe -m pip install customtkinter==6.0.0 darkdetect==0.8.0
  # 開発時に new_env で起動確認する場合は new_env にも同じものを入れる
  new_env\Scripts\python.exe -m pip install customtkinter==6.0.0 darkdetect==0.8.0
  ```
- Windows 10/11
- 約15GBの空きディスク容量
- Inno Setup 6 (`C:\Program Files (x86)\Inno Setup 6\ISCC.exe`)

## ファイル構成

```
D:\whisper\
  ├── audio_transcription.py          # メインアプリ（標準版）
  ├── audio_transcription_turbo.py    # Turbo版ラッパー（識別子のみ差し替え）
  ├── TND_audio_transcription.spec    # 標準版 PyInstaller spec（onedir）
  ├── TND_audio_transcription_turbo.spec  # Turbo版 spec
  ├── TND_AudioTranscription01.ico    # アプリアイコン
  ├── models/                         # Whisper large-v3（約3GB）
  ├── models_turbo/                   # Whisper large-v3-turbo（約1.5GB）
  ├── models_diarization/             # pyannote 話者分離モデル（約32MB）
  ├── build_env/                      # ビルド用 venv（torch 2.11 CPU / pyannote.audio 4.x）
  ├── new_env/                        # 開発用 venv（GPU版torch。ビルドには使わない）
  └── installer/                      # Inno Setup インストーラー一式
      ├── setup_standard.iss
      ├── setup_turbo.iss
      ├── update_standard.iss
      ├── update_turbo.iss
      └── build_installers.ps1
```

> **注:** `setup.py` / `uninstall.py`（および `setup_turbo.py` / `uninstall_turbo.py`）は
> PyInstaller で `setup.exe` / `uninstall.exe` を作る旧方式のインストーラーです。
> **非推奨** であり、現在の配布物ビルドには使用しません。参考としてリポジトリに残置しています。

## spec ファイルの要点（変更時の注意）

ビルドオプションはすべて spec に集約されています（コマンドラインの `--add-data` 等は不要）。

- **onedir 構成**: `EXE(..., exclude_binaries=True, upx=False)` + `COLLECT(..., upx=False)`。
  **UPX は絶対に有効化しないこと**（torch の DLL が壊れる）
- **datas**: onnxruntime / faster_whisper の VAD 資材は `build_env/...` から取得。
  `collect_data_files('pyannote.audio')` で telemetry/config.yaml 等の非 .py データを回収。
  `collect_data_files('customtkinter')` で customtkinter の `assets\`（色テーマ JSON
  `themes\*.json`、部品描画用フォント `CustomTkinter_shapes_font.otf`、Roboto フォント、アイコン）を回収。
  これらは .py ではないため import 解析では拾われず、欠けると customtkinter の import 時
  （テーマ JSON の読込）に失敗して GUI が起動しない
- **hiddenimports**: pyannote は config.yaml のクラス名文字列を importlib で動的解決するため、
  pyannote 系・lightning 系・`scipy._external.array_api_compat` を `collect_submodules` で網羅
- **excludes**:
  - `torchcodec` … build_env では FFmpeg DLL 不整合で import 不能。アプリは音声を
    torch.Tensor で直接渡すため不要
  - `matplotlib` … pyannote の学習用コードからのみ参照。推論では未使用
- **ASIO DLL の除去**: Analysis 直後に `libportaudio64bit-asio.dll` を a.binaries / a.datas
  からフィルタ除去。アプリは WASAPI/MME のみ使用し、sounddevice は環境変数
  `SD_ENABLE_ASIO` 設定時のみ ASIO 版をロードするため動作に影響なし
  （Steinberg ASIO SDK 由来バイナリを配布物に含めないためのライセンス対応）

## オフライン保証（アプリ側実装）

`audio_transcription.py` 冒頭で以下を設定済み。**変更・削除しないこと**:

- `HF_HUB_OFFLINE=1` / `TRANSFORMERS_OFFLINE=1` … モデル読込時のネットワークアクセス遮断
- `PYANNOTE_METRICS_ENABLED=false` … pyannote.audio は既定でテレメトリ
  （otel.pyannote.ai への使用状況送信）が**有効**なため、import 前に明示的に無効化

## ビルド手順

### Step 1: モデルの準備（未取得の場合のみ）

```powershell
cd D:\whisper

# 標準版 Whisper モデル
build_env\Scripts\python.exe -c "from faster_whisper import WhisperModel; WhisperModel('large-v3', device='cpu', compute_type='int8', download_root='models')"

# Turbo版 Whisper モデル
build_env\Scripts\python.exe -c "from faster_whisper import WhisperModel; WhisperModel('large-v3-turbo', device='cpu', compute_type='int8', download_root='models_turbo')"
```

話者分離モデル（`models_diarization/`）は pyannote/speaker-diarization-community-1 の
HuggingFace キャッシュ形式フォルダです。取得には HF トークンが必要なため、既存フォルダを
そのまま使用します（`snapshots/<hash>/config.yaml` が存在することを確認）。

### Step 2: メインアプリのビルド（onedir）

**必ず `build_env` の pyinstaller を使うこと**（new_env は GPU版 torch のため配布物が壊れる）。

```powershell
cd D:\whisper

# 標準版
build_env\Scripts\pyinstaller.exe --noconfirm --clean TND_audio_transcription.spec

# Turbo版
build_env\Scripts\pyinstaller.exe --noconfirm --clean TND_audio_transcription_turbo.spec
```

出力は onedir フォルダになります（単一EXEではない）:

```
dist\TND_audio_transcription\
  ├── TND_audio_transcription.exe   # 本体（起動ブートストラップ、約60MB）
  └── _internal\                    # DLL・ライブラリ一式（約700MB。torch が約320MB）
```

### Step 3: セルフテスト（ビルド検証ゲート・必須）

凍結EXEの同梱漏れを起動だけで検出できる `--selftest` を必ず実行します。
モデルフォルダが EXE の隣に必要なため、一時的にジャンクションを張って実行します:

```powershell
# 標準版
cmd /c mklink /J dist\TND_audio_transcription\models models
cmd /c mklink /J dist\TND_audio_transcription\models_diarization models_diarization
dist\TND_audio_transcription\TND_audio_transcription.exe --selftest
# → 終了コード 0 / logs\selftest_*.log に "RESULT: ALL_OK" が出ること
#   [OK] faster_whisper / [OK] sounddevice / [OK] pyannote の3項目を確認

# 検証後はジャンクションを外す（rmdir はリンクのみ削除、実体は消えない）
cmd /c rmdir dist\TND_audio_transcription\models
cmd /c rmdir dist\TND_audio_transcription\models_diarization
Remove-Item -Recurse -Force dist\TND_audio_transcription\logs
```

Turbo版も同様に（models は `models_turbo` を張る）:

```powershell
cmd /c mklink /J dist\TND_audio_transcription_turbo\models models_turbo
cmd /c mklink /J dist\TND_audio_transcription_turbo\models_diarization models_diarization
dist\TND_audio_transcription_turbo\TND_audio_transcription_turbo.exe --selftest
cmd /c rmdir dist\TND_audio_transcription_turbo\models
cmd /c rmdir dist\TND_audio_transcription_turbo\models_diarization
Remove-Item -Recurse -Force dist\TND_audio_transcription_turbo\logs
```

selftest の3項目:

1. **faster_whisper** … 同梱 Whisper モデルの実体解決（model.bin の存在まで確認）
2. **sounddevice** … PortAudio DLL のロード + デバイス列挙（マイク0台でも成功）
3. **pyannote** … 話者分離モデルのロード + 5秒ダミー波形でのパイプライン実行完走

あわせて ASIO DLL が除去されていることを確認:

```powershell
# 何も出力されなければ OK
Get-ChildItem -Recurse dist\TND_audio_transcription\_internal -Filter "*asio*"
```

customtkinter の色テーマ JSON が同梱されていることを確認（`--selftest` は GUI を起動しないため、
ここで確認する）:

```powershell
# blue.json / dark-blue.json / gold.json / green.json の 4 つが表示されれば OK
Get-ChildItem dist\TND_audio_transcription\_internal\customtkinter\assets\themes -Filter "*.json"
```

### Step 4: 配布用フォルダの作成

Inno Setup の入力ソースとなる、バージョン別の配布用フォルダを組み立てます。
onedir 出力は `app\` サブフォルダに丸ごと格納します。

```powershell
$v = "1.7.0"

# ---- 標準版 ----
New-Item -ItemType Directory -Path "dist\TND_AudioTranscription_v$v\app" -Force

# onedir 出力一式（EXE + _internal）を app\ にコピー
Copy-Item -Recurse "dist\TND_audio_transcription\*" "dist\TND_AudioTranscription_v$v\app\"

# アイコン・README
Copy-Item "TND_AudioTranscription01.ico" "dist\TND_AudioTranscription_v$v\"
Copy-Item "README.txt" "dist\TND_AudioTranscription_v$v\"

# Whisperモデル（シンボリックリンクが実体ファイルに変換される）
Copy-Item -Recurse "models" "dist\TND_AudioTranscription_v$v\models"
# blobsフォルダを削除（snapshotsに実体があるので不要、サイズ半減）
Remove-Item -Recurse -Force "dist\TND_AudioTranscription_v$v\models\models--Systran--faster-whisper-large-v3\blobs"

# 話者分離モデル
Copy-Item -Recurse "models_diarization" "dist\TND_AudioTranscription_v$v\models_diarization"

# ---- Turbo版 ----
New-Item -ItemType Directory -Path "dist\TND_AudioTranscription_turbo_v$v\app" -Force
Copy-Item -Recurse "dist\TND_audio_transcription_turbo\*" "dist\TND_AudioTranscription_turbo_v$v\app\"
Copy-Item "TND_AudioTranscription01.ico" "dist\TND_AudioTranscription_turbo_v$v\"
Copy-Item "README_turbo.txt" "dist\TND_AudioTranscription_turbo_v$v\README.txt"
Copy-Item -Recurse "models_turbo" "dist\TND_AudioTranscription_turbo_v$v\models"
Remove-Item -Recurse -Force "dist\TND_AudioTranscription_turbo_v$v\models\models--mobiuslabsgmbh--faster-whisper-large-v3-turbo\blobs"
if (Test-Path "dist\TND_AudioTranscription_turbo_v$v\models\.locks") { Remove-Item -Recurse -Force "dist\TND_AudioTranscription_turbo_v$v\models\.locks" }
Copy-Item -Recurse "models_diarization" "dist\TND_AudioTranscription_turbo_v$v\models_diarization"
```

完成形（v1.6.0 実測: 標準 約3.7GB / Turbo 約2.3GB）:

```
TND_AudioTranscription_v1.7.0/
  ├── app/                            # onedir 出力（EXE + _internal\、約770MB）
  │   ├── TND_audio_transcription.exe
  │   └── _internal/
  ├── TND_AudioTranscription01.ico
  ├── README.txt
  ├── models/                         # Whisperモデル（約2.9GB）
  └── models_diarization/             # 話者分離モデル（約32MB）
```

> **注:** 同梱ライブラリのライセンス全文 `THIRD_PARTY_LICENSES.txt` は `{app}`（EXE と同じフォルダ）に
> 同梱します（BSD/MIT の「バイナリ配布時にライセンス文を添付する」条件のため）。
> インストーラーの [Files] がリポジトリ直下の `THIRD_PARTY_LICENSES.txt` を直接参照する
> （`Source: "..\THIRD_PARTY_LICENSES.txt"`、.iss のフォルダ基準）ので、上の配布用フォルダへのコピーは不要です。
> アプリ内では画面最下段の「ライセンス情報」→「全文を開く」で表示できます。

### Step 5: Inno Setup によるインストーラー作成

Step 4 で組み立てた配布用フォルダを入力として、`installer\build_installers.ps1` を実行します。
標準版・Turbo版それぞれのフルインストーラーと差分更新インストーラー、計4本を一括ビルドします。

```powershell
cd installer
.\build_installers.ps1 -Version 1.7.0
```

既定では以下のフォルダをソースとして参照します（`-StandardDir` / `-TurboDir` で明示指定も可能）:

- 標準版: `..\dist\TND_AudioTranscription_v<Version>`
- Turbo版: `..\dist\TND_AudioTranscription_turbo_v<Version>`

ビルドが成功すると、`dist\` 直下に次の4本が生成されます。

```
dist\
  ├── TND_AudioTranscription-setup-1.7.0.exe          # 標準版フル（約3.2GB）
  ├── TND_AudioTranscription-update-1.7.0.exe         # 標準版差分更新（約234MB）
  ├── TND_AudioTranscription_turbo-setup-1.7.0.exe     # Turbo版フル（約1.8GB）
  └── TND_AudioTranscription_turbo-update-1.7.0.exe    # Turbo版差分更新（約234MB）
```

インストーラーの [Files] 構成（v1.6.0〜）:

- **setup（フル）**: `app\*`（本体一式）+ README + アイコン + `models\*` + `models_diarization\*`
- **update（差分）**: `app\*` + README + `models_diarization\*`。
  **models（3GB級）は含めない**のが従来からの方針。models_diarization は
  v1.6.0 の新規追加物のため update にも含める（v1.5.0 からの更新で話者分離を使えるように）。
  update 版の本体が約115MB→234MB に増えたのは onedir 化で torch 等を同梱するため

### Step 6: サイレントインストール実機検証（出荷ゲート）

Turbo版フルインストーラーで install → selftest → 起動 → uninstall を検証します。

> **注意:** `/VERYSILENT` 等のスイッチは **PowerShell から実行すること**。
> Git Bash 経由だと `/VERYSILENT` がパスに変換されてサイレントにならず、
> 通常のウィザードが起動してしまう（v1.6.0 検証時に実際に発生）。

```powershell
# サイレントインストール
Start-Process -FilePath "dist\TND_AudioTranscription_turbo-setup-1.7.0.exe" `
  -ArgumentList "/VERYSILENT","/SUPPRESSMSGBOXES","/NORESTART" -Wait

# インストール先でセルフテスト（終了コード 0 / RESULT: ALL_OK を確認）
& "$env:LOCALAPPDATA\TND_AudioTranscription_turbo\TND_audio_transcription_turbo.exe" --selftest

# 通常起動スモーク（タイトルバー確認後、手動または Stop-Process で終了）

# サイレントアンインストール
Start-Process -FilePath "$env:LOCALAPPDATA\TND_AudioTranscription_turbo\unins000.exe" `
  -ArgumentList "/VERYSILENT","/SUPPRESSMSGBOXES","/NORESTART" -Wait
```

## ユーザー向けインストール手順

1. `TND_AudioTranscription-setup-<version>.exe`（Turbo版は `TND_AudioTranscription_turbo-setup-<version>.exe`）をダブルクリック
2. インストール先を確認（デフォルト: 標準版 `C:\Users\<ユーザー名>\AppData\Local\TND_AudioTranscription`、Turbo版は末尾 `_turbo`）
3. 「インストール」ボタンをクリック
4. デスクトップに作成されたショートカットからアプリを起動

既存インストールを更新する場合は、対応する `*-update-<version>.exe` を実行してください
（本体一式・README.txt・話者分離モデルが更新され、Whisperモデル（models）や
ユーザー設定は変更されません）。

## アンインストール手順

1. 「設定」→「アプリ」→「インストールされているアプリ」
2. 「TND AI議事録アプリ」（Turbo版は「TND AI議事録アプリ (Turbo版)」）を選択
3. 「アンインストール」をクリック

## トラブルシューティング

### selftest が pyannote 項目で ModuleNotFoundError

pyannote 系の動的 import が spec の hiddenimports から漏れています。
エラーに出たモジュールの親パッケージを spec の `collect_submodules` 対象に追加して
再ビルド → `--selftest` 再実行、を繰り返して潰します
（v1.6.0 では `scipy._external.array_api_compat` がこのパターンで追加された）。

### selftest が pyannote 項目で FileNotFoundError (config.yaml 等)

非 .py のデータファイルが同梱されていません。`collect_data_files('<パッケージ>')` を
spec の datas に追加します。

### EXEが起動しない

spec の `console=False` を一時的に `True` にしてビルドし、コンソールのエラーを確認します。

### ビルド時に `silero_vad_v6.onnx` が見つからない

faster-whisperのバージョンによってファイル名が異なります。以下で確認:

```powershell
dir build_env\Lib\site-packages\faster_whisper\assets\
```

## Turbo版と標準版の併存の仕組み

`audio_transcription_turbo.py` は標準版モジュールをimportした後、アプリ名（タイトルバー）と
ライセンス表記のみ turbo 用に差し替えるだけの薄いラッパーです。

インストール先・レジストリキー・表示名といった**インストール時の識別子分離**は
Inno Setup スクリプト側で行います。`setup_standard.iss` と `setup_turbo.iss`
（および `update_standard.iss` / `update_turbo.iss`）はそれぞれ別の `AppId`（固定GUID）、
`DefaultDirName`（`{localappdata}\TND_AudioTranscription` / `..._turbo`）、`AppName`
（表示名）を持つため、標準版とTurbo版は同一PCに衝突なく併存インストールできます。
