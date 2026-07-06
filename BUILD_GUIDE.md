# ビルド手順書 (v1.4.0)

このドキュメントでは、音声文字起こしアプリ v1.4.0 の配布用パッケージをビルドする手順を説明します。

## 前提条件

- Python 3.10〜3.12
- Windows 10/11
- 約10GBの空きディスク容量

## ファイル構成

```
D:\whisper\
  ├── audio_transcription.py    # メインアプリ
  ├── setup.py                  # インストーラー
  ├── uninstall.py              # アンインストーラー
  ├── TND_AudioTranscription01.ico  # アプリアイコン
  ├── models/                   # Whisperモデル（約3GB）
  └── new_env/                  # Python仮想環境
```

## ビルド手順

### Step 1: 仮想環境の準備

```powershell
cd D:\whisper

# 仮想環境を有効化
new_env\Scripts\Activate.ps1

# 必要に応じてパッケージを確認
pip list
```

### Step 2: モデルの準備

modelsフォルダが存在しない場合は、以下でダウンロード：

```powershell
python -c "from faster_whisper import WhisperModel; WhisperModel('large-v3', device='cpu', compute_type='int8', download_root='models')"
```

### Step 3: メインアプリのビルド

```powershell
pyinstaller --onefile ^
  --add-data "new_env\Lib\site-packages\onnxruntime;onnxruntime" ^
  --add-data "new_env\Lib\site-packages\faster_whisper\vad.py;faster_whisper" ^
  --add-data "new_env\Lib\site-packages\faster_whisper\assets\silero_vad_v6.onnx;faster_whisper\assets" ^
  --noconsole ^
  --icon "TND_AudioTranscription01.ico" ^
  --name "TND_audio_transcription" ^
  audio_transcription.py
```

**1行バージョン:**
```powershell
pyinstaller --onefile --add-data "new_env\Lib\site-packages\onnxruntime;onnxruntime" --add-data "new_env\Lib\site-packages\faster_whisper\vad.py;faster_whisper" --add-data "new_env\Lib\site-packages\faster_whisper\assets\silero_vad_v6.onnx;faster_whisper\assets" --noconsole --icon "TND_AudioTranscription01.ico" --name "TND_audio_transcription" audio_transcription.py
```

### Step 4: インストーラーのビルド

```powershell
pyinstaller --onefile --noconsole --name "setup" setup.py
```

### Step 5: アンインストーラーのビルド

```powershell
pyinstaller --onefile --noconsole --name "uninstall" uninstall.py
```

### Step 6: 配布用フォルダの作成

```powershell
# 配布用フォルダを作成
New-Item -ItemType Directory -Path "dist\TND_AudioTranscription_v1.4.0" -Force

# ファイルをコピー
Copy-Item "dist\TND_audio_transcription.exe" "dist\TND_AudioTranscription_v1.4.0\"
Copy-Item "dist\setup.exe" "dist\TND_AudioTranscription_v1.4.0\"
Copy-Item "dist\uninstall.exe" "dist\TND_AudioTranscription_v1.4.0\"

# modelsフォルダをコピー（シンボリックリンクが実体ファイルに変換される）
Copy-Item -Recurse "models" "dist\TND_AudioTranscription_v1.4.0\"

# blobsフォルダを削除（snapshotsに実体があるので不要、サイズ半減）
Remove-Item -Recurse -Force "dist\TND_AudioTranscription_v1.4.0\models\models--Systran--faster-whisper-large-v3\blobs"

# アイコンファイルをコピー
Copy-Item "TND_AudioTranscription01.ico" "dist\TND_AudioTranscription_v1.4.0\"

# README.txtをコピー（存在する場合）
Copy-Item "README.txt" "dist\TND_AudioTranscription_v1.4.0\" -ErrorAction SilentlyContinue
```

**サイズ確認（約2.8GBになっていることを確認）:**
```powershell
(Get-ChildItem -Recurse "dist\TND_AudioTranscription_v1.4.0\models" | Measure-Object -Property Length -Sum).Sum / 1GB
```

## 配布用パッケージの内容

```
TND_AudioTranscription_v1.4.0/
  ├── setup.exe                       # ← ユーザーはこれを実行
  ├── TND_audio_transcription.exe     # メインアプリ
  ├── uninstall.exe                   # アンインストーラー
  ├── TND_AudioTranscription01.ico    # アプリアイコン
  ├── README.txt                      # ユーザー向け説明書（任意）
  └── models/                         # Whisperモデル（約3GB）
```

> **注:** ライセンス情報はアプリ内メニュー「ヘルプ」→「ライセンス情報」から確認できるため、配布パッケージには含めません。

## ユーザー向けインストール手順

1. 配布フォルダを展開
2. `setup.exe` をダブルクリック
3. インストール先を確認（デフォルト: `C:\Users\<ユーザー名>\AppData\Local\TND_AudioTranscription`）
4. 「インストール」ボタンをクリック
5. デスクトップに作成されたショートカットからアプリを起動

## アンインストール手順

以下のいずれかの方法でアンインストールできます：

### 方法1: コントロールパネルから
1. 「設定」→「アプリ」→「インストールされているアプリ」
2. 「TND AI議事録アプリ」を選択
3. 「アンインストール」をクリック

### 方法2: 直接実行
1. インストールフォルダ内の `uninstall.exe` を実行

## トラブルシューティング

### ビルド時に `silero_vad_v6.onnx` が見つからない

faster-whisperのバージョンによってファイル名が異なります。以下で確認：

```powershell
dir new_env\Lib\site-packages\faster_whisper\assets\
```

### EXEが起動しない

コンソール付きでビルドしてエラーを確認：

```powershell
# --noconsole を外してビルド
pyinstaller --onefile ... --name "TND_audio_transcription_debug" audio_transcription.py
```

## Turbo版のビルド

Turbo版は標準版のソース（`audio_transcription.py` / `setup.py` / `uninstall.py`）を
そのままimportし、識別子（インストール先・レジストリキー・表示名・EXE名）と
ライセンス表記のみ差し替える薄いラッパー（`audio_transcription_turbo.py` /
`setup_turbo.py` / `uninstall_turbo.py`）として実装されています。そのため、
標準版と同じ同梱データ（onnxruntime、faster_whisperのVAD関連ファイル）が必要です。

### Step 1: Turbo版モデルの取得

```powershell
python -c "from faster_whisper import WhisperModel; WhisperModel('large-v3-turbo', device='cpu', compute_type='int8', download_root='models_turbo')"
```

### Step 2: Turbo版EXEのビルド

ラッパーが標準版ソースをimportするだけのため、`--add-data` は標準版と同じ群を指定します。

```powershell
pyinstaller --onefile ^
  --add-data "new_env\Lib\site-packages\onnxruntime;onnxruntime" ^
  --add-data "new_env\Lib\site-packages\faster_whisper\vad.py;faster_whisper" ^
  --add-data "new_env\Lib\site-packages\faster_whisper\assets\silero_vad_v6.onnx;faster_whisper\assets" ^
  --noconsole ^
  --icon "TND_AudioTranscription01.ico" ^
  --name "TND_audio_transcription_turbo" ^
  audio_transcription_turbo.py
```

```powershell
pyinstaller --onefile --noconsole --name "setup_turbo" setup_turbo.py
```

```powershell
pyinstaller --onefile --noconsole --name "uninstall_turbo" uninstall_turbo.py
```

### Step 3: 配布用フォルダの作成

```powershell
New-Item -ItemType Directory -Path "dist\TND_AudioTranscription_turbo_v1.4.0" -Force

Copy-Item "dist\TND_audio_transcription_turbo.exe" "dist\TND_AudioTranscription_turbo_v1.4.0\"
Copy-Item "dist\setup_turbo.exe" "dist\TND_AudioTranscription_turbo_v1.4.0\setup.exe"
Copy-Item "dist\uninstall_turbo.exe" "dist\TND_AudioTranscription_turbo_v1.4.0\uninstall.exe"

# models_turboフォルダを models としてコピー
Copy-Item -Recurse "models_turbo" "dist\TND_AudioTranscription_turbo_v1.4.0\models"

# blobsフォルダを削除（snapshotsに実体があるので不要）
Remove-Item -Recurse -Force "dist\TND_AudioTranscription_turbo_v1.4.0\models\models--mobiuslabsgmbh--faster-whisper-large-v3-turbo\blobs"

Copy-Item "TND_AudioTranscription01.ico" "dist\TND_AudioTranscription_turbo_v1.4.0\"
Copy-Item "README.txt" "dist\TND_AudioTranscription_turbo_v1.4.0\" -ErrorAction SilentlyContinue
```

### 配布フォルダ構成

```
TND_AudioTranscription_turbo_v1.4.0/
  ├── setup.exe                             # ← ユーザーはこれを実行（Turbo版インストーラー）
  ├── TND_audio_transcription_turbo.exe     # Turbo版メインアプリ
  ├── uninstall.exe                         # Turbo版アンインストーラー
  ├── TND_AudioTranscription01.ico          # アプリアイコン
  ├── README.txt                            # ユーザー向け説明書（任意）
  └── models/                               # Whisper large-v3-turbo モデル（約1.5GB）
```

### 標準版との併存の仕組み

`audio_transcription_turbo.py` / `setup_turbo.py` は標準版モジュールをimportした後、
`APP_NAME`（インストール先フォルダ名・レジストリキーに使用）、`APP_DISPLAY_NAME`
（表示名）、`APP_EXE_NAME`（本体EXE名）を turbo 用の値に上書きするだけで、
標準版と衝突せず同一PCに併存インストールできるようにしています。
