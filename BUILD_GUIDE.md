# ビルド手順書 (v1.4.0)

このドキュメントでは、音声文字起こしアプリ v1.4.0 の配布用パッケージをビルドする手順を説明します。

## 前提条件

- Python 3.10〜3.12
- Windows 10/11
- 約10GBの空きディスク容量
- Inno Setup 6 (`C:\Program Files (x86)\Inno Setup 6\ISCC.exe`)

## ファイル構成

```
D:\whisper\
  ├── audio_transcription.py    # メインアプリ
  ├── setup.py                  # 旧インストーラー（非推奨、参考用）
  ├── uninstall.py               # 旧アンインストーラー（非推奨、参考用）
  ├── TND_AudioTranscription01.ico  # アプリアイコン
  ├── models/                   # Whisperモデル（約3GB）
  ├── new_env/                  # Python仮想環境
  └── installer/                # Inno Setup インストーラー一式
      ├── setup_standard.iss
      ├── setup_turbo.iss
      ├── update_standard.iss
      ├── update_turbo.iss
      └── build_installers.ps1
```

> **注:** `setup.py` / `uninstall.py`（および `setup_turbo.py` / `uninstall_turbo.py`）は
> PyInstaller で `setup.exe` / `uninstall.exe` を作る旧方式のインストーラーです。
> **非推奨** であり、現在の配布物ビルドには使用しません。参考としてリポジトリに残置しています。
> 現在の配布物は `installer/` 以下の Inno Setup スクリプトでビルドします（Step 5 参照）。

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

### Step 4: 配布用フォルダの作成

Inno Setup の入力ソースとなる、バージョン別の配布用フォルダを組み立てます。
（`setup.exe` / `uninstall.exe` のコピーは不要になったため行いません。インストーラー本体は Step 5 で別途生成します。）

```powershell
# 配布用フォルダを作成
New-Item -ItemType Directory -Path "dist\TND_AudioTranscription_v1.4.0" -Force

# 本体EXEをコピー
Copy-Item "dist\TND_audio_transcription.exe" "dist\TND_AudioTranscription_v1.4.0\"

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

### Step 5: Inno Setup によるインストーラー作成

Step 4 で組み立てた配布用フォルダを入力として、`installer\build_installers.ps1` を実行します。
標準版・Turbo版それぞれのフルインストーラーと差分更新インストーラー、計4本を一括ビルドします。

```powershell
cd installer
.\build_installers.ps1 -Version 1.4.0
```

既定では以下のフォルダをソースとして参照します（`-StandardDir` / `-TurboDir` で明示指定も可能）:

- 標準版: `..\dist\TND_AudioTranscription_v<Version>`
- Turbo版: `..\dist\TND_AudioTranscription_turbo_v<Version>`

ビルドが成功すると、`dist\` 直下に次の4本が生成されます。

```
dist\
  ├── TND_AudioTranscription-setup-1.4.0.exe          # 標準版フルインストーラー
  ├── TND_AudioTranscription-update-1.4.0.exe         # 標準版差分更新インストーラー
  ├── TND_AudioTranscription_turbo-setup-1.4.0.exe     # Turbo版フルインストーラー
  └── TND_AudioTranscription_turbo-update-1.4.0.exe    # Turbo版差分更新インストーラー
```

## 配布用パッケージの内容

`dist\TND_AudioTranscription_v1.4.0\`（Step 5 の入力ソース）:

```
TND_AudioTranscription_v1.4.0/
  ├── TND_audio_transcription.exe     # メインアプリ
  ├── TND_AudioTranscription01.ico    # アプリアイコン
  ├── README.txt                      # ユーザー向け説明書（任意）
  └── models/                         # Whisperモデル（約3GB）
```

> **注:** ライセンス情報はアプリ内メニュー「ヘルプ」→「ライセンス情報」から確認できるため、配布パッケージには含めません。

エンドユーザーに配布するのは、上記フォルダそのものではなく Step 5 で生成される
`TND_AudioTranscription-setup-<version>.exe`（新規インストール用）または
`TND_AudioTranscription-update-<version>.exe`（既存インストールの差分更新用）です。

## ユーザー向けインストール手順

1. `TND_AudioTranscription-setup-<version>.exe`（Turbo版は `TND_AudioTranscription_turbo-setup-<version>.exe`）をダブルクリック
2. インストール先を確認（デフォルト: 標準版 `C:\Users\<ユーザー名>\AppData\Local\TND_AudioTranscription`、Turbo版は末尾 `_turbo`）
3. 「インストール」ボタンをクリック
4. デスクトップに作成されたショートカットからアプリを起動

既存インストールを更新する場合は、対応する `*-update-<version>.exe` を実行してください
（本体EXEとREADME.txtのみが上書きされ、models フォルダやユーザー設定は変更されません）。

## アンインストール手順

1. 「設定」→「アプリ」→「インストールされているアプリ」
2. 「TND AI議事録アプリ」（Turbo版は「TND AI議事録アプリ (Turbo版)」）を選択
3. 「アンインストール」をクリック

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

Turbo版は標準版のソース（`audio_transcription.py`）をそのままimportし、
アプリ名（タイトルバー）とライセンス表記のみ差し替える薄いラッパー
（`audio_transcription_turbo.py`）として実装されています。そのため、
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

> Turbo版インストーラー（`setup_turbo.exe` / `uninstall_turbo.exe`）の PyInstaller ビルドは
> 旧方式のため不要です。Turbo版インストーラーは標準版と同じ Step 5
> （`installer\build_installers.ps1`）で `setup_turbo.iss` / `update_turbo.iss` として
> まとめてビルドされます。

### Step 3: 配布用フォルダの作成

```powershell
New-Item -ItemType Directory -Path "dist\TND_AudioTranscription_turbo_v1.4.0" -Force

Copy-Item "dist\TND_audio_transcription_turbo.exe" "dist\TND_AudioTranscription_turbo_v1.4.0\"

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
  ├── TND_audio_transcription_turbo.exe     # Turbo版メインアプリ
  ├── TND_AudioTranscription01.ico          # アプリアイコン
  ├── README.txt                            # ユーザー向け説明書（任意）
  └── models/                               # Whisper large-v3-turbo モデル（約1.5GB）
```

このフォルダが Step 5（`installer\build_installers.ps1` の `-TurboDir`）の入力ソースになります。

### 標準版との併存の仕組み

`audio_transcription_turbo.py` は標準版モジュールをimportした後、アプリ名（タイトルバー）と
ライセンス表記のみ turbo 用に差し替えるだけの薄いラッパーです。

インストール先・レジストリキー・表示名といった**インストール時の識別子分離**は、
旧方式（`setup_turbo.py` の `APP_NAME` / `APP_DISPLAY_NAME` / `APP_EXE_NAME` 上書き）ではなく、
現在は Inno Setup スクリプト側で行います。`setup_standard.iss` と `setup_turbo.iss`
（および `update_standard.iss` / `update_turbo.iss`）はそれぞれ別の `AppId`（固定GUID）、
`DefaultDirName`（`{localappdata}\TND_AudioTranscription` / `..._turbo`）、`AppName`
（表示名）を持つため、標準版とTurbo版は同一PCに衝突なく併存インストールできます。
