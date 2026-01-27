# Audio Transcription Tool

音声ファイルを自動で文字起こしするGUIアプリケーションです。

## 特徴

* **対応フォーマット**: WAV, MP3, M4A, MP4
* **高精度文字起こし**: faster-whisper (large-v3モデル) を使用
* **自動分割処理**: 長い音声ファイルを1分ごとに分割して処理
* **Excel出力**: 文字起こし結果をExcelファイル(.xlsx)に出力
* **GUI対応**: tkinterによるシンプルなファイル選択インターフェース
* **プログレスバー**: 処理状況をリアルタイムで表示（v1.2.0〜）
* **自動フォルダ表示**: 処理完了後に出力フォルダを自動で開く（v1.2.0〜）

## 更新履歴

### v1.2.0 (2026-01-27)
- **新機能**: プログレスバーによる処理状況の可視化
- **新機能**: 処理完了後に出力フォルダを自動で開く
- **改善**: モデル読み込みをループ外に移動し、処理速度を大幅に向上
- **改善**: GUIウィンドウのリサイズに対応
- **修正**: UIフリーズを防ぐため、処理を別スレッドで実行

### v1.1.1 (2026-01-26)
- 初回リリース

## 必要要件

* Python 3.10〜3.12（3.13以降は非推奨）
* 約3GBの空きディスク容量（モデルファイル用）

## インストール

```bash
# リポジトリをクローン
git clone https://github.com/HTanoda/audio-transcription.git
cd audio-transcription

# 仮想環境を作成（推奨）
python -m venv venv

# 仮想環境を有効化
# Windows (PowerShell)
venv\Scripts\Activate.ps1
# Windows (コマンドプロンプト)
venv\Scripts\activate.bat
# macOS/Linux
source venv/bin/activate

# 依存パッケージをインストール
pip install -r requirements.txt
```

## 使い方

```bash
python audio_transcription.py
```

1. 起動するとGUIウィンドウが表示されます
2. 「入力ファイル」の「選択」ボタンで音声ファイルを選択
3. 「出力フォルダ」の「選択」ボタンで出力先フォルダを選択
4. 「文字起こし開始」ボタンをクリック
5. プログレスバーで処理状況を確認
6. 処理完了後、自動で出力フォルダが開きます

## 出力形式

Excelファイルには以下の列が含まれます：

| No | 音声ファイル | 変換結果 |
| --- | --- | --- |
| 0 | ファイルパス（リンク付き） | 文字起こしテキスト |
| 1 | ... | ... |

## EXE化（配布用）

PyInstallerを使用してスタンドアロンの実行ファイルを作成できます。

```bash
# PyInstallerをインストール
pip install pyinstaller

# モデルをダウンロード（初回のみ）
python -c "from faster_whisper import WhisperModel; WhisperModel('large-v3', device='cpu', compute_type='int8', download_root='models')"

# EXEをビルド
pyinstaller --onefile \
  --add-data "models;models" \
  --add-data "venv\Lib\site-packages\onnxruntime;onnxruntime" \
  --add-data "venv\Lib\site-packages\faster_whisper\vad.py;faster_whisper" \
  --add-data "venv\Lib\site-packages\faster_whisper\assets\silero_vad_v6.onnx;faster_whisper\assets" \
  --copy-metadata imageio \
  --copy-metadata imageio-ffmpeg \
  --noconsole \
  --name "音声文字起こし" \
  audio_transcription.py
```

生成されたEXEは `dist` フォルダに出力されます。

## 依存ライブラリ

| ライブラリ | バージョン | 用途 | ライセンス |
| --- | --- | --- | --- |
| [faster-whisper](https://github.com/SYSTRAN/faster-whisper) | - | 音声認識 | MIT |
| [moviepy](https://github.com/Zulko/moviepy) | 1.0.3 | 音声ファイル処理 | MIT |
| [pandas](https://github.com/pandas-dev/pandas) | - | データ処理 | BSD-3-Clause |
| [openpyxl](https://openpyxl.readthedocs.io/) | - | Excel出力 | MIT |

> **注意**: moviepyは2.x系ではなく1.0.3を使用してください。2.x系では`moviepy.editor`が廃止されています。

## 設定のカスタマイズ

`audio_transcription.py` 内の以下の設定を変更できます：

```python
# 分割間隔（秒）- デフォルト: 60秒
split_interval = 1 * 60

# Whisperモデル設定
model = WhisperModel("large-v3", device="cpu", compute_type="int8")

# 初期プロンプト（認識精度向上のためのヒント）
initial_prompt="高島宗一郎です。こんにちは、今日はよろしくお願いします。"
```

### モデルサイズの変更

処理速度と精度のトレードオフに応じてモデルを変更できます：

| モデル | サイズ | 精度 | 速度 |
| --- | --- | --- | --- |
| tiny | 78.2MB | 低 | 超高速 |
| base | 148MB | 低 | 高速 |
| small | 486MB | バランス | バランス |
| medium | 1.53GB | 高精度 | 少し遅い |
| large-v3 | 3.09GB | 最高精度 | 遅い |

## GPU対応

CUDAが利用可能な環境では、以下のように変更することでGPU処理が可能です：

```python
model = WhisperModel("large-v3", device="cuda", compute_type="float16")
```

## トラブルシューティング

### `No module named 'moviepy.editor'` エラー

moviepy 2.x がインストールされています。1.0.3にダウングレードしてください：

```bash
pip uninstall moviepy -y
pip install moviepy==1.0.3
```

### PyInstallerでビルドしたEXEが起動しない

コンソール付きでビルドしてエラーを確認してください：

```bash
# --noconsole を外してビルド
pyinstaller --onefile ... audio_transcription.py
```

### Pythonバージョンの問題

Python 3.13以降では依存パッケージのビルドに問題が発生する場合があります。Python 3.10〜3.12の使用を推奨します。

## ライセンス

このプロジェクトは [MIT License](LICENSE) の下で公開されています。

## 貢献

Issue や Pull Request を歓迎します。

## 謝辞

* [OpenAI Whisper](https://github.com/openai/whisper) - 音声認識モデル
* [SYSTRAN/faster-whisper](https://github.com/SYSTRAN/faster-whisper) - 高速推論実装
