# Audio Transcription Tool

音声ファイルを自動で文字起こしするGUIアプリケーションです。

## 特徴

- **対応フォーマット**: WAV, MP3, M4A, MP4
- **高精度文字起こし**: faster-whisper (large-v3モデル) を使用
- **自動分割処理**: 長い音声ファイルを1分ごとに分割して処理
- **Excel出力**: 文字起こし結果をExcelファイル(.xlsx)に出力
- **GUI対応**: tkinterによるシンプルなファイル選択インターフェース

## 必要要件

- Python 3.9以上
- 約3GBの空きディスク容量（モデルファイル用）

## インストール

```bash
# リポジトリをクローン
git clone https://github.com/yourusername/audio-transcription.git
cd audio-transcription

# 依存パッケージをインストール
pip install -r requirements.txt
```

## 使い方

```bash
python audio_transcription.py
```

1. 起動するとファイル選択ダイアログが表示されます
2. 文字起こしする音声ファイルを選択してください
3. 出力先フォルダを選択してください
4. 処理が完了すると、選択したフォルダに以下が出力されます：
   - 分割された音声ファイル（1分ごと）
   - 文字起こし結果のExcelファイル（`*_output.xlsx`）

## 出力形式

Excelファイルには以下の列が含まれます：

| No | 音声ファイル | 変換結果 |
|----|-------------|---------|
| 0  | ファイルパス（リンク付き）| 文字起こしテキスト |
| 1  | ... | ... |

## 依存ライブラリ

| ライブラリ | 用途 | ライセンス |
|-----------|------|----------|
| [faster-whisper](https://github.com/SYSTRAN/faster-whisper) | 音声認識 | MIT |
| [moviepy](https://github.com/Zulko/moviepy) | 音声ファイル処理 | MIT |
| [pandas](https://github.com/pandas-dev/pandas) | データ処理 | BSD-3-Clause |
| [openpyxl](https://openpyxl.readthedocs.io/) | Excel出力 | MIT |

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
|-------|-------|------|------|
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

## ライセンス

このプロジェクトは [MIT License](LICENSE) の下で公開されています。

## 貢献

Issue や Pull Request を歓迎します。

## 謝辞

- [OpenAI Whisper](https://github.com/openai/whisper) - 音声認識モデル
- [SYSTRAN/faster-whisper](https://github.com/SYSTRAN/faster-whisper) - 高速推論実装
