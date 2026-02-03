# TND AI議事録アプリ (audio-transcription)

会議の音声ファイルをAIが自動で文字起こしするアプリケーションです。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

## 📌 特徴

- **かんたん操作** - ファイルを選んでボタンを押すだけ
- **高精度な文字起こし** - 最新のAI音声認識モデル（Whisper large-v3）を使用
- **Excel出力** - 文字起こし結果をExcelファイルで出力
- **オフライン動作** - インターネット接続なしで使用可能
- **プログレスバー** - 処理状況をリアルタイムで確認

---

## 💻 動作環境

| 項目 | 要件 |
|------|------|
| OS | Windows 10 / 11 |
| メモリ | 8GB以上推奨 |
| ストレージ | 約4GBの空き容量 |

---

## 📥 インストール方法（一般ユーザー向け）

### Step 1: ダウンロード

[Releases](https://github.com/HTanoda/audio-transcription/releases) から最新版のZIPファイルをダウンロードします。

### Step 2: ZIPファイルを展開

ダウンロードした `TND_AudioTranscription_v1.2.1.zip` を右クリックし、「すべて展開」を選択します。

### Step 3: セットアップを実行

展開したフォルダ内の **`setup.exe`** をダブルクリックします。

1. インストール先を確認します（通常は変更不要です）
2. 「デスクトップにショートカットを作成」にチェックが入っていることを確認
3. **「インストール」** ボタンをクリック
4. モデルファイル（約3GB）のコピーに数分かかります

### Step 4: 完了

デスクトップに **「TND AI議事録アプリ」** のショートカットが作成されます。

---

## 🎯 使い方

1. デスクトップの **「TND AI議事録アプリ」** アイコンをダブルクリック
2. 「入力ファイル」の **「選択」** ボタンで音声ファイルを選択
3. 「出力フォルダ」の **「選択」** ボタンで保存先を選択
4. **「文字起こし開始」** ボタンをクリック
5. 処理完了後、自動的に出力フォルダが開きます

**対応フォーマット:** WAV, MP3, M4A, MP4

**処理時間の目安:** 音声1分あたり約1〜3分（PCの性能により異なります）

---

## 📊 出力ファイル

| ファイル | 内容 |
|---------|------|
| `○○_output.xlsx` | 文字起こし結果（Excel形式） |
| `○○_0.wav` など | 分割された音声ファイル（1分ごと） |

---

## 🗑️ アンインストール方法

### 方法1: Windowsの設定から（推奨）

1. **「設定」** → **「アプリ」** → **「インストールされているアプリ」**
2. 「TND AI議事録アプリ」を探す
3. **「アンインストール」** をクリック

### 方法2: アンインストーラーを直接実行

インストールフォルダ内の `uninstall.exe` をダブルクリック

---

## 🔧 開発者向け情報

### リポジトリ構成

```
audio-transcription/
  ├── audio_transcription.py    # メインアプリケーション
  ├── setup.py                  # インストーラー
  ├── uninstall.py              # アンインストーラー
  ├── requirements.txt          # 依存パッケージ
  ├── BUILD_GUIDE.md            # ビルド手順書
  ├── README.md                 # このファイル
  ├── LICENSE                   # MITライセンス
  └── THIRD_PARTY_LICENSES.txt  # サードパーティライセンス
```

### 開発環境のセットアップ

```bash
# リポジトリをクローン
git clone https://github.com/HTanoda/audio-transcription.git
cd audio-transcription

# 仮想環境を作成
python -m venv venv
venv\Scripts\activate

# 依存パッケージをインストール
pip install -r requirements.txt

# Whisperモデルをダウンロード（初回のみ、約3GB）
python -c "from faster_whisper import WhisperModel; WhisperModel('large-v3', device='cpu', compute_type='int8', download_root='models')"
```

### 実行（開発時）

```bash
python audio_transcription.py
```

### ビルド

詳細は [BUILD_GUIDE.md](BUILD_GUIDE.md) を参照してください。

```bash
# メインアプリのビルド例
pyinstaller --onefile --noconsole --icon "TND_AudioTranscription.ico" --name "TND_audio_transcription" audio_transcription.py
```

---

## 📝 更新履歴

### v1.2.1 (2025-02-03)
- **新機能:** インストーラー/アンインストーラーを追加
- **新機能:** モデル外部化による起動速度の改善
- **改善:** アプリアイコンを追加
- **改善:** 配布サイズの最適化（約2.8GB）

### v1.2.0
- **新機能:** プログレスバーによる処理状況の可視化
- **新機能:** 処理完了後に出力フォルダを自動で開く
- **改善:** モデル読み込みをループ外に移動し処理速度大幅向上
- **改善:** GUIウィンドウのリサイズ対応
- **修正:** UIフリーズ防止のため別スレッド実行

### v1.1.1
- 初期リリース

---

## ❓ よくある質問

### Q: 初回起動時に時間がかかります
**A:** 初回はAIモデルの読み込みに時間がかかります（30秒〜1分程度）。2回目以降は高速に起動します。

### Q: 「モデルフォルダが見つかりません」と表示されます
**A:** アプリを再インストールしてください。

### Q: 文字起こしの精度が低いです
**A:** 以下の点をご確認ください：
- 音声がクリアに録音されているか
- 背景ノイズが大きくないか
- 話者が複数同時に話していないか

### Q: インターネット接続は必要ですか？
**A:** いいえ、不要です。AIモデルはローカルにインストールされているため、オフラインで使用できます。

---

## 📜 ライセンス

このプロジェクトは [MIT License](LICENSE) の下で公開されています。

使用しているサードパーティライブラリのライセンスは [THIRD_PARTY_LICENSES.txt](THIRD_PARTY_LICENSES.txt) を参照してください。

---

## 📞 サポート

問題が発生した場合や、ご要望がある場合は [Issues](https://github.com/HTanoda/audio-transcription/issues) までご連絡ください。
