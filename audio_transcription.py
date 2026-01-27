import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from faster_whisper import WhisperModel
from moviepy.editor import AudioFileClip
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime
import threading

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


class AudioTranscriptionApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("TND audio_transcription v1.2.0")
        self.root.geometry("500x350")
        self.root.resizable(True, True)
        
        self.input_file_path = None
        self.output_folder_path = None
        self.is_processing = False
        
        self.setup_ui()
    
    def setup_ui(self):
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 入力ファイル選択
        file_frame = ttk.LabelFrame(main_frame, text="入力ファイル", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.file_label = ttk.Label(file_frame, text="ファイルが選択されていません")
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.file_btn = ttk.Button(file_frame, text="選択", command=self.select_input_file)
        self.file_btn.pack(side=tk.RIGHT)
        
        # 出力フォルダ選択
        folder_frame = ttk.LabelFrame(main_frame, text="出力フォルダ", padding="10")
        folder_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.folder_label = ttk.Label(folder_frame, text="フォルダが選択されていません")
        self.folder_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.folder_btn = ttk.Button(folder_frame, text="選択", command=self.select_output_folder)
        self.folder_btn.pack(side=tk.RIGHT)
        
        # プログレスバーフレーム
        progress_frame = ttk.LabelFrame(main_frame, text="処理状況", padding="10")
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.status_label = ttk.Label(progress_frame, text="待機中...")
        self.status_label.pack(fill=tk.X, pady=(0, 5))
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', length=400)
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))
        
        self.progress_detail = ttk.Label(progress_frame, text="")
        self.progress_detail.pack(fill=tk.X)
        
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
            segments, _ = model.transcribe(output_file, 
                                           beam_size=5, 
                                           language='ja', 
                                           temperature=0, 
                                           vad_filter=True, 
                                           initial_prompt="高島宗一郎です。こんにちは、今日はよろしくお願いします。")

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
