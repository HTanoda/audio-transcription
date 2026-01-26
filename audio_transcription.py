"""
Audio Transcription Tool
音声ファイルを自動で文字起こしするGUIアプリケーション

Copyright (c) 2025 Hiroki
Released under the MIT License
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from faster_whisper import WhisperModel
from moviepy.editor import AudioFileClip
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime


def resource_path(relative_path):
    """PyInstallerでパッケージ化された際のリソースパスを解決する"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def split_audio_file(file_path, output_folder):
    """
    音声ファイルを分割して文字起こしを行う
    
    Args:
        file_path: 入力音声ファイルのパス
        output_folder: 出力フォルダのパス
    """
    # 入力ファイルの拡張子を取得
    file_extension = os.path.splitext(file_path)[1].lower()
    file_name = os.path.splitext(os.path.basename(file_path))[0]

    # 音声ファイルを読み込む
    audio = AudioFileClip(file_path)

    # 分割する時間間隔（１分）を取得
    split_interval = 1 * 60  # 秒単位

    # 分割した音声ファイルを保存するフォルダを作成
    os.makedirs(output_folder, exist_ok=True)

    list1 = ["", "", ""]
    df = pd.DataFrame([list1])
    df.columns = ['No', '音声ファイル', '変換結果']

    # 音声ファイルを分割する
    for i, start_time in enumerate(range(0, int(audio.duration), split_interval)):
        # 分割開始位置と終了位置を計算
        end_time = min(start_time + split_interval, audio.duration)
        # 音声を分割
        split_audio = audio.subclip(start_time, end_time)
        # 出力ファイル名を作成
        output_file = os.path.join(output_folder, f"{file_name}_{i}{file_extension}")
        # 分割した音声ファイルを保存
        if file_extension == ".wav":
            split_audio.write_audiofile(output_file, codec='pcm_s16le')
        elif file_extension == ".mp3":
            split_audio.write_audiofile(output_file, codec='libmp3lame')
        elif file_extension == ".m4a":
            split_audio.write_audiofile(output_file, codec='aac')
        elif file_extension == ".mp4":
            split_audio.write_audiofile(output_file, codec='aac')
        else:
            raise ValueError("サポートされていない音声形式です。")
        print(f"分割ファイル {output_file} を保存しました。")

        # 音声ファイルを文字変換
        print("start:", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

        model_dir = resource_path("models")
        model = WhisperModel("large-v3", device="cpu", compute_type="int8", download_root=model_dir)
        
        segments, _ = model.transcribe(
            output_file, 
            beam_size=5, 
            language='ja', 
            temperature=0, 
            vad_filter=True, 
            initial_prompt="高島宗一郎です。こんにちは、今日はよろしくお願いします。"
        )

        transcription = ''
        for segment in segments:
            transcription = transcription + str(segment.text) + '\n'
        print(transcription)
        # 結果をdfにセット
        df.loc[i] = [i, output_file, transcription]

    # excelへ書き出し
    print("end:", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
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


def main():
    """メイン関数"""
    root = tk.Tk()
    root.withdraw()

    filetypes = [("音声ファイル", "*.wav;*.mp3;*.m4a;*.mp4")]
    input_file_path = filedialog.askopenfilename(
        title="音声ファイルを選択", 
        filetypes=filetypes, 
        initialdir='./'
    )

    if not input_file_path:
        messagebox.showinfo("キャンセル", "ファイルが選択されませんでした。")
        return

    output_folder_path = filedialog.askdirectory(
        title="出力するフォルダを選択", 
        initialdir='./'
    )

    if not output_folder_path:
        messagebox.showinfo("キャンセル", "出力フォルダが選択されませんでした。")
        return

    try:
        split_audio_file(input_file_path, output_folder_path)
        messagebox.showinfo("完了", "処理が完了しました。")
    except Exception as e:
        messagebox.showerror("エラー", str(e))


if __name__ == "__main__":
    main()
