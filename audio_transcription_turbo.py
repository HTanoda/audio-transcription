"""Turbo版エントリポイント。

標準版 audio_transcription をそのまま利用し、アプリ名（タイトルバー）と
ライセンス表記のみ large-v3-turbo 用に差し替える。ロジックは完全共通。
"""
import audio_transcription as base

base.APP_NAME = "TND_AudioTranscription_turbo"
base.APP_TITLE = f"TND audio_transcription turbo v{base.APP_VERSION}"
base.LICENSE_TEXT = base.LICENSE_TEXT.replace(
    "■ Whisper Model (faster-whisper-large-v3)\n"
    "MIT License  https://huggingface.co/Systran/faster-whisper-large-v3",
    "■ Whisper Model (faster-whisper-large-v3-turbo)\n"
    "MIT License  https://huggingface.co/mobiuslabsgmbh/faster-whisper-large-v3-turbo\n"
    "（OpenAI Whisper large-v3-turbo の CTranslate2 変換版）",
)


if __name__ == "__main__":
    base.main()
