"""Turbo版インストーラー。

標準版 setup をそのまま利用し、アプリ識別子（インストール先・レジストリキー・
表示名・メインEXE名）のみ turbo 用に差し替える。標準版と併存インストール可能。
"""
import setup as base

base.APP_NAME = "TND_AudioTranscription_turbo"
base.APP_DISPLAY_NAME = "TND AI議事録アプリ (Turbo版)"
base.APP_EXE_NAME = "TND_audio_transcription_turbo.exe"


if __name__ == "__main__":
    base.main()
