"""Turbo版アンインストーラー。

標準版 uninstall をそのまま利用し、アプリ識別子のみ turbo 用に差し替える。
標準版とは別のレジストリキー・インストール先を対象とする。
"""
import uninstall as base

base.APP_NAME = "TND_AudioTranscription_turbo"
base.APP_DISPLAY_NAME = "TND AI議事録アプリ (Turbo版)"
base.APP_EXE_NAME = "TND_audio_transcription_turbo.exe"


if __name__ == "__main__":
    base.main()
