import os
import sys
import shutil
import tkinter as tk
from tkinter import ttk, messagebox
import winreg
import threading
import subprocess
import time

# アプリケーション情報
APP_NAME = "TND_AudioTranscription"
APP_DISPLAY_NAME = "TND AI議事録アプリ"
APP_EXE_NAME = "TND_audio_transcription.exe"


def is_app_running():
    """メインアプリが実行中かチェック"""
    try:
        result = subprocess.run(
            ['tasklist', '/FI', f'IMAGENAME eq {APP_EXE_NAME}', '/NH'],
            capture_output=True,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        return APP_EXE_NAME.lower() in result.stdout.lower()
    except:
        return False


def kill_app_process():
    """メインアプリのプロセスを終了"""
    try:
        subprocess.run(
            ['taskkill', '/F', '/IM', APP_EXE_NAME],
            capture_output=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        return True
    except:
        return False


def get_install_dir_from_registry():
    """レジストリからインストール先を取得"""
    try:
        uninstall_key = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        key_path = f"{uninstall_key}\\{APP_NAME}"
        
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            install_dir, _ = winreg.QueryValueEx(key, "InstallLocation")
            return install_dir
    except FileNotFoundError:
        return None
    except Exception as e:
        print(f"レジストリ読み取りエラー: {e}")
        return None


def remove_shortcut(shortcut_path):
    """ショートカットを削除"""
    if os.path.exists(shortcut_path):
        try:
            os.remove(shortcut_path)
            return True
        except:
            return False
    return True


def unregister_from_registry():
    """レジストリからアンインストール情報を削除"""
    try:
        uninstall_key = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        key_path = f"{uninstall_key}\\{APP_NAME}"
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
        return True
    except FileNotFoundError:
        return True  # キーが存在しない場合は成功とみなす
    except Exception as e:
        print(f"レジストリ削除エラー: {e}")
        return False


class UninstallerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(f"{APP_DISPLAY_NAME} アンインストール")
        self.root.geometry("450x320")
        self.root.resizable(True, True)
        
        # レジストリからインストール先を取得
        self.install_dir = get_install_dir_from_registry()
        
        if not self.install_dir:
            messagebox.showerror(
                "エラー",
                f"{APP_DISPLAY_NAME} のインストール情報が見つかりません。\n\n"
                f"アプリケーションがインストールされていないか、\n"
                f"既にアンインストール済みの可能性があります。"
            )
            self.root.destroy()
            return
        
        if not os.path.exists(self.install_dir):
            messagebox.showerror(
                "エラー",
                f"インストールフォルダが見つかりません:\n{self.install_dir}\n\n"
                f"既に削除されている可能性があります。\n"
                f"レジストリ情報のみ削除します。"
            )
            # レジストリだけ削除
            unregister_from_registry()
            self.install_dir = None  # GUIを表示しないようにする
            self.root.destroy()
            return
        
        self.setup_ui()
    
    def setup_ui(self):
        # タイトル
        title_frame = ttk.Frame(self.root, padding="20")
        title_frame.pack(fill=tk.X)
        
        ttk.Label(
            title_frame, 
            text=f"{APP_DISPLAY_NAME}",
            font=("", 14, "bold")
        ).pack()
        ttk.Label(
            title_frame, 
            text="アンインストール",
            font=("", 10)
        ).pack()
        
        # 区切り線
        ttk.Separator(self.root, orient='horizontal').pack(fill=tk.X, padx=20)
        
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(
            main_frame, 
            text=f"以下のフォルダからアプリケーションを削除します:\n\n{self.install_dir}",
            wraplength=400
        ).pack(anchor=tk.W)
        
        # プログレスバー
        self.progress_frame = ttk.Frame(main_frame)
        self.progress_frame.pack(fill=tk.X, pady=(20, 0))
        
        self.status_label = ttk.Label(self.progress_frame, text="")
        self.status_label.pack(fill=tk.X)
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate', length=400)
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))
        
        # ボタン
        button_frame = ttk.Frame(self.root, padding="20")
        button_frame.pack(fill=tk.X)
        
        self.uninstall_btn = ttk.Button(
            button_frame, 
            text="アンインストール", 
            command=self.confirm_uninstall
        )
        self.uninstall_btn.pack(side=tk.RIGHT)
        
        ttk.Button(
            button_frame, 
            text="キャンセル", 
            command=self.root.quit
        ).pack(side=tk.RIGHT, padx=(0, 10))
    
    def update_progress(self, value, status):
        self.progress_bar['value'] = value
        self.status_label.config(text=status)
        self.root.update_idletasks()
    
    def confirm_uninstall(self):
        # アプリが実行中かチェック
        if is_app_running():
            result = messagebox.askquestion(
                "確認",
                f"{APP_DISPLAY_NAME} が実行中です。\n\n"
                f"アンインストールを続行するには、アプリを終了する必要があります。\n"
                f"アプリを終了してアンインストールを続行しますか？"
            )
            if result == 'yes':
                kill_app_process()
                # 少し待機してプロセス終了を確認
                time.sleep(1)
                if is_app_running():
                    messagebox.showerror(
                        "エラー",
                        f"アプリを終了できませんでした。\n\n"
                        f"手動で {APP_DISPLAY_NAME} を終了してから、\n"
                        f"再度アンインストールを実行してください。"
                    )
                    return
            else:
                return
        
        result = messagebox.askquestion(
            "確認",
            f"{APP_DISPLAY_NAME} をアンインストールしますか？\n\n"
            f"モデルファイルを含むすべてのファイルが削除されます。"
        )
        if result == 'yes':
            self.uninstall_btn.config(state=tk.DISABLED)
            thread = threading.Thread(target=self.uninstall)
            thread.daemon = True
            thread.start()
    
    def uninstall(self):
        try:
            # 念のため再度アプリ終了を試みる
            if is_app_running():
                kill_app_process()
                time.sleep(1)
            
            # ショートカットを削除
            self.root.after(0, lambda: self.update_progress(20, "ショートカットを削除中..."))
            
            # デスクトップ
            desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
            remove_shortcut(os.path.join(desktop, f"{APP_DISPLAY_NAME}.lnk"))
            
            # スタートメニュー
            start_menu = os.path.join(
                os.environ['APPDATA'], 
                'Microsoft\\Windows\\Start Menu\\Programs'
            )
            remove_shortcut(os.path.join(start_menu, f"{APP_DISPLAY_NAME}.lnk"))
            
            # レジストリから削除
            self.root.after(0, lambda: self.update_progress(40, "システムから登録解除中..."))
            unregister_from_registry()
            
            # アプリケーションフォルダを削除
            self.root.after(0, lambda: self.update_progress(60, "ファイルを削除中...\n（モデルファイルの削除に時間がかかります）"))
            
            if os.path.exists(self.install_dir):
                shutil.rmtree(self.install_dir)
            
            self.root.after(0, lambda: self.update_progress(100, "アンインストール完了"))
            
            # 完了メッセージ
            self.root.after(0, lambda: self.on_uninstall_complete())
            
        except PermissionError as e:
            self.root.after(0, lambda: messagebox.showerror(
                "エラー", 
                f"ファイルの削除に失敗しました。\n\n"
                f"{APP_DISPLAY_NAME} が実行中の可能性があります。\n"
                f"アプリを終了してから再度お試しください。"
            ))
            self.root.after(0, lambda: self.uninstall_btn.config(state=tk.NORMAL))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("エラー", f"アンインストールに失敗しました:\n{e}"))
            self.root.after(0, lambda: self.uninstall_btn.config(state=tk.NORMAL))
    
    def on_uninstall_complete(self):
        messagebox.showinfo(
            "完了",
            f"{APP_DISPLAY_NAME} のアンインストールが完了しました。"
        )
        self.root.quit()
    
    def run(self):
        if self.install_dir:  # インストール先が見つかった場合のみ実行
            self.root.mainloop()


def main():
    app = UninstallerApp()
    app.run()


if __name__ == "__main__":
    main()
