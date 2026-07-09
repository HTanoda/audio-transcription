import os
import sys
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import winreg
import threading

# アプリケーション情報
APP_NAME = "TND_AudioTranscription"
APP_DISPLAY_NAME = "TND AI議事録アプリ"
APP_VERSION = "1.4.2"
APP_EXE_NAME = "TND_audio_transcription.exe"
UNINSTALLER_NAME = "uninstall.exe"
APP_ICON_NAME = "TND_AudioTranscription01.ico"


def get_default_install_dir():
    """デフォルトのインストール先を取得"""
    local_app_data = os.environ.get('LOCALAPPDATA', os.path.expanduser('~\\AppData\\Local'))
    return os.path.join(local_app_data, APP_NAME)


def get_source_dir():
    """インストール元（setup.exeのあるフォルダ）を取得"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def create_shortcut(target_path, shortcut_path, description="", icon_path=None):
    """ショートカットを作成（PowerShell使用）"""
    # PowerShellスクリプトでショートカットを作成
    ps_script = f'''
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("{shortcut_path}")
$Shortcut.TargetPath = "{target_path}"
$Shortcut.Description = "{description}"
$Shortcut.WorkingDirectory = "{os.path.dirname(target_path)}"
'''
    if icon_path:
        ps_script += f'$Shortcut.IconLocation = "{icon_path}"\n'
    ps_script += '$Shortcut.Save()'
    
    # PowerShellを実行
    import subprocess
    result = subprocess.run(
        ['powershell', '-ExecutionPolicy', 'Bypass', '-Command', ps_script],
        capture_output=True,
        text=True,
        creationflags=subprocess.CREATE_NO_WINDOW
    )
    return result.returncode == 0


def register_uninstaller(install_dir):
    """コントロールパネルの「プログラムと機能」に登録"""
    try:
        uninstall_key = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        key_path = f"{uninstall_key}\\{APP_NAME}"
        
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, APP_DISPLAY_NAME)
            winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, APP_VERSION)
            winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, "TND")
            winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, install_dir)
            winreg.SetValueEx(key, "UninstallString", 0, winreg.REG_SZ, 
                            os.path.join(install_dir, UNINSTALLER_NAME))
            winreg.SetValueEx(key, "NoModify", 0, winreg.REG_DWORD, 1)
            winreg.SetValueEx(key, "NoRepair", 0, winreg.REG_DWORD, 1)
        return True
    except Exception as e:
        print(f"レジストリ登録エラー: {e}")
        return False


class InstallerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(f"{APP_DISPLAY_NAME} セットアップ")
        self.root.geometry("500x450")
        self.root.resizable(True, True)
        
        self.install_dir = tk.StringVar(value=get_default_install_dir())
        self.create_desktop_shortcut = tk.BooleanVar(value=True)
        self.create_startmenu_shortcut = tk.BooleanVar(value=True)
        
        self.setup_ui()
    
    def setup_ui(self):
        # タイトル
        title_frame = ttk.Frame(self.root, padding="20")
        title_frame.pack(fill=tk.X)
        
        ttk.Label(
            title_frame, 
            text=f"{APP_DISPLAY_NAME} v{APP_VERSION}",
            font=("", 14, "bold")
        ).pack()
        ttk.Label(
            title_frame, 
            text="セットアップウィザード",
            font=("", 10)
        ).pack()
        
        # 区切り線
        ttk.Separator(self.root, orient='horizontal').pack(fill=tk.X, padx=20)
        
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # インストール先
        ttk.Label(main_frame, text="インストール先:").pack(anchor=tk.W)
        
        dir_frame = ttk.Frame(main_frame)
        dir_frame.pack(fill=tk.X, pady=(5, 15))
        
        self.dir_entry = ttk.Entry(dir_frame, textvariable=self.install_dir, width=50)
        self.dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(dir_frame, text="参照...", command=self.browse_dir).pack(side=tk.RIGHT, padx=(5, 0))
        
        # オプション
        ttk.Label(main_frame, text="オプション:").pack(anchor=tk.W, pady=(10, 5))
        
        ttk.Checkbutton(
            main_frame, 
            text="デスクトップにショートカットを作成",
            variable=self.create_desktop_shortcut
        ).pack(anchor=tk.W)
        
        ttk.Checkbutton(
            main_frame, 
            text="スタートメニューに登録",
            variable=self.create_startmenu_shortcut
        ).pack(anchor=tk.W)
        
        # プログレスバー
        self.progress_frame = ttk.LabelFrame(main_frame, text="進捗", padding="10")
        self.progress_frame.pack(fill=tk.X, pady=(20, 0))
        
        self.status_label = ttk.Label(self.progress_frame, text="インストール準備完了")
        self.status_label.pack(fill=tk.X)
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate', length=400)
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))
        
        # ボタン
        button_frame = ttk.Frame(self.root, padding="20")
        button_frame.pack(fill=tk.X)
        
        self.install_btn = ttk.Button(
            button_frame, 
            text="インストール", 
            command=self.start_install
        )
        self.install_btn.pack(side=tk.RIGHT)
        
        ttk.Button(
            button_frame, 
            text="キャンセル", 
            command=self.root.quit
        ).pack(side=tk.RIGHT, padx=(0, 10))
    
    def browse_dir(self):
        directory = filedialog.askdirectory(
            title="インストール先を選択",
            initialdir=self.install_dir.get()
        )
        if directory:
            self.install_dir.set(os.path.join(directory, APP_NAME))
    
    def update_progress(self, value, status):
        self.progress_bar['value'] = value
        self.status_label.config(text=status)
        self.root.update_idletasks()
    
    def start_install(self):
        self.install_btn.config(state=tk.DISABLED)
        self.dir_entry.config(state=tk.DISABLED)
        
        thread = threading.Thread(target=self.install)
        thread.daemon = True
        thread.start()
    
    def install(self):
        try:
            source_dir = get_source_dir()
            install_dir = self.install_dir.get()
            
            # インストール先フォルダを作成
            self.root.after(0, lambda: self.update_progress(10, "インストール先を準備中..."))
            os.makedirs(install_dir, exist_ok=True)
            
            # メインEXEをコピー
            self.root.after(0, lambda: self.update_progress(20, "アプリケーションをコピー中..."))
            src_exe = os.path.join(source_dir, APP_EXE_NAME)
            if os.path.exists(src_exe):
                shutil.copy2(src_exe, install_dir)
            else:
                raise FileNotFoundError(f"{APP_EXE_NAME} が見つかりません")
            
            # アンインストーラーをコピー
            self.root.after(0, lambda: self.update_progress(25, "アンインストーラーをコピー中..."))
            src_uninstaller = os.path.join(source_dir, UNINSTALLER_NAME)
            if os.path.exists(src_uninstaller):
                shutil.copy2(src_uninstaller, install_dir)
            
            # README.txtをコピー（存在する場合）
            src_readme = os.path.join(source_dir, "README.txt")
            if os.path.exists(src_readme):
                shutil.copy2(src_readme, install_dir)
            
            # アイコンファイルをコピー（存在する場合）
            src_icon = os.path.join(source_dir, APP_ICON_NAME)
            if os.path.exists(src_icon):
                shutil.copy2(src_icon, install_dir)
            
            # modelsフォルダをコピー（時間がかかる）
            self.root.after(0, lambda: self.update_progress(30, "モデルファイルをコピー中...\n（約3GB、数分かかります）"))
            src_models = os.path.join(source_dir, "models")
            dst_models = os.path.join(install_dir, "models")
            
            if os.path.exists(src_models):
                if os.path.exists(dst_models):
                    shutil.rmtree(dst_models)
                shutil.copytree(src_models, dst_models)
            else:
                raise FileNotFoundError("models フォルダが見つかりません")
            
            self.root.after(0, lambda: self.update_progress(80, "ショートカットを作成中..."))
            
            # アイコンのパス
            icon_path = os.path.join(install_dir, APP_ICON_NAME)
            if not os.path.exists(icon_path):
                icon_path = None  # アイコンがなければEXEのアイコンを使用
            
            # デスクトップショートカット
            if self.create_desktop_shortcut.get():
                desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
                shortcut_path = os.path.join(desktop, f"{APP_DISPLAY_NAME}.lnk")
                target_path = os.path.join(install_dir, APP_EXE_NAME)
                create_shortcut(target_path, shortcut_path, APP_DISPLAY_NAME, icon_path)
            
            # スタートメニュー
            if self.create_startmenu_shortcut.get():
                start_menu = os.path.join(
                    os.environ['APPDATA'], 
                    'Microsoft\\Windows\\Start Menu\\Programs'
                )
                shortcut_path = os.path.join(start_menu, f"{APP_DISPLAY_NAME}.lnk")
                target_path = os.path.join(install_dir, APP_EXE_NAME)
                create_shortcut(target_path, shortcut_path, APP_DISPLAY_NAME, icon_path)
            
            # レジストリに登録
            self.root.after(0, lambda: self.update_progress(90, "システムに登録中..."))
            register_uninstaller(install_dir)
            
            self.root.after(0, lambda: self.update_progress(100, "インストール完了！"))
            
            self.root.after(0, lambda: self.on_install_complete(install_dir))
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("エラー", f"インストールに失敗しました:\n{e}"))
            self.root.after(0, lambda: self.install_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.dir_entry.config(state=tk.NORMAL))
    
    def on_install_complete(self, install_dir):
        result = messagebox.askquestion(
            "インストール完了",
            f"インストールが完了しました。\n\n"
            f"インストール先: {install_dir}\n\n"
            f"今すぐアプリケーションを起動しますか？"
        )
        if result == 'yes':
            exe_path = os.path.join(install_dir, APP_EXE_NAME)
            os.startfile(exe_path)
        self.root.quit()
    
    def run(self):
        self.root.mainloop()


def main():
    # 必要なファイルの存在確認
    source_dir = get_source_dir()
    required_files = [APP_EXE_NAME, "models"]
    
    missing = []
    for f in required_files:
        path = os.path.join(source_dir, f)
        if not os.path.exists(path):
            missing.append(f)
    
    if missing:
        messagebox.showerror(
            "エラー",
            f"必要なファイルが見つかりません:\n\n" + "\n".join(missing) + 
            f"\n\nsetup.exe と同じフォルダに配置してください。"
        )
        return
    
    app = InstallerApp()
    app.run()


if __name__ == "__main__":
    main()
