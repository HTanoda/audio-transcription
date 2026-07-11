; -- setup_turbo.iss --
; TND AI議事録アプリ (Turbo版) フルインストーラー
;
; ビルド例:
;   ISCC.exe /DAppVersion=1.5.0 /DSourceDir=..\dist\TND_AudioTranscription_turbo_v1.5.0 setup_turbo.iss
;
#ifndef AppVersion
  #define AppVersion "1.5.0"
#endif
#ifndef SourceDir
  #define SourceDir "..\dist\TND_AudioTranscription_turbo_v" + AppVersion
#endif

#define MyAppName      "TND AI議事録アプリ (Turbo版)"
#define MyAppExeName   "TND_audio_transcription_turbo.exe"
#define MyAppIcoName   "TND_AudioTranscription01.ico"
#define MyAppPublisher "HIROKI TANODA (TND)"
#define MyAppCopyright "Copyright (c) 2026 HIROKI TANODA (TND)"

[Setup]
AppId={{C2620700-450A-4C11-84E3-072A9C2A834E}
AppName={#MyAppName}
AppVersion={#AppVersion}
AppPublisher={#MyAppPublisher}
AppCopyright={#MyAppCopyright}
DefaultDirName={localappdata}\TND_AudioTranscription_turbo
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\dist
OutputBaseFilename=TND_AudioTranscription_turbo-setup-{#AppVersion}
Compression=lzma2
SolidCompression=no
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
WizardStyle=modern
SetupIconFile={#SourceDir}\{#MyAppIcoName}
UninstallDisplayIcon={app}\{#MyAppIcoName}
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "デスクトップにショートカットを作成"; GroupDescription: "追加アイコン:"

[Files]
Source: "{#SourceDir}\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\README.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\{#MyAppIcoName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\models\*"; DestDir: "{app}\models"; Flags: recursesubdirs createallsubdirs ignoreversion nocompression

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppIcoName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppIcoName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "アプリを起動"; Flags: nowait postinstall skipifsilent

[InstallDelete]
; 旧Python製インストーラの残骸を掃除
Type: files; Name: "{app}\uninstall.exe"

[Registry]
; 旧Python製インストーラのアンインストール登録を削除（アプリ一覧の二重表示防止）
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\TND_AudioTranscription_turbo"; Flags: deletekey dontcreatekey

[UninstallDelete]
Type: filesandordirs; Name: "{app}\logs"
Type: files; Name: "{app}\settings.json"
Type: files; Name: "{app}\hotwords.json"
