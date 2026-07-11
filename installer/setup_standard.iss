; -- setup_standard.iss --
; TND AI議事録アプリ (標準版) フルインストーラー
;
; ビルド例:
;   ISCC.exe /DAppVersion=1.5.0 /DSourceDir=..\dist\TND_AudioTranscription_v1.5.0 setup_standard.iss
;
#ifndef AppVersion
  #define AppVersion "1.5.0"
#endif
#ifndef SourceDir
  #define SourceDir "..\dist\TND_AudioTranscription_v" + AppVersion
#endif

#define MyAppName      "TND AI議事録アプリ"
#define MyAppExeName   "TND_audio_transcription.exe"
#define MyAppIcoName   "TND_AudioTranscription01.ico"
#define MyAppPublisher "HIROKI TANODA (TND)"
#define MyAppCopyright "Copyright (c) 2026 HIROKI TANODA (TND)"

[Setup]
AppId={{FC2F6FBE-882B-42B0-8E63-B1B3B049F613}
AppName={#MyAppName}
AppVersion={#AppVersion}
AppPublisher={#MyAppPublisher}
AppCopyright={#MyAppCopyright}
DefaultDirName={localappdata}\TND_AudioTranscription
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\dist
OutputBaseFilename=TND_AudioTranscription-setup-{#AppVersion}
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
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\TND_AudioTranscription"; Flags: deletekey dontcreatekey

[UninstallDelete]
Type: filesandordirs; Name: "{app}\logs"
Type: files; Name: "{app}\settings.json"
Type: files; Name: "{app}\hotwords.json"
