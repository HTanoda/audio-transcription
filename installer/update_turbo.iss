; -- update_turbo.iss --
; TND AI議事録アプリ (Turbo版) 差分更新インストーラー
; 本体一式（onedir の app\* = EXE + _internal\）と README.txt を上書きする
; （models は含まない）。models_diarization は v1.6.0 の新規追加物のため
; 差分更新でも同梱する。
; フルインストーラー (setup_turbo.iss) と同じ AppId を使うが、
; このインストーラー自体はアンインストール登録を行わない
; （フル版のアンインストール登録を壊さないため）。
;
; ビルド例:
;   ISCC.exe /DAppVersion=1.7.3 /DSourceDir=..\dist\TND_AudioTranscription_turbo_v1.7.3 update_turbo.iss
;
#ifndef AppVersion
  #define AppVersion "1.7.3"
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
AppName={#MyAppName} (更新)
AppVersion={#AppVersion}
AppPublisher={#MyAppPublisher}
AppCopyright={#MyAppCopyright}
DefaultDirName={localappdata}\TND_AudioTranscription_turbo
DisableDirPage=no
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
DisableReadyPage=yes
OutputDir=..\dist
OutputBaseFilename=TND_AudioTranscription_turbo-update-{#AppVersion}
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
WizardStyle=modern
Uninstallable=no
CreateUninstallRegKey=no
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
; フル版と同じ。差分更新だけを入れた PC (フル版未導入) でもショートカットを作れるようにする (v1.7.3)
Name: "desktopicon"; Description: "デスクトップにショートカットを作成"; GroupDescription: "追加アイコン:"

[Files]
; PyInstaller onedir 出力一式（本体EXE + _internal\）
Source: "{#SourceDir}\app\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion
Source: "{#SourceDir}\README.txt"; DestDir: "{app}"; Flags: ignoreversion
; ショートカット用アイコン（フル版と同じ）
Source: "{#SourceDir}\{#MyAppIcoName}"; DestDir: "{app}"; Flags: ignoreversion
; 同梱ライブラリのライセンス全文（BSD/MIT の添付条件）。相対パスはこの .iss のフォルダ基準
Source: "..\THIRD_PARTY_LICENSES.txt"; DestDir: "{app}"; Flags: ignoreversion
; 話者分離モデル（v1.6.0 新規追加。models本体3GB級は差分更新の対象外という
; 既存の方針を踏襲するが、models_diarizationは新規追加物のため含める）
Source: "{#SourceDir}\models_diarization\*"; DestDir: "{app}\models_diarization"; Flags: recursesubdirs createallsubdirs ignoreversion nocompression

[Icons]
; フル版と同じ名前・参照先で作る（既存のショートカットは同じ内容で上書きされるだけ）
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppIcoName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppIcoName}"; Tasks: desktopicon

[Code]
function NextButtonClick(CurPageID: Integer): Boolean;
var
  ExePath: string;
begin
  Result := True;
  if CurPageID = wpSelectDir then
  begin
    ExePath := AddBackslash(WizardForm.DirEdit.Text) + '{#MyAppExeName}';
    if not FileExists(ExePath) then
    begin
      if MsgBox(
        '選択されたフォルダに ' + '{#MyAppExeName}' + ' が見つかりません。' + #13#10 +
        '既存インストールを更新する場合は、正しいインストール先を選択してください。' + #13#10#13#10 +
        '新規インストール先として、このまま続行しますか？',
        mbConfirmation, MB_YESNO) = IDNO then
        Result := False;
    end;
  end;
end;
