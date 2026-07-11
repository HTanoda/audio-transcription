; -- update_standard.iss --
; TND AI議事録アプリ (標準版) 差分更新インストーラー
; 本体EXEとREADME.txtのみを上書きする（modelsは含まない）。
; フルインストーラー (setup_standard.iss) と同じ AppId を使うが、
; このインストーラー自体はアンインストール登録を行わない
; （フル版のアンインストール登録を壊さないため）。
;
; ビルド例:
;   ISCC.exe /DAppVersion=1.5.0 /DSourceDir=..\dist\TND_AudioTranscription_v1.5.0 update_standard.iss
;
#ifndef AppVersion
  #define AppVersion "1.5.0"
#endif
#ifndef SourceDir
  #define SourceDir "..\dist\TND_AudioTranscription_v" + AppVersion
#endif

#define MyAppName      "TND AI議事録アプリ"
#define MyAppExeName   "TND_audio_transcription.exe"
#define MyAppPublisher "HIROKI TANODA (TND)"
#define MyAppCopyright "Copyright (c) 2026 HIROKI TANODA (TND)"

[Setup]
AppId={{FC2F6FBE-882B-42B0-8E63-B1B3B049F613}
AppName={#MyAppName} (更新)
AppVersion={#AppVersion}
AppPublisher={#MyAppPublisher}
AppCopyright={#MyAppCopyright}
DefaultDirName={localappdata}\TND_AudioTranscription
DisableDirPage=no
DisableProgramGroupPage=yes
DisableReadyPage=yes
OutputDir=..\dist
OutputBaseFilename=TND_AudioTranscription-update-{#AppVersion}
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

[Files]
Source: "{#SourceDir}\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourceDir}\README.txt"; DestDir: "{app}"; Flags: ignoreversion

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
