<#
.SYNOPSIS
    TND AI議事録アプリ（標準版 / Turbo版）の Inno Setup インストーラーを一括ビルドする。
.DESCRIPTION
    setup_standard.iss / setup_turbo.iss / update_standard.iss / update_turbo.iss の
    4本を ISCC.exe でビルドし、dist フォルダに出力する。
    Windows PowerShell 5.1 互換（&& / ?? は使用しない）。
.PARAMETER Version
    ビルド対象のアプリバージョン（既定: 1.6.0）
.PARAMETER StandardDir
    標準版の配布用ソースフォルダ（既定: ..\dist\TND_AudioTranscription_v<Version>）
.PARAMETER TurboDir
    Turbo版の配布用ソースフォルダ（既定: ..\dist\TND_AudioTranscription_turbo_v<Version>）
.EXAMPLE
    .\build_installers.ps1
.EXAMPLE
    .\build_installers.ps1 -Version 1.6.0 -StandardDir "..\dist\TND_AudioTranscription_v1.6.0" -TurboDir "..\dist\TND_AudioTranscription_turbo_v1.6.0"
#>
param(
    [string]$Version = "1.6.0",
    [string]$StandardDir = "",
    [string]$TurboDir = ""
)

$ErrorActionPreference = "Stop"

$ScriptDir = $PSScriptRoot
$ISCC = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"

if ([string]::IsNullOrWhiteSpace($StandardDir)) {
    $StandardDir = Join-Path $ScriptDir "..\dist\TND_AudioTranscription_v$Version"
}
if ([string]::IsNullOrWhiteSpace($TurboDir)) {
    $TurboDir = Join-Path $ScriptDir "..\dist\TND_AudioTranscription_turbo_v$Version"
}

if (-not (Test-Path -LiteralPath $ISCC)) {
    Write-Error "ISCC.exe が見つかりません: $ISCC"
    exit 1
}

if (-not (Test-Path -LiteralPath $StandardDir)) {
    Write-Error "標準版ソースフォルダが見つかりません: $StandardDir"
    exit 1
}

if (-not (Test-Path -LiteralPath $TurboDir)) {
    Write-Error "Turbo版ソースフォルダが見つかりません: $TurboDir"
    exit 1
}

Write-Host "=== ビルド対象 ===" -ForegroundColor Cyan
Write-Host "Version      : $Version"
Write-Host "StandardDir  : $StandardDir"
Write-Host "TurboDir     : $TurboDir"
Write-Host ""

$targets = @(
    @{ Name = "標準版フルインストーラー";     Script = "setup_standard.iss";    SourceDir = $StandardDir },
    @{ Name = "Turbo版フルインストーラー";     Script = "setup_turbo.iss";       SourceDir = $TurboDir },
    @{ Name = "標準版差分更新インストーラー";  Script = "update_standard.iss";   SourceDir = $StandardDir },
    @{ Name = "Turbo版差分更新インストーラー"; Script = "update_turbo.iss";      SourceDir = $TurboDir }
)

foreach ($target in $targets) {
    $issPath = Join-Path $ScriptDir $target.Script
    if (-not (Test-Path -LiteralPath $issPath)) {
        Write-Error "iss ファイルが見つかりません: $issPath"
        exit 1
    }

    Write-Host "--- $($target.Name) ($($target.Script)) ---" -ForegroundColor Yellow
    & $ISCC "/DAppVersion=$Version" "/DSourceDir=$($target.SourceDir)" $issPath
    if ($LASTEXITCODE -ne 0) {
        Write-Error "ビルド失敗: $($target.Script) (ExitCode=$LASTEXITCODE)"
        exit $LASTEXITCODE
    }
    Write-Host ""
}

Write-Host "=== 全4本のビルドが完了しました ===" -ForegroundColor Green
Write-Host ("出力先: {0}" -f (Join-Path $ScriptDir "..\dist"))
