@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul

rem ===== 名前を変えるときは、下の2行を編集してください =====
set "FOLDER_NAME=業務量モニター"
set "FILE_NAME=業務量モニター"
rem FILE_NAME はフォルダ直下のファイル名です。拡張子は省略できます。
rem 編集後は UTF-8 で保存してください。名前に % を使う場合は %% と書きます。

set "MONITOR_INSTALLER=%~f0"
rem このバッチ後半の PowerShell 部分を実行します。
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -NonInteractive -Command "$ErrorActionPreference='Stop'; $text=Get-Content -LiteralPath $env:MONITOR_INSTALLER -Raw -Encoding UTF8; $parts=$text -split '(?m)^:MONITOR_SCRIPT\r?$',2; if($parts.Count -ne 2){throw 'Batch file is incomplete.'}; & ([ScriptBlock]::Create($parts[1]))"
set "MONITOR_RESULT=%errorlevel%"
echo(
pause
exit /b %MONITOR_RESULT%

:MONITOR_SCRIPT
$ErrorActionPreference = 'Stop'
try {
    $folderName = $env:FOLDER_NAME
    $fileName = $env:FILE_NAME
    foreach ($name in @($folderName, $fileName)) {
        if ([string]::IsNullOrWhiteSpace($name) -or
            $name.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0 -or
            $name -match '[. ]$') {
            throw '先頭の設定には、パスを含まないフォルダ名とファイル名を入力してください。'
        }
    }

    $source = [IO.Path]::Combine([IO.Path]::GetDirectoryName($env:MONITOR_INSTALLER), $folderName)
    if (-not [IO.Directory]::Exists($source)) {
        throw "バッチと同じ場所にコピー元フォルダが見つかりません。 $source"
    }

    # 配布元にあるファイルから選び、過去にコピーしたファイルの影響を避けます。
    $files = @(Get-ChildItem -LiteralPath $source -File -Force)
    $matchesByName = @($files | Where-Object { $_.Name -eq $fileName })
    if ($matchesByName.Count -eq 0) {
        $matchesByName = @($files | Where-Object { $_.BaseName -eq $fileName })
    }
    if ($matchesByName.Count -eq 0) {
        throw "コピー元に「$fileName」が見つかりません。フォルダ直下のファイル名を確認してください。 $source"
    }
    if ($matchesByName.Count -gt 1) {
        $choices = $matchesByName.Name -join '、'
        throw "候補が複数あります。先頭の FILE_NAME に拡張子まで指定してください。候補は $choices です。"
    }
    $launchFile = $matchesByName[0]

    $shell = New-Object -ComObject Shell.Application
    $downloads = $shell.NameSpace('shell:Downloads').Self.Path
    $desktop = [Environment]::GetFolderPath('DesktopDirectory')
    if (-not $downloads -or -not [IO.Directory]::Exists($desktop)) {
        throw 'Windows のダウンロードフォルダまたはデスクトップの場所を取得できません。'
    }
    $destination = [IO.Path]::GetFullPath([IO.Path]::Combine($downloads, $folderName))
    $target = [IO.Path]::Combine($destination, $launchFile.Name)
    $linkName = $launchFile.BaseName
    if (-not $linkName) { $linkName = $launchFile.Name }
    $linkPath = [IO.Path]::Combine($desktop, ($linkName + '.lnk'))

    if ($source -ne $destination) {
        if ($destination.StartsWith($source + '\', [StringComparison]::OrdinalIgnoreCase) -or
            $source.StartsWith($destination + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw 'コピー元とコピー先が入れ子です。バッチと配布フォルダを別の場所に置いて実行してください。'
        }
        Write-Host "コピー先  $destination"
        $robocopy = [IO.Path]::Combine($env:SystemRoot, 'System32\robocopy.exe')
        # サブフォルダを含めて同名ファイルを上書きします。コピー先の余分な項目は残します。
        & $robocopy $source $destination /E /IS /IT /IM /R:2 /W:1 /NFL /NDL /NP /NJH /NJS
        if ($LASTEXITCODE -ge 8) {
            throw "コピーが完了しませんでした（Robocopy 終了コード $LASTEXITCODE）。"
        }
    } else {
        Write-Host 'すでにダウンロードフォルダに配置されています。'
    }
    if (-not [IO.File]::Exists($target)) {
        throw "コピー先でショートカットの対象を確認できません。 $target"
    }

    $wsh = New-Object -ComObject WScript.Shell
    $link = $wsh.CreateShortcut($linkPath)
    $link.TargetPath = $target
    $link.WorkingDirectory = $destination
    $link.Arguments = ''
    $link.Save()
    if (-not [IO.File]::Exists($linkPath)) {
        throw "ショートカットを保存できませんでした。 $linkPath"
    }
    Write-Host "完了しました。 $linkPath"
    Write-Host "リンク先  $target"
    exit 0
} catch {
    Write-Host ("エラー  " + $_.Exception.Message) -ForegroundColor Red
    exit 1
}
