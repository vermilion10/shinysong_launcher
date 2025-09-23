$currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $lang = (Get-UICulture).Name
    switch -wildcard ($lang) {
        "zh-*" {
            Write-Host -ForegroundColor Red "错误：权限不足。"
            Write-Host -ForegroundColor Yellow "请右键单击此脚本，然后选择 '以管理员身份运行'。"
        }
        "ja-*" {
            Write-Host -ForegroundColor Red "エラー：管理者権限がありません。"
            Write-Host -ForegroundColor Yellow "このスクリプトを右クリックし、「管理者として実行」を選択してください。"
        }
        default { 
            Write-Host -ForegroundColor Red "Error: Insufficient permissions."
            Write-Host -ForegroundColor Yellow "Please right-click this script and select 'Run as administrator'."
        }
    }
    Write-Host "`nPress any key to exit..."
    pause | Out-Null
    exit
}

$lang = (Get-UICulture).Name
$gameName = "Gakumas"
$shortcutName = "Gakumas.lnk"
$taskName = "GakumasLauncherTask"

switch -wildcard ($lang) {
    "zh-*" {
        $gameName = "学园偶像大师"
        $shortcutName = "学园偶像大师.lnk"
        $logFile = "$env:APPDATA\dmmgameplayer5\logs\dll.log"
        $messages = @{
            LogFileNotFound = "错误: DMMGamePlayer 日志文件不存在 ($logFile)"
            LogLineNotFound = "错误: 在 DMM 日志中找不到《学园偶像大师》的运行记录。请先通过 DMM 客户端正常启动一次游戏。"
            ExtractionError = "错误: 无法从日志中提取完整的启动信息。日志格式可能已更改。"
            SuccessExtract = "成功提取游戏信息:"
            Path = "  - 路径: "
            Directory = "  - 目录: "
            CreatingTask = "`n正在创建/更新计划任务 `"$taskName`" ..."
            TaskSuccess = "计划任务设置成功！"
            CreatingShortcut = "`n正在创建桌面快捷方式..."
            ShortcutSuccess = "`n成功! 已在桌面上创建《学园偶像大师》的快捷方式。"
            PressKeyToExit = "`n按任意键退出..."
            UnexpectedError = "`n在执行过程中发生意外错误："
        }
    }
    "ja-*" {
        $gameName = "学マス"
        $shortcutName = "学園アイドルマスター.lnk"
        $logFile = "$env:APPDATA\dmmgameplayer5\logs\dll.log"
        $messages = @{
            LogFileNotFound = "エラー：DMMGamePlayerのログファイルが見つかりません ($logFile)"
            LogLineNotFound = "エラー：DMMログに「学園アイドルマスター」の実行記録が見つかりません。DMMクライアントから一度ゲームを正常に起動してください。"
            ExtractionError = "エラー：ログから完全な起動情報を抽出できません。ログの形式が変更された可能性があります。"
            SuccessExtract = "ゲーム情報の抽出に成功しました:"
            Path = "  - パス: "
            Directory = "  - ディレクトリ: "
            CreatingTask = "`nスケジュールされたタスク `"$taskName`" を作成/更新しています..."
            TaskSuccess = "スケジュールされたタスクの設定に成功しました！"
            CreatingShortcut = "`nデスクトップショートカットを作成しています..."
            ShortcutSuccess = "`n完了！デスクトップに「学園アイドルマスター」のショートカットが作成されました。"
            PressKeyToExit = "`nキーを押して終了..."
            UnexpectedError = "`n実行中に予期せぬエラーが発生しました："
        }
    }
    default {
        $gameName = "GAKUM@S"
        $shortcutName = "GAKUM@S.lnk"
        $logFile = "$env:APPDATA\dmmgameplayer5\logs\dll.log"
        $messages = @{
            LogFileNotFound = "Error: DMMGamePlayer log file does not exist ($logFile)"
            LogLineNotFound = "Error: Could not find any run records for 'GAKUM@S' in the DMM log. Please launch the game normally once through the DMM client."
            ExtractionError = "Error: Failed to extract complete launch information from the log. The log format may have changed."
            SuccessExtract = "Successfully extracted game information:"
            Path = "  - Path: "
            Directory = "  - Directory: "
            CreatingTask = "`nCreating/updating scheduled task `"$taskName`" ..."
            TaskSuccess = "Scheduled task set up successfully!"
            CreatingShortcut = "`nCreating desktop shortcut..."
            ShortcutSuccess = "`nSuccess! A shortcut for 'GAKUM@S' has been created on the desktop."
            PressKeyToExit = "`nPress any key to exit..."
            UnexpectedError = "`nAn unexpected error occurred during execution:"
        }
    }
}

$ErrorActionPreference = "Stop"

try {
    if (-not (Test-Path $logFile)) {
        throw $messages.LogFileNotFound
    }

    $lastLaunchLine = Select-String -Path $logFile -Pattern "Execute of:: gakumas exe" | Select-Object -Last 1 -ExpandProperty Line
    if (-not $lastLaunchLine) {
        throw $messages.LogLineNotFound
    }

    $regex = 'exe:\s*(?<exe_path>.*?gakumas\.exe).*?/viewer_id=(?<viewer_id>[^\s]+).*?/open_id=(?<open_id>[^\s]+).*?/pf_access_token=(?<pf_token>[^\s]+)'
    if ($lastLaunchLine -notmatch $regex) {
        throw $messages.ExtractionError
    }

    $exePath = $matches.exe_path
    $arguments = "/viewer_id=$($matches.viewer_id) /open_id=$($matches.open_id) /pf_access_token=$($matches.pf_token)"
    $workingDir = Split-Path $exePath -Parent

    Write-Host $messages.SuccessExtract
    Write-Host "$($messages.Path)$exePath"
    Write-Host "$($messages.Directory)$workingDir"

    Write-Host $messages.CreatingTask
    
    $taskAction = New-ScheduledTaskAction -Execute $exePath -Argument $arguments -WorkingDirectory $workingDir
    $taskPrincipal = New-ScheduledTaskPrincipal -UserId (Get-CimInstance -ClassName Win32_ComputerSystem).UserName -LogonType Interactive -RunLevel Highest
    $taskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Seconds 0)
    
    Register-ScheduledTask -TaskName $taskName -Action $taskAction -Principal $taskPrincipal -Settings $taskSettings -Force -ErrorAction Stop
    
    Write-Host $messages.TaskSuccess

    Write-Host $messages.CreatingShortcut
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $shortcutPath = Join-Path $desktopPath $shortcutName

    $exeDir = Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)

    $vbsPath = Join-Path $exeDir "Launch_$taskName.vbs"

$vbsContent = @"
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "schtasks /run /tn ""$taskName""", 0, False
"@

Set-Content -Path $vbsPath -Value $vbsContent -Encoding ASCII

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $vbsPath
    $shortcut.WorkingDirectory = $workingDir
    $shortcut.IconLocation = $exePath
    $shortcut.Save()

    Write-Host $messages.ShortcutSuccess
}
catch {

    Write-Host -ForegroundColor Red $messages.UnexpectedError
    Write-Host -ForegroundColor Red $_.Exception.Message
}
finally {
    Write-Host $messages.PressKeyToExit
    pause | Out-Null
}
