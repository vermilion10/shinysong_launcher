$ErrorActionPreference = "Stop"
$logFile = "$env:APPDATA\dmmgameplayer5\logs\dll.log"
$exeName = "gakumas.exe"
$shortcutName = "学园偶像大师.lnk"

if (-not (Test-Path $logFile)) {
    Write-Host "错误: DMMGamePlayer 日志文件不存在 ($logFile)"
    pause
    exit
}

$lastLaunchLine = Select-String -Path $logFile -Pattern "Execute of:: gakumas exe" | Select-Object -Last 1 -ExpandProperty Line

if (-not $lastLaunchLine) {
    Write-Host "错误: 在 DMM 日志中找不到《学园偶像大师》的运行记录"
    pause
    exit
}

$regex = 'exe:\s*(?<exe_path>.*?gakumas\.exe).*?/viewer_id=(?<viewer_id>[^\s]+).*?/open_id=(?<open_id>[^\s]+).*?/pf_access_token=(?<pf_token>[^\s]+)'

if ($lastLaunchLine -match $regex) {
    $arguments = "/viewer_id=$($matches.viewer_id) /open_id=$($matches.open_id) /pf_access_token=$($matches.pf_token)"
    $exePath = $matches.exe_path
    $workingDir = Split-Path $exePath -Parent

    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $shortcutPath = Join-Path $desktopPath $shortcutName
    
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $exePath
    $shortcut.Arguments = $arguments
    $shortcut.WorkingDirectory = $workingDir
    $shortcut.IconLocation = $exePath
    $shortcut.Save()

    Write-Host "`n成功! 已在桌面上创建《学园偶像大师》的快捷方式"
}
else {
    Write-Host "错误: 无法从日志中提取完整的启动信息，请先正常启动游戏一次"
}

pause
