$ErrorActionPreference = "Stop"
$logFile = "$env:APPDATA\dmmgameplayer5\logs\dll.log"
$exeName = "imasscprism.exe"
$shortcutName = "ShaniSong.lnk"

if (-not (Test-Path $logFile)) {
    Write-Host "Error: DMMGamePlayer log file not found."
    pause
    exit
}

$lastLaunchLine = Select-String -Path $logFile -Pattern "Execute of::.*?$exeName" | Select-Object -Last 1 -ExpandProperty Line

if (-not $lastLaunchLine) {
    Write-Host "Error: No ShaniSong launch history found in the DMM log."
    Write-Host "Solution: Open ShaniSong via DMM Player at least once, then try again."
    pause
    exit
}

$regex = 'exe:\s*(?<exe_path>.*?' + [regex]::Escape($exeName) + ').*?/viewer_id=(?<viewer_id>[^\s]+).*?/open_id=(?<open_id>[^\s]+).*?/pf_access_token=(?<pf_token>[^\s]+)'


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

    Write-Host "`nSuccess! The shortcut '$shortcutName' has been created on the Desktop."
}
else {
    Write-Host "Error: Failed to retrieve token. The log structure may be different or the token has not been formed yet."
    Write-Host "Ensure you are logged in and have entered the game's main menu via DMM before running this script."
}

pause

