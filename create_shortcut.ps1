# create_shortcut.ps1  -  Create a desktop shortcut for Automatic Writing Assistant
# Run:  powershell -ExecutionPolicy Bypass -File create_shortcut.ps1

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

# Detect whether we have a built .exe or fall back to python script
$exe = Join-Path $root "dist\AutomaticWritingAssistant.exe"
if (Test-Path $exe) {
    $targetPath = $exe
    $arguments  = ""
    $workDir    = Join-Path $root "dist"
} else {
    $targetPath = (Get-Command python).Source
    $arguments  = "`"$(Join-Path $root 'app.py')`""
    $workDir    = $root
}

$iconFile = Join-Path $root "icon.ico"
$desktop  = [Environment]::GetFolderPath("Desktop")
$lnkPath  = Join-Path $desktop "Automatic Writing Assistant.lnk"

$shell    = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($lnkPath)
$shortcut.TargetPath       = $targetPath
$shortcut.Arguments        = $arguments
$shortcut.WorkingDirectory = $workDir
$shortcut.Description      = "Automatic Writing Assistant v2.0"
$shortcut.WindowStyle      = 1   # Normal window

if (Test-Path $iconFile) {
    $shortcut.IconLocation = "$iconFile,0"
}

$shortcut.Save()

Write-Host ""
Write-Host "Desktop shortcut created: $lnkPath"
Write-Host "Target: $targetPath $arguments"
Write-Host ""
Write-Host "You can now double-click the icon on your Desktop to launch the app."
