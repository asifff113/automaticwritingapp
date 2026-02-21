Param(
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

python -m pip install --upgrade pip
python -m pip install -r requirements-build.txt
python -m pip install -r requirements.txt

if ($Clean) {
    python .\build_release.py --clean
}
else {
    python .\build_release.py
}

Write-Host ""
Write-Host "Build complete."
Write-Host "EXE: dist\AutomaticWritingAssistant.exe"
Write-Host "ZIP: dist\AutomaticWritingAssistant-windows.zip"
