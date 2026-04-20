$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$pythonExe = Join-Path $projectRoot "tools\python311-embed\python.exe"
$appScript = Join-Path $projectRoot "desktop_app\sermon_studio.py"
$distDir = Join-Path $projectRoot "dist"
$packageDir = Join-Path $distDir "FGMC-Youtube-Automation-package"
$toolsDir = Join-Path $packageDir "tools"
$builtExe = Join-Path $distDir "FGMC-Youtube-Automation.exe"
$packageExe = Join-Path $packageDir "FGMC-Youtube-Automation.exe"
$ffmpegExe = "C:\Users\Jimmy-Gram\AppData\Local\Microsoft\WinGet\Packages\Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe\ffmpeg-8.1-full_build\bin\ffmpeg.exe"
$ffprobeExe = "C:\Users\Jimmy-Gram\AppData\Local\Microsoft\WinGet\Packages\Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe\ffmpeg-8.1-full_build\bin\ffprobe.exe"
$ytDlpExe = Join-Path $projectRoot "tools\python311-embed\Scripts\yt-dlp.exe"

& $pythonExe -m PyInstaller `
  --noconfirm `
  --onefile `
  --windowed `
  --name "FGMC-Youtube-Automation" `
  $appScript

New-Item -ItemType Directory -Force -Path $toolsDir | Out-Null
Copy-Item $builtExe $packageExe -Force

if ((Test-Path $ffmpegExe) -and -not (Test-Path (Join-Path $toolsDir "ffmpeg.exe"))) {
  Copy-Item $ffmpegExe (Join-Path $toolsDir "ffmpeg.exe") -Force
}
if ((Test-Path $ffprobeExe) -and -not (Test-Path (Join-Path $toolsDir "ffprobe.exe"))) {
  Copy-Item $ffprobeExe (Join-Path $toolsDir "ffprobe.exe") -Force
}
Copy-Item $ytDlpExe (Join-Path $toolsDir "yt-dlp.exe") -Force

Remove-Item $builtExe -Force -ErrorAction SilentlyContinue
