@echo off
setlocal EnableExtensions DisableDelayedExpansion
set "NpuWinPs=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "NpuWinPs=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
"%NpuWinPs%" -NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File "%~dp0NPU-KMS-Monitor.ps1" -Mode Menu
exit /b %errorlevel%
