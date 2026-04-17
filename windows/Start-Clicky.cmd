@echo off
setlocal

if not exist "%~dp0bin\Release\net10.0-windows\win-arm64\publish\ClippyRW.exe" (
  call "%~dp0Build-Clicky.cmd"
  if errorlevel 1 exit /b %errorlevel%
)

start "" "%~dp0bin\Release\net10.0-windows\win-arm64\publish\ClippyRW.exe"
exit /b 0
