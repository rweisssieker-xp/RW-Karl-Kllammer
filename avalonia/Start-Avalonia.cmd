@echo off
setlocal
call "%~dp0Build-Avalonia.cmd"
if errorlevel 1 exit /b %errorlevel%

dotnet "%~dp0bin\Release\net10.0-windows\CarolusNexus.dll"
