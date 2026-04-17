@echo off
setlocal
where dotnet >nul 2>nul
if errorlevel 1 (
  echo dotnet SDK not found.
  exit /b 1
)

rem Framework-dependent publish (smaller; requires .NET 10 runtime on machine)
dotnet publish "%~dp0ClippyRW.Avalonia.csproj" -c Release -r win-x64 --self-contained false /p:PublishTrimmed=false /nologo
if errorlevel 1 exit /b %errorlevel%

echo.
echo Published to: %~dp0bin\Release\net10.0-windows\win-x64\publish\
echo See docs\RELEASE.md for self-contained and versioning options.
exit /b 0
