@echo off
setlocal
where dotnet >nul 2>nul
if errorlevel 1 (
  echo dotnet SDK not found.
  exit /b 1
)

dotnet publish "%~dp0ClippyRW.Windows.csproj" -c Release -r win-arm64 --self-contained false /nologo /clp:NoSummary
if errorlevel 1 exit /b %errorlevel%

copy /Y "%~dp0bin\Release\net10.0-windows\win-arm64\publish\ClippyRW.exe" "%~dp0ClippyRW.exe" >nul

exit /b %errorlevel%
