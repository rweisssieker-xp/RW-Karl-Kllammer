@echo off
setlocal
where dotnet >nul 2>nul
if errorlevel 1 (
  echo dotnet SDK not found.
  exit /b 1
)

dotnet build "%~dp0ClippyRW.Avalonia.csproj" -c Release /nologo /clp:NoSummary
exit /b %errorlevel%
