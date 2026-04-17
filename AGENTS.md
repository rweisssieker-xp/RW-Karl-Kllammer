# Carolus Nexus - Agent Instructions

`CLAUDE.md` points to this file.

## Overview

This repository should now be treated as an Avalonia desktop app project first.

The primary UI lives in `avalonia/` and:

- targets `.NET 10`
- is the only UI surface that should receive ongoing feature and layout work
- provides the responsive, cross-platform desktop shell
- reads current state from the existing `windows/` configuration and data folders
- is the migration target away from the old absolute-positioned WinForms UI

The legacy Windows runtime in `windows/` still contains platform-specific integrations and data formats. It currently:

- reads API secrets from `windows/.env`
- shows an always-on cursor companion overlay with state-based visuals
- can move the companion to detected on-screen targets from Claude point tags
- records microphone audio and transcribes it with ElevenLabs speech-to-text or a local Whisper Python install
- calls Anthropic, OpenAI, or OpenAI-compatible APIs directly for screenshot + vision chat
- calls ElevenLabs directly for TTS
- can route one-shot requests to a local Codex CLI run when the prompt contains `nimm codex`
- can route one-shot requests to a local Claude Code CLI run when the prompt contains `nimm claude code`
- can route one-shot requests to a local OpenClaw CLI run when the prompt contains `nimm openclaw`
- can attach screenshots to Codex runs for prompts like `nimm codex mit screen`
- stores reusable prompt recipes in `windows/data/automation-recipes.json`
- uses `playground/` in the repo root as the default Codex working directory
- writes Codex run logs to `codex output/` in the repo root
- writes Claude Code and OpenClaw run logs to `codex output/` in the repo root
- stores non-secret local settings in `windows/data/settings.json`
- uses a tray icon, a hold-to-talk button, and a global push-to-talk hotkey

## Key Files

| File | Purpose |
|------|---------|
| `avalonia/ClippyRW.Avalonia.csproj` | Primary Avalonia desktop app project. |
| `avalonia/Views/MainWindow.axaml` | Main responsive shell UI. |
| `avalonia/ViewModels/MainWindowViewModel.cs` | Main shell view model. |
| `avalonia/Services/OperatorWorkspaceService.cs` | Reads and manages existing repo/runtime state for the Avalonia app. |
| `avalonia/README.md` | Avalonia app usage and migration notes. |
| `windows/Clicky.Windows.cs` | Legacy Windows runtime source retained for platform-specific integrations that are not yet ported. |
| `windows/.env.example` | Template for local API secrets. |
| `windows/README.md` | Legacy Windows runtime notes and data/integration references. |
| `SOUL.md` | Editable personality layer for Karl Klammer, the companion persona of Carolus Nexus. |
| `NOTICE.md` | Short provenance note for the restarted Karl-Klammer repository. |
| `.gitignore` | Ignores local secrets, generated state, and Windows build artifacts. |

## Build & Run

```cmd
cd avalonia
Build-Avalonia.cmd
Start-Avalonia.cmd
```

## Conventions

- Treat `avalonia/` as the active app.
- Do not add new UI work to the WinForms surface in `windows/` unless the user explicitly asks for legacy maintenance.
- Keep the repo focused on the desktop app family only: the active Avalonia app plus the legacy Windows runtime/integration layer.
- Do not reintroduce macOS, Xcode, Worker, or Cloudflare-specific code unless the user explicitly asks for it.
- Keep secrets out of source files. Use `windows/.env`.
- Keep generated local state in `windows/data/`.
- Keep `playground/` out of git.
- Keep `codex output/` out of git.
- Keep `windows/ClippyRW.exe` and other generated local artifacts out of git.

## Self-Update

When the Avalonia app structure or the legacy Windows integration layer changes materially, update this file.
