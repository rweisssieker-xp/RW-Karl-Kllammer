# Carolus Nexus

This folder contains the active desktop UI for Carolus Nexus / Karl Klammer.

It is the primary app surface going forward. Right now it still consumes state from the existing `windows/` runtime folders while the remaining Windows-specific integrations are ported. Carolus Nexus is the product name; Karl Klammer is the companion persona and overlay alias.

Current shape:

- responsive container-based shell instead of absolute WinForms coordinates
- Windows-targeted Avalonia UI on `.NET 10`
- reads current repo state from the existing Windows app folders
- surfaces provider, mode, knowledge, ritual, history, diagnostics, local agent-console status, the restored assistant runtime, and a cursor-following companion overlay

## Build

```cmd
cd avalonia
Build-Avalonia.cmd
```

## Run

```cmd
cd avalonia
Start-Avalonia.cmd
```

## Current Scope

- responsive operator UI with tabs for ask, dashboard, setup, knowledge, rituals, history, diagnostics, console, and live context
- compact vs wide layout mode
- probes the existing repo for:
  - `windows/.env`
  - `windows/data/settings.json`
  - local knowledge docs
  - recipe count
  - ritual/watch memory count
  - action history count
  - diagnostics export count
- can import/delete/reindex local knowledge files
- can filter, edit, clone, archive, dry-run and execute saved rituals
- includes a structured ritual engine with:
  - ritual metadata such as source type, risk, tags, guards, knowledge sources, parameters and steps
  - operator-memory metadata such as confidence score, last app / form context and estimated saved minutes
  - ritual step editing for `app|...` and `ax.*` actions with wait / retry / ifapp / if_form / if_dialog / if_tab / on_fail
  - ritual parameter editing for reusable placeholders like `{{customer_account}}`
  - ritual runtime with `run ritual`, `run next step` and `dry run`
  - ritual stats and discovery filters for category, source and risk
  - history -> ritual, watch -> ritual and knowledge -> ritual promotion flows
  - first teach-mode draft flow that turns newly captured semantic actions into a ritual draft
- can filter history and diagnostics
- can run the screenshot-aware assistant flow against Anthropic, OpenAI, or an OpenAI-compatible endpoint
- can include current Windows screenshots and local knowledge chunks in the ask flow
- can run provider smoke tests from the UI
- can record live microphone input from the Avalonia app, transcribe it, and ask automatically on release
- can transcribe imported audio files via ElevenLabs STT or local Whisper
- can synthesize assistant responses to an MP3 via ElevenLabs and open the generated file
- can launch local Codex, Claude Code, and OpenClaw runs and write their output to `codex output/`
- includes a Windows tray icon and a global push-to-talk hotkey based on `PUSH_TO_TALK_KEY`
- includes a cursor-following overlay companion window with state transitions for ready, listening, transcribing, thinking, speaking, and error
- surfaces proactive Karl suggestions, operator metrics and trust-first execution hints in the main UI
- includes a live desktop inspector in the `Live Context` tab for listing controls, reading form/dialog summaries, focusing the active window, and trying simple control actions
- includes a first AX 2012 / Microsoft Dynamics AX fat-client adapter with:
  - active AX window detection
  - AX context snapshotting for forms, dialogs, tabs, fields, actions, and grid candidates
  - `ax.*` plan actions in the Ask flow such as `ax.read_context`, `ax.read_form`, `ax.read_field`, `ax.set_field`, `ax.click_action`, `ax.open_tab`, `ax.open_lookup`, `ax.read_grid`, `ax.select_grid_row`, `ax.confirm_dialog`, `ax.cancel_dialog`, `ax.wait_for_form`, `ax.wait_for_dialog`, and `ax.wait_for_text`
  - extra plan guards `if_form`, `if_dialog`, `if_tab`, and `on_fail=skip`
  - AX-aware inspector shortcuts in the `Live Context` tab
  - AX ritual metadata via category, risk, source, app / form / dialog / tab guards
  - persistent AX action history entries with execution results
  - direct `save plan as ritual` flow from the Ask tab
  - RAG/SOP-aware prompting that can turn local AX work instructions into cautious `ax.*` plans
  - AX-first ritual execution through the dedicated ritual runtime instead of raw mouse replay
  - risk-aware blocking for high-risk ritual full runs and sensitive send/post/book actions

## Still Not Ported

- deeper AX-specific UI Automation / MSAA-backed control discovery beyond the current Win32-first hybrid heuristics
- robust AX grid row selection, lookup handling, and end-to-end posting workflows across all AX surfaces
- higher-confidence teach-once semantic AX recorder beyond the current history-backed teach-mode draft flow

Those still live in the legacy Windows runtime for now. The intended migration path is:

1. keep the retained Windows runtime logic as a temporary integration source
2. extract shared engine/services into a core library
3. keep Avalonia as the primary shell
4. move Windows-only automation into a platform adapter layer
