# Karl Klammer For Windows — Build & Setup

Build, configure, and run the native Windows client. For a feature overview, see the [root README](../README.md).

## Requirements

- Windows
- one assistant provider key: Anthropic or OpenAI
- optional: an OpenAI-compatible base URL if you want to target other LLMs through a compatible endpoint
- optional: an ElevenLabs API key and voice ID
- optional: local Whisper (Python) if you want `STT_PROVIDER=whisper`
- optional: local Codex CLI for `nimm codex ...`
- optional: local Claude Code CLI for `nimm claude code ...`
- optional: local OpenClaw CLI for `nimm openclaw ...`

## Build

Double-click `Build-Clicky.cmd`, or from `cmd.exe`:

```cmd
path\to\clicky\windows\Build-Clicky.cmd
```

This publishes the WinForms app from `ClippyRW.Windows.csproj` with the installed .NET 10 SDK and also copies `ClippyRW.exe` into this folder for convenience.

## First Run

1. Copy `.env.example` to `.env` in this folder.
2. Fill in the required keys (see below).
3. Run `Start-Clicky.cmd` (or `ClippyRW.exe` directly).
4. Click `reload .env`, then `test apis`.
5. Optionally drop local docs into `windows/data/knowledge/` and click `reindex knowledge`.
6. Choose a provider, model, mode, and optional saved recipe in the app UI.
7. Either type a prompt and click `ask about my screen`, or hold the `hold to talk` button / push-to-talk hotkey (default `F8`).
8. To hand off a one-shot task, start the prompt with `nimm codex ...`, `nimm codex mit screen ...`, `nimm claude code ...`, or `nimm openclaw ...`.
9. For guided setup, open the new `setup wizard` from the main UI.

Without Codex/Claude Code/OpenClaw installed, the normal assistant still works — only the matching handoff flow is unavailable. Without local Whisper, set `STT_PROVIDER=elevenlabs`.

## `.env` Variables

### Assistant Providers

- `ANTHROPIC_API_KEY` — direct Anthropic screenshot + vision flow
- `OPENAI_API_KEY` — OpenAI or OpenAI-compatible screenshot + vision flow
- `OPENAI_BASE_URL` — optional OpenAI-compatible API base URL; leave at the default for OpenAI itself

### Speech

- `ELEVENLABS_API_KEY` — optional for `STT_PROVIDER=elevenlabs` and cloud TTS
- `ELEVENLABS_VOICE_ID` — optional voice id for ElevenLabs TTS

### Optional

| Variable | Default | Purpose |
|---|---|---|
| `STT_PROVIDER` | `whisper` | `elevenlabs` or `whisper` |
| `CODEX_COMMAND` | `codex.cmd` | Path or name of local Codex CLI |
| `CLAUDE_CODE_COMMAND` | `claude` | Path or name of local Claude Code CLI |
| `CODEX_WORKDIR` | `playground/` (repo root) | Working dir for Codex runs |
| `CODEX_TIMEOUT_SECONDS` | `900` | Timeout for Codex runs |
| `OPENCLAW_COMMAND` | `openclaw` | Path or name of local OpenClaw CLI |
| `OPENCLAW_SESSION_KEY` | `main` | Agent id / session key for OpenClaw |
| `OPENCLAW_TIMEOUT_SECONDS` | `120` | Timeout for OpenClaw runs |
| `WHISPER_PYTHON` | `python` | Python command for local Whisper |
| `WHISPER_MODEL` | `base` | Whisper model name |
| `WHISPER_LANGUAGE` | `de` | Speech language hint |
| `PUSH_TO_TALK_KEY` | `F8` | Global push-to-talk hotkey |

## Notes

- Secrets live in `windows/.env` next to the executable — never commit this file
- Local settings are stored in `windows/data/settings.json`
- Local knowledge files can be placed in `windows/data/knowledge/`
- The local knowledge index is stored in `windows/data/knowledge-index.json`
- Saved prompt recipes are stored in `windows/data/automation-recipes.json`
- Watch sessions are stored in `windows/data/watch-sessions.json`
- Confirmed desktop actions are logged in `windows/data/action-history.json`
- Repeated watch sessions and confirmed actions can surface as `learned rituals` in the main UI so you can save them as recipes
- Learned rituals now carry active-app context so repeated patterns can stay tied to the window/app they came from
- Some learned rituals can be replayed directly as confirmed action chains with the `replay` button
- Karl now shows a proactive `next idea` based on the active app context and learned rituals
- `use context idea` loads Karl's proactive suggestion into the prompt, while `run context idea` can execute a replayable context action chain
- Codex writes generated files to `playground/` unless `CODEX_WORKDIR` is set
- Run logs: `codex output/karl-klammer-codex-*.txt`, `karl-klammer-claude-code-*.txt`, `karl-klammer-openclaw-*.txt`
- If the global push-to-talk key cannot be registered, the on-screen hold button still works
- If direct ElevenLabs playback fails, the app falls back to local Windows speech when possible
- The first Whisper run can take longer if the selected model still needs to download
- `mode` changes Karl Klammer's system instructions: `companion`, `agent`, `automation`, or `watch`
- `use local knowledge` enables a simple local retrieval layer over files in `windows/data/knowledge/`
- The local knowledge manager supports `.txt`, `.md`, `.log`, `.json`, `.csv`, `.pdf`, and `.docx`
- `reindex knowledge` rebuilds the local chunk index before future asks
- The main UI now includes a small knowledge manager with filename/content search, preview, import, refresh, per-doc reindex, and remove
- Retrieved local sources are shown in the response area as a quick source summary
- The main UI now also exposes operator dialogs for `ritual manager`, `action history`, `window inspector`, `provider + voice`, `diagnostics`, and `setup wizard`
- `ritual manager` lets you inspect, edit, rename, delete, and run saved rituals/recipes in a dedicated editor
- `action history` now supports search, prompt reload, and replay of confirmed past actions
- `window inspector` now supports free-form semantic app actions in addition to `list_controls`, `read_form`, `read_table`, `read_dialog`, and `read_selected_row`
- `provider + voice` can reload `.env`, run a smoke test, and test TTS directly from the dialog
- `diagnostics` now supports filtering, copying all events, and exporting a diagnostics snapshot to `windows/data/diagnostics-*.log`
- `setup wizard` gives a small first-run health summary for provider keys, voice setup, knowledge usage, and mode selection
- If `save automation hints` is enabled, Karl Klammer can append reusable recipe suggestions which are stored automatically
- In `watch` mode, Karl Klammer logs prompt/response/screen summaries so repeated workflows can turn into future recipes
- Karl Klammer can now suggest `[ACTION:move]` or `[ACTION:click]`; the app always asks for confirmation before executing them
- Karl Klammer can also suggest confirmed `[ACTION:open|target]` and `[ACTION:type|text]` actions
- Karl Klammer can also suggest `[ACTION:hotkey|shortcut]` for confirmed keyboard shortcuts
- Karl Klammer can also suggest semantic app actions like `[ACTION:app|focus_address]`, which map differently depending on the active app
- For Windows fat-client apps, Karl Klammer can also use semantic control actions like `[ACTION:app|focus_window]`, `[ACTION:app|focus_control:Save]`, `[ACTION:app|click_control:OK]`, or `[ACTION:app|type_control:Customer=Alice]`
- Fat-client adapters now expose higher-level actions like `[ACTION:app|save]`, `[ACTION:app|confirm]`, `[ACTION:app|cancel]`, `[ACTION:app|next_field]`, and `[ACTION:app|next_tab]` for WinForms, WPF, Java, Qt, and classic desktop windows
- For WinForms/WPF-style desktop apps, Karl Klammer also tries Windows Accessibility/MSAA before falling back to raw child-window control matching
- Karl Klammer can also inspect desktop clients with `[ACTION:app|list_controls]`, `[ACTION:app|read_control:FieldName]`, and `[ACTION:app|activate_tab:TabName]`
- Karl Klammer can also return rough structured summaries with `[ACTION:app|read_form]`, `[ACTION:app|read_table]`, `[ACTION:app|read_dialog]`, and `[ACTION:app|read_selected_row]`
- Action confirmations now include a lightweight risk level so typing, clicking, and app actions in sensitive apps are easier to judge
- The main window shows supported semantic app actions for the currently active app kind
- Current app adapters cover browser, explorer, ide, mail, messenger, and a generic fallback set
- In IDE-heavy or obviously code-oriented prompts, Karl Klammer can auto-route to Codex without requiring the explicit trigger phrase
- Karl Klammer can suggest short confirmed action chains with `[ACTIONS:...]`, for example move + click or open + hotkey
- Action chains now understand lightweight workflow directives like `wait=500` and `ifapp=chrome` before the actual action token
- `use ritual` loads a learned pattern back into the prompt box so you can run or refine it immediately
- `replay` runs a learned ritual's stored action chain after confirmation
- The main window now shows the currently active app/window so you can see which context Karl is learning from
- `ClippyRW.exe` is a local build artifact and should stay out of git
- The published .NET 10 app lives under `windows/bin/Release/net10.0-windows/win-arm64/publish/`
- This is still a Windows app; moving to real cross-platform would require replacing WinForms and Windows API calls with a cross-platform UI/runtime layer
