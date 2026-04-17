# Carolus Nexus

An Avalonia desktop assistant shell on `.NET 10`, backed by the existing Windows runtime data and integrations.

Carolus Nexus is being moved to a responsive Avalonia UI while retaining the existing local data, knowledge, ritual, diagnostics, and Windows automation foundations from the legacy runtime. Karl Klammer remains the companion persona and on-screen operator alias.

## Current App Focus

- `avalonia/` is now the active UI app
- responsive, container-based layout on `.NET 10`
- reads repo/runtime state from `windows/.env` and `windows/data/`
- now includes the main ask/runtime flow, screen-aware provider calls, live push-to-talk recording, imported-audio transcription, tray/hotkey routing, response speech generation, and a cursor-following companion overlay
- surfaces provider, model, knowledge, rituals, history, diagnostics, local agent status, proactive ritual suggestions, and operator metrics
- intended to remain the primary desktop experience

## Quick Start

1. Clone the repository.
2. Open the [`avalonia`](avalonia) folder.
3. Follow [`avalonia/README.md`](avalonia/README.md) for build and setup.

## Legacy Windows Runtime

The old WinForms runtime still exists in [`windows`](windows) and currently holds the Windows-specific integration layer.

- It still contains the deepest Windows-only automation and overlay code.
- Its data/config layout remains the source of truth used by the Avalonia app today.
- It should be treated as a legacy runtime/integration source, not the main UI target.

See [`windows/README.md`](windows/README.md) only when working on those retained platform-specific pieces.

## Local-First Direction

Carolus Nexus is designed so it can evolve toward a fully local setup.

Already local today:

- speech-to-text via local Whisper
- Codex, Claude Code, and OpenClaw one-shot handoffs

Still cloud-backed:

- the main screenshot-aware assistant flow (Anthropic / OpenAI / compatible APIs)
- TTS playback (ElevenLabs)

Also local in the current desktop workflow and inherited data layer:

- local knowledge retrieval over files in `windows/data/knowledge/`
- local ritual memory, watch logs, diagnostics logs, and action history
- Windows fat-client inspection and semantic app actions via desktop control/accessibility paths
- ritual confidence, saved-minutes heuristics, and trust-first execution policies across operator flows

The intended direction is a fully local stack: local Whisper + a local vision-capable chat model (e.g. Ollama) + a local TTS engine.

## Project Layout

```text
avalonia/
  ClippyRW.Avalonia.csproj
  Build-Avalonia.cmd
  Start-Avalonia.cmd
  README.md
windows/
  Clicky.Windows.cs
  Build-Clicky.cmd
  Start-Clicky.cmd
  .env.example
  README.md
SOUL.md
```

`SOUL.md` is an optional personality file. If present, Karl Klammer, the companion persona of Carolus Nexus, loads it and injects it into the Anthropic system prompt.

## Known Limitations

- The active UI is Avalonia, but several deep integrations still depend on Windows-specific code and data in `windows/`
- Avalonia is now Windows-targeted in this repo because the remaining runtime and automation layers are Windows-specific
- no installer yet
- deep fat-client control execution is only partially ported: Avalonia now has a desktop inspector and simple control actions, but not the full legacy Windows automation stack
- Codex, Claude Code, and OpenClaw handoffs are one-shot background runs, not persistent sessions
- handoff triggers are tuned for German speech variants (e.g. `nimm codex`, `nimm claude code`, `nimm openclaw`)
- speech and vision features depend on external API availability

## Credits

The core idea of an **always-on companion sitting next to the mouse cursor** is inspired by [farzaa/clicky](https://github.com/farzaa/clicky), a macOS/Swift menu-bar assistant by Farza Majeed (MIT). Thanks for that spark.

Carolus Nexus started as a larger C# / WinForms rewrite with orchestration of three local CLI agents (Codex, Claude Code, OpenClaw), direct Anthropic and ElevenLabs integrations, local Whisper STT, and a tray + hotkey workflow for Windows. Karl Klammer remains the companion identity inside the product. The repo originally started from a local clone of the upstream, so small remnants (folder names, minor snippets) may still trace back to it.

See [`NOTICE.md`](NOTICE.md) for details on origin and attribution.

## License

MIT. See [`LICENSE`](LICENSE) and [`NOTICE.md`](NOTICE.md) for licensing and provenance notes.
