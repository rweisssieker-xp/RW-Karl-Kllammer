# Epics: shipping, installer, and local-first stack

Aligned with root [`README.md`](../README.md) (*Known Limitations*, *Local-First Direction*).

## Epic A — Shipping and distribution

**Goal:** A repeatable way to install or update Carolus Nexus on Windows without cloning the repo.

- Define release artifact(s): e.g. self-contained `win-x64` publish folder, or MSIX / Inno Setup / WiX installer (pick one and document). Baseline: [`docs/RELEASE.md`](RELEASE.md) and [`avalonia/Publish-Release.cmd`](../avalonia/Publish-Release.cmd).
- Versioning: align `AssemblyInformationalVersion` / file version with git tags.
- First-run story: link to `windows/.env.example`, `windows/data/` layout, and optional portable vs installed paths.
- CI (optional): build + publish on tag.

**Exit criteria:** A maintainer can produce a release from documented commands; an end user can run the app without building from source.

## Epic B — Local vision + chat

**Goal:** Reduce dependence on cloud APIs for the main ask flow.

- Provider abstraction already leans toward OpenAI-compatible endpoints; extend docs and UI for a single “local base URL + model” profile (e.g. Ollama, LM Studio).
- Vision: define how screenshots are encoded and which local models are supported; gate features when the model is text-only.
- Performance: document VRAM/RAM expectations and image size limits.

**Exit criteria:** Ask flow works end-to-end with a documented local OpenAI-compatible vision model; failures are clear in UI.

## Epic C — Local TTS

**Goal:** Spoken responses without ElevenLabs.

- Options: Windows SAPI (parity with legacy fallback), Piper, Coqui, or OS-native engines.
- Integrate with existing `SpeakResponses` / synthesis button; keep ElevenLabs as optional cloud tier.

**Exit criteria:** With cloud keys cleared, the operator can still hear responses via the chosen local engine.

## Epic D — Observability

**Goal:** Easier support without reading raw JSON logs.

- Export bundle: settings snapshot (redacted), last N diagnostics lines, last ritual run log — from Diagnostics tab.

**Exit criteria:** One-click or single-folder export for support.
