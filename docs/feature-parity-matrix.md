# Feature parity: Legacy WinForms vs Avalonia

Reference sources: [`windows/Clicky.Windows.cs`](../windows/Clicky.Windows.cs), [`avalonia/ViewModels/MainWindowViewModel.cs`](../avalonia/ViewModels/MainWindowViewModel.cs), [`CarolusNexus.Core/AgentHandoffTriggers.cs`](../CarolusNexus.Core/AgentHandoffTriggers.cs).

| Area | Legacy (WinForms) | Avalonia | Parity |
|------|-------------------|----------|--------|
| Ask + vision (Anthropic / OpenAI / compatible) | Yes | Yes | Full |
| Local knowledge RAG | Yes | Yes | Full |
| Rituals / recipes / watch / history / diagnostics | Yes (separate dialogs) | Yes (tabs) + support ZIP | Full (UX differs) |
| Setup wizard | Dedicated dialog | Setup tab + toggles (auto-route, auto-speak) | Functional equivalent |
| Codex / Claude Code / OpenClaw in Ask | Trigger + `IntentRouter` | [`RunAssistantAsync`](avalonia/ViewModels/MainWindowViewModel.cs) + [`LocalAgentRunService`](avalonia/Services/LocalAgentRunService.cs) + optional `-i` screens | Full |
| Codex / CLI from Console tab | Yes | Console tab | Full |
| Auto-route IDE/coding → Codex/OpenClaw | `IntentRouter` | `AgentHandoffTriggers.DetectIntentRoute` (toggle `AutoRouteLocalAgents`) | Full |
| TTS primary | ElevenLabs → WMP COM | ElevenLabs → MP3 + shell open | Partial (playback path differs) |
| TTS fallback | `SAPI.SpVoice` | [`WindowsSpeechFallback`](CarolusNexus.Platform.Windows/WindowsSpeechFallback.cs) | Full |
| Speak after cloud reply | Companion + audio path | Optional `SpeakAfterAsk` + `SpeakResponses` | Partial (UX differs) |
| Push-to-talk / tray | Yes | Yes | Full |
| Companion overlay + point tags | Yes | Yes | Full |
| Window inspector / AX | Rich legacy | Desktop inspector + AX v1 (see *Still Not Ported* in [`avalonia/README.md`](../avalonia/README.md)) | Partial |
| Foreground **app kind** (`ifapp`, ritual `GuardApp`) | `ActiveWindowService` → [`AppKindDetector`](../CarolusNexus.Core/AppKindDetector.cs) (via Core reference) | [`WindowsForegroundWindow`](../CarolusNexus.Platform.Windows/WindowsForegroundWindow.cs) → same detector | Full |

## Code anchors

- **Legacy** ask pipeline: `RunAskFlowAsync` in `Clicky.Windows.cs` (CLI triggers, then `IntentRouter.DetectRoute`, then cloud ask).
- **Avalonia** ask pipeline: `MainWindowViewModel.RunAssistantAsync` — order: OpenClaw / Claude Code / Codex explicit triggers → optional auto-route (`AutoRouteLocalAgents`) → `_assistantRuntimeService.AskAsync`.
- **Handoff logic**: [`AgentHandoffTriggers`](CarolusNexus.Core/AgentHandoffTriggers.cs) (shared core).
- **App kind** (browser, `ax`, `creo`, `babtec`, `catia`, `nx`, …): [`AppKindDetector.FromProcessName`](../CarolusNexus.Core/AppKindDetector.cs); legacy sets `AppKind` in `ActiveWindowService.GetActiveWindowInfo`; Avalonia uses [`WindowsForegroundWindow`](../CarolusNexus.Platform.Windows/WindowsForegroundWindow.cs).
- **TTS**: `MainWindowViewModel.SynthesizeCurrentResponseAsync` and `SpeakAfterCloudAskAsync`; ElevenLabs in [`AssistantRuntimeService`](avalonia/Services/AssistantRuntimeService.cs).

## Resolved follow-ups (historical)

Earlier drafts asked to port `IntentRouter`, CLI checks, and SAPI fallback into Avalonia — these are **done** as above.

## Remaining gaps (high level)

- AX depth (grid/lookup/posting reliability, UIA/MSAA): see [`backlog-ax.md`](backlog-ax.md) and *Still Not Ported* in [`avalonia/README.md`](../avalonia/README.md).
- Installer / cross-platform / local vision LLM: root [`README.md`](../README.md) and [`epics-shipping-local-first.md`](epics-shipping-local-first.md).
